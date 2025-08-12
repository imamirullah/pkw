// server.js
const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json()); // parse JSON bodies

// ----------------- MONGO -----------------
mongoose.connect('mongodb+srv://skillsha00:Amirkhan%401212%2312@cluster0.wucdenw.mongodb.net/excelupload?retryWrites=true&w=majority', {
  useNewUrlParser: true,
  useUnifiedTopology: true
}).then(() => console.log('✅ MongoDB connected'))
  .catch(err => console.error('❌ MongoDB connection error:', err.message));

// ----------------- SCHEMA -----------------
const userSchema = new mongoose.Schema({
  name: String,
  designation: String,
  workingArea: String,
  validUpto: Date,
  codeNo: String,
  adhaarNo: String
}, { timestamps: true });

const User = mongoose.model('User', userSchema);

// ----------------- MULTER (memory) -----------------
const upload = multer({ storage: multer.memoryStorage() });

// ----------------- HELPERS -----------------

// Normalize header: lowercase, remove punctuation and extra spaces
function normalizeHeader(h) {
  if (!h) return '';
  return h.toString().toLowerCase().replace(/[^\w\s]|_/g, '').replace(/\s+/g, ' ').trim();
}

// Map possible header names to canonical keys we use
function buildHeaderMap(headers) {
  // headers is array like ['Name','Designation',...]
  const map = {}; // normalized -> original
  for (const h of headers) {
    map[normalizeHeader(h)] = h;
  }
  return map;
}

// tries to read value from row by multiple possible header names
function getRowValue(row, headerMap, possibilities) {
  for (const p of possibilities) {
    const norm = normalizeHeader(p);
    if (headerMap[norm] && row.hasOwnProperty(headerMap[norm])) {
      return row[headerMap[norm]];
    }
  }
  // fallback: try direct keys (in case sheet_to_json already normalized)
  for (const key of Object.keys(row)) {
    if (normalizeHeader(key) === normalizeHeader(possibilities[0])) return row[key];
  }
  return undefined;
}

// Convert Excel date serial or string 'DD-MM-YYYY' / 'YYYY-MM-DD' to JS Date (or null)
function parseExcelDate(val) {
  if (val == null) return null;
  if (typeof val === 'number') {
    // Excel serial -> JS date
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(date.getTime()) ? null : date;
  }
  if (typeof val === 'string') {
    // Try common formats: DD-MM-YYYY, D-M-YYYY, YYYY-MM-DD
    const s = val.trim();
    // handle DD-MM-YYYY
    const dmy = /^(\d{1,2})-(\d{1,2})-(\d{4})$/;
    const ymd = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
    if (dmy.test(s)) {
      const [, dd, mm, yyyy] = s.match(dmy);
      const date = new Date(`${yyyy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}T00:00:00Z`);
      return isNaN(date.getTime()) ? null : date;
    }
    if (ymd.test(s)) {
      const [, yyyy, mm, dd] = s.match(ymd);
      const date = new Date(`${yyyy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}T00:00:00Z`);
      return isNaN(date.getTime()) ? null : date;
    }
    // last resort
    const parsed = new Date(s);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  // other types
  return null;
}

// ----------------- UPLOAD ENDPOINT -----------------
app.post('/upload', upload.single('excel'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ message: 'No file uploaded' });

    // Read workbook from buffer
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    // get rows as array of objects
    const sheetData = XLSX.utils.sheet_to_json(sheet, { defval: '' }); // defval to avoid undefined

    if (!sheetData.length) return res.json({ message: 'No rows found in Excel' });

    // Build header map from first row's keys
    const headers = Object.keys(sheetData[0]);
    const headerMap = buildHeaderMap(headers);

    // Acceptable variations for each field
    const NAME_KEYS = ['Name', 'name'];
    const DESIGNATION_KEYS = ['Designation', 'designation', 'job title'];
    const WORKING_AREA_KEYS = ['Working Area', 'working area', 'workingarea', 'area'];
    const VALID_KEYS = ['Valid Up-to', 'Valid Upto', 'Valid Upto', 'Valid U pto', 'valid upto', 'valid up to', 'validupto', 'valid'];
    const CODE_KEYS = ['Code No', 'Code No.', 'CodeNo', 'code no', 'codeno'];
    const ADHAAR_KEYS = ['Adhaar no', 'Aadhaar No', 'Aadhaar', 'adhaar', 'aadhaar no', 'aadhaar'];

    let insertedCount = 0;
    let skippedCount = 0;
    const skippedDetails = [];

    for (const row of sheetData) {
      // read capable of different header names
      const rawCode = getRowValue(row, headerMap, CODE_KEYS) || '';
      const rawAdhaar = getRowValue(row, headerMap, ADHAAR_KEYS) || '';

      const codeNo = rawCode ? String(rawCode).trim() : null;
      // normalize Aadhaar: remove spaces
      const adhaarNo = rawAdhaar ? String(rawAdhaar).replace(/\s+/g, '').trim() : null;

      if (!codeNo && !adhaarNo) {
        skippedCount++;
        skippedDetails.push({ reason: 'missing code & adhaar', row });
        continue;
      }

      // Check duplicates: if any existing with same codeNo (case-insensitive) OR same adhaarNo (exact)
      const dupQuery = [];
      if (codeNo) dupQuery.push({ codeNo: new RegExp(`^${escapeRegExp(codeNo)}$`, 'i') });
      if (adhaarNo) dupQuery.push({ adhaarNo: adhaarNo });

      const existing = dupQuery.length ? await User.findOne({ $or: dupQuery }) : null;
      if (existing) {
        skippedCount++;
        skippedDetails.push({ reason: 'duplicate', found: existing._id, row });
        continue;
      }

      // Build user object
      const userObj = {
        name: (getRowValue(row, headerMap, NAME_KEYS) || '').toString().trim(),
        designation: (getRowValue(row, headerMap, DESIGNATION_KEYS) || '').toString().trim(),
        workingArea: (getRowValue(row, headerMap, WORKING_AREA_KEYS) || '').toString().trim(),
        validUpto: parseExcelDate(getRowValue(row, headerMap, VALID_KEYS)),
        codeNo: codeNo || '',
        adhaarNo: adhaarNo || ''
      };

      // Save
      await User.create(userObj);
      insertedCount++;
    }

    return res.json({
      message: `Upload complete. Inserted: ${insertedCount}, Skipped: ${skippedCount}`,
      skippedDetails: skippedDetails.slice(0, 10) // return a few examples if needed
    });

  } catch (err) {
    console.error('Upload error:', err);
    return res.status(500).json({ message: 'Error processing file', error: err.message });
  }
});

// helper to escape RegExp
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ----------------- CRUD API -----------------




app.get('/search', async (req, res) => {
    const q = req.query.q;
    const regex = new RegExp(q, 'i'); // case-insensitive
    const users = await User.find({
        $or: [
            { codeNo: regex },
            { adhaarNo: regex }
        ]
    });
    res.json(users);
});

// get all users
app.get('/users', async (req, res) => {
  try {
    const users = await User.find().sort({ createdAt: -1 }).lean();
    res.json(users);
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Error fetching users' });
  }
});

// get by codeNo or adhaar: we keep route /user/:query for compatibility
app.get('/user/:query', async (req, res) => {
  try {
    const q = req.params.query.trim();
    const users = await User.find({
      $or: [
        { codeNo: new RegExp(`^${escapeRegExp(q)}$`, 'i') },
        { adhaarNo: q.replace(/\s+/g, '') }
      ]
    }).lean();
    if (!users.length) return res.status(404).json({ message: 'No user found' });
    res.json({ message: `Found ${users.length}`, data: users });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Error searching user' });
  }
});

// add single user (from admin panel)
app.post('/user', async (req, res) => {
  try {
    const { name, designation, workingArea, validUpto, codeNo, adhaarNo } = req.body;
    if (!codeNo && !adhaarNo) return res.status(400).json({ message: 'Code No or Aadhaar required' });

    // duplicate check
    const dup = await User.findOne({
      $or: [
        codeNo ? { codeNo: new RegExp(`^${escapeRegExp(codeNo)}$`, 'i') } : null,
        adhaarNo ? { adhaarNo: String(adhaarNo).replace(/\s+/g, '') } : null
      ].filter(Boolean)
    });

    if (dup) return res.status(409).json({ message: 'Duplicate code or aadhaar exists' });

    const user = new User({
      name: name || '',
      designation: designation || '',
      workingArea: workingArea || '',
      validUpto: validUpto ? parseExcelDate(validUpto) : null,
      codeNo: codeNo ? String(codeNo).trim() : '',
      adhaarNo: adhaarNo ? String(adhaarNo).replace(/\s+/g, '') : ''
    });
    await user.save();
    res.status(201).json({ message: 'User created', user });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Error creating user' });
  }
});

// update user
app.put('/user/:id', async (req, res) => {
  try {
    const id = req.params.id;
    const { name, designation, workingArea, validUpto, codeNo, adhaarNo } = req.body;

    // check duplicate against others
    const dupQuery = [];
    if (codeNo) dupQuery.push({ codeNo: new RegExp(`^${escapeRegExp(codeNo)}$`, 'i') });
    if (adhaarNo) dupQuery.push({ adhaarNo: String(adhaarNo).replace(/\s+/g, '') });

    if (dupQuery.length) {
      const dup = await User.findOne({ $and: [{ _id: { $ne: id } }, { $or: dupQuery }] });
      if (dup) return res.status(409).json({ message: 'Another user with same code/adhaar exists' });
    }

    const updated = await User.findByIdAndUpdate(id, {
      name: name || '',
      designation: designation || '',
      workingArea: workingArea || '',
      validUpto: validUpto ? parseExcelDate(validUpto) : null,
      codeNo: codeNo ? String(codeNo).trim() : '',
      adhaarNo: adhaarNo ? String(adhaarNo).replace(/\s+/g, '') : ''
    }, { new: true });

    if (!updated) return res.status(404).json({ message: 'User not found' });
    res.json({ message: 'User updated', user: updated });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Error updating user' });
  }
});

// delete single user
app.delete('/user/:id', async (req, res) => {
  try {
    const id = req.params.id;
    await User.findByIdAndDelete(id);
    res.json({ message: 'User deleted' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Error deleting user' });
  }
});

// bulk delete - expects { ids: [id1,id2,...] }
app.post('/users/bulk-delete', async (req, res) => {
  try {
    const { ids } = req.body;
    if (!Array.isArray(ids) || !ids.length) return res.status(400).json({ message: 'No ids provided' });
    await User.deleteMany({ _id: { $in: ids } });
    res.json({ message: `Deleted ${ids.length} users` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: 'Error bulk deleting' });
  }
});

// ----------------- START -----------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Server running on http://localhost:${PORT}`));
