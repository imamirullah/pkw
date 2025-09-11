const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { MongoClient, ObjectId } = require("mongodb");
const cors = require("cors");

const app = express();
const port = 3000;

// ✅ Enable CORS
app.use(cors());
app.use(express.json());

// 🔗 MongoDB Atlas connection
const uri =
  "mongodb+srv://skillsha00:Amirkhan%401212%2312@cluster0.wucdenw.mongodb.net/excelupload?retryWrites=true&w=majority";

const client = new MongoClient(uri);
let db, collection;

async function connectDB() {
  try {
    await client.connect();
    db = client.db("excelupload");
    collection = db.collection("employees");
    console.log("✅ MongoDB connected");
  } catch (err) {
    console.error("❌ MongoDB connection failed:", err);
  }
}
connectDB();

// ✅ Helper function
function cleanString(value, toUpper = false) {
  if (!value) return "";
  let str = value.toString().trim();
  return toUpper ? str.toUpperCase() : str;
}

// ✅ Multer config
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// ✅ Upload & save XLS data to MongoDB
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];

    // ✅ normalize headers
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      defval: "",
      raw: false,
      header: 1, // rows as arrays
    });

    const headers = data[0].map((h) =>
      h.toString().trim().toLowerCase().replace(/\s+/g, "")
    ); // normalize headers

    const rows = data.slice(1);

    const mappedData = rows.map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] ? row[i].toString().trim() : "";
      });

      return {
        name: cleanString(obj["name"], true), // always CAPITAL
        designation: cleanString(obj["designation"]),
        workingArea: cleanString(obj["workingarea"]),
        validUpto: cleanString(obj["validupto"]),
        codeNo: cleanString(obj["codeno"]),
        adhaarNo: cleanString(obj["adhaarno"]),
        adress: cleanString(obj["address"]),
      };
    });

    if (mappedData.length > 0) {
      await collection.insertMany(mappedData);
      res.json({
        message: "✅ Data uploaded successfully",
        inserted: mappedData.length,
      });
    } else {
      res.status(400).json({ message: "⚠️ No data found in file" });
    }
  } catch (error) {
    console.error("❌ Upload error:", error);
    res.status(500).json({ message: "❌ Error uploading file" });
  }
});

// ✅ Fetch all data
app.get("/employees", async (req, res) => {
  try {
    const employees = await collection.find({}).toArray();
    res.json(employees);
  } catch (error) {
    console.error(error);
    res.status(500).json({ message: "❌ Error fetching data" });
  }
});

// ✅ Search by name, codeNo, or adhaarNo
app.get("/search", async (req, res) => {
  try {
    const { name, codeNo, adhaarNo } = req.query;
    let query = {};

    if (name) query.name = { $regex: name, $options: "i" };
    if (codeNo) query.codeNo = codeNo;
    if (adhaarNo) query.adhaarNo = adhaarNo;

    const result = await collection.find(query).toArray();
    res.json(result);
  } catch (error) {
    console.error(error);
    res.status(500).json({ message: "❌ Error searching data" });
  }
});

// ✅ Delete by MongoDB _id
app.delete("/delete/:id", async (req, res) => {
  try {
    const id = req.params.id;
    await collection.deleteOne({ _id: new ObjectId(id) });
    res.json({ message: "✅ Record deleted" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ message: "❌ Error deleting data" });
  }
});

// ✅ Delete all data
app.delete("/delete-all", async (req, res) => {
  try {
    await collection.deleteMany({});
    res.json({ message: "✅ All records deleted" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ message: "❌ Error deleting all data" });
  }
});
// ✅ Bulk delete
app.post("/delete-many", async (req, res) => {
  try {
    const { ids } = req.body;
    if (!ids || ids.length === 0) {
      return res.status(400).json({ message: "No IDs provided" });
    }

    await db.collection("employees").deleteMany({
      _id: { $in: ids.map(id => new ObjectId(id)) }
    });

    res.json({ message: `${ids.length} records deleted successfully` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: "Error deleting records" });
  }
});


app.listen(port, () => {
  console.log(`🚀 Server running on http://localhost:${port}`);
});
