import express from 'express';
import mongoose from 'mongoose';
import cors from 'cors';
import dotenv from 'dotenv';
import multer from 'multer';
import * as XLSX from 'xlsx';
import { v2 as cloudinary } from 'cloudinary';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Configure Cloudinary
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME || 'mq6c2sfk',
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

const upload = multer({ limits: { fileSize: 50 * 1024 * 1024 } });

// Health check
app.get('/', (req, res) => {
  res.send('✅ Kirana Database API is running!');
});

// ------------------------------------
// MONGODB SCHEMA (11 EXACT COLUMNS)
// ------------------------------------
const productSchema = new mongoose.Schema(
  {
    Image: { type: String, default: '' },
    Name: { type: String, required: true, index: true },
    Price: { type: String, default: '' },
    'Original Price': { type: String, default: '' },
    Quantity: { type: String, default: '' },
    'Sub-Category': { type: String, default: '' },
    Category: { type: String, default: '' },
    'Hindi Name': { type: String, default: '', index: true },
    'Hinglish Name': { type: String, default: '', index: true },
    'Indian Category': { type: String, default: '' },
    'Indian Sub-Category': { type: String, default: '' },
  },
  { timestamps: true }
);

const MasterProduct = mongoose.model('MasterProduct', productSchema);
let databaseReady = Promise.resolve(false);

// Helper to extract values case-insensitively
const getVal = (row, possibleKeys) => {
  const rowKeys = Object.keys(row);
  for (const pKey of possibleKeys) {
    const matched = rowKeys.find(
      (k) => k.trim().toLowerCase() === pKey.toLowerCase()
    );
    if (matched !== undefined && row[matched] !== undefined) {
      return String(row[matched]).trim();
    }
  }
  return '';
};

// -----------------------------------------------------------
// 🚀 AUTOMATIC PERMANENT FILE SEEDER INTO MONGODB
// -----------------------------------------------------------
async function autoSeedPermanentFile() {
  try {
    const count = await MasterProduct.countDocuments();
    if (count > 0) {
      console.log(`✅ MongoDB already has ${count.toLocaleString()} items. No seeding required.`);
      return;
    }

    const possiblePaths = [
      path.join(__dirname, 'catalog.xlsx'),
      path.join(__dirname, 'catalog.XLSX'),
      path.join(__dirname, 'data.xlsx')
    ];

    let filePath = possiblePaths.find((p) => fs.existsSync(p));

    if (!filePath) {
      throw new Error('catalog.xlsx was not found in the backend folder.');
    }

    console.log(`⏳ Found permanent file: ${filePath}`);
    console.log('⏳ Auto-loading 33,000+ items directly from code into MongoDB...');

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (rawData.length === 0) {
      throw new Error('catalog.xlsx does not contain any rows.');
    }

    const normalized = rawData.map((row) => ({
      Image: getVal(row, ['Image', 'imageUrl', 'image', 'img']),
      Name: getVal(row, ['Name', 'name', 'productName', 'Title']),
      Price: getVal(row, ['Price', 'lingPr', 'sellingPrice', 'price', 'sp']),
      'Original Price': getVal(row, ['Original Price', 'OriginalPrice', 'mrp', 'MRP']),
      Quantity: getVal(row, ['Quantity', 'quantity', 'qty', 'weight', 'size']),
      'Sub-Category': getVal(row, ['Sub-Category', 'SubCategory', 'subcategory']),
      Category: getVal(row, ['Category', 'category']),
      'Hindi Name': getVal(row, ['Hindi Name', 'HindiName', 'Hindi']),
      'Hinglish Name': getVal(row, ['Hinglish Name', 'HinglishName', 'Hinglish']),
      'Indian Category': getVal(row, ['Indian Category', 'IndianCategory']),
      'Indian Sub-Category': getVal(row, ['Indian Sub-Category', 'IndianSubCategory']),
    }));

    // Bulk insert in chunks of 2,000 for high speed
    const CHUNK_SIZE = 2000;
    for (let i = 0; i < normalized.length; i += CHUNK_SIZE) {
      const chunk = normalized.slice(i, i + CHUNK_SIZE);
      await MasterProduct.insertMany(chunk, { ordered: false });
    }

    const totalCount = await MasterProduct.countDocuments();
    console.log(`🎉 SUCCESS: Auto-seeded ${totalCount.toLocaleString()} items permanently into MongoDB!`);
  } catch (err) {
    console.error('Auto-seed error:', err);
    throw err;
  }
}

// ------------------------------------
// ☁️ CLOUDINARY UPLOAD ROUTE
// ------------------------------------
app.post('/api/upload-image', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No image provided' });

    const b64 = Buffer.from(req.file.buffer).toString('base64');
    const dataURI = `data:${req.file.mimetype};base64,${b64}`;

    const result = await cloudinary.uploader.upload(dataURI, {
      folder: 'kirana_store',
      resource_type: 'image',
    });

    res.json({ url: result.secure_url });
  } catch (err) {
    res.status(500).json({ error: err.message || 'Upload failed' });
  }
});

// ------------------------------------
// MASTER ROUTES
// ------------------------------------
app.get('/api/master/count', async (req, res) => {
  try {
    const count = await MasterProduct.countDocuments();
    res.json({ count });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Return the catalog seeded from catalog.xlsx for the frontend search index.
app.get('/api/master/catalog', async (req, res) => {
  try {
    if (!(await databaseReady)) {
      return res.status(503).json({ error: 'Catalog database is not ready yet. Please retry shortly.' });
    }

    const catalog = await MasterProduct.find().select('-__v').lean();
    if (catalog.length === 0) {
      return res.status(503).json({ error: 'Catalog database is empty.' });
    }

    res.json(catalog);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/master/search', async (req, res) => {
  try {
    const q = req.query.q ? req.query.q.trim() : '';
    if (!q) {
      const sample = await MasterProduct.find().limit(24);
      return res.json(sample);
    }

    const regex = new RegExp(q, 'i');
    const results = await MasterProduct.find({
      $or: [
        { Name: regex },
        { 'Hindi Name': regex },
        { 'Hinglish Name': regex },
        { 'Sub-Category': regex },
        { 'Indian Sub-Category': regex },
      ],
    }).limit(24);

    res.json(results);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------
// ➕ IMPORT DATASET (ADDS MORE ITEMS TO EXISTING ONES)
// -----------------------------------------------------------
app.post('/api/master/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (rawData.length === 0) return res.status(400).json({ error: 'File is empty' });

    const normalized = rawData.map((row) => ({
      Image: getVal(row, ['Image', 'imageUrl', 'image', 'img']),
      Name: getVal(row, ['Name', 'name', 'productName', 'Title']),
      Price: getVal(row, ['Price', 'lingPr', 'sellingPrice', 'price', 'sp']),
      'Original Price': getVal(row, ['Original Price', 'OriginalPrice', 'mrp', 'MRP']),
      Quantity: getVal(row, ['Quantity', 'quantity', 'qty', 'weight', 'size']),
      'Sub-Category': getVal(row, ['Sub-Category', 'SubCategory', 'subcategory']),
      Category: getVal(row, ['Category', 'category']),
      'Hindi Name': getVal(row, ['Hindi Name', 'HindiName', 'Hindi']),
      'Hinglish Name': getVal(row, ['Hinglish Name', 'HinglishName', 'Hinglish']),
      'Indian Category': getVal(row, ['Indian Category', 'IndianCategory']),
      'Indian Sub-Category': getVal(row, ['Indian Sub-Category', 'IndianSubCategory']),
    }));

    // NOTE: Does NOT delete existing items! Appends new items to MongoDB
    const CHUNK_SIZE = 2000;
    for (let i = 0; i < normalized.length; i += CHUNK_SIZE) {
      const chunk = normalized.slice(i, i + CHUNK_SIZE);
      await MasterProduct.insertMany(chunk, { ordered: false });
    }

    const totalCount = await MasterProduct.countDocuments();
    res.json({ message: 'Success', count: totalCount, added: normalized.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ------------------------------------
// START SERVER & CONNECT MONGODB
// ------------------------------------
const PORT = process.env.PORT || 5000;

app.listen(PORT, '0.0.0.0', () => {
  console.log(`🚀 Server is listening on port ${PORT}`);
});

if (!process.env.MONGODB_URI) {
  console.error('⚠️ WARNING: MONGODB_URI is missing!');
} else {
  databaseReady = mongoose
    .connect(process.env.MONGODB_URI)
    .then(async () => {
      console.log('✅ Connected to MongoDB Atlas');
      // Automatically load permanent file into MongoDB if empty
      await autoSeedPermanentFile();
      return true;
    })
    .catch((err) => {
      console.error('❌ MongoDB initialization failed:', err.message);
      return false;
    });
}