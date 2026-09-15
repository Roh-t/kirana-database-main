import express from 'express';
import mongoose from 'mongoose';
import cors from 'cors';
import dotenv from 'dotenv';
import multer from 'multer';
import * as XLSX from 'xlsx';
import { v2 as cloudinary } from 'cloudinary';

dotenv.config();

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

// ------------------------------------
// MONGODB SCHEMA (11 EXACT COLUMNS)
// ------------------------------------
const productSchema = new mongoose.Schema({
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
  'Indian Sub-Category': { type: String, default: '' }
}, { timestamps: true });

const MasterProduct = mongoose.model('MasterProduct', productSchema);

// ------------------------------------
// ☁️ CLOUDINARY IMAGE UPLOAD ROUTE
// ------------------------------------
app.post('/api/upload-image', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No image file provided' });
    }

    // Convert file buffer to Base64 Data URI
    const b64 = Buffer.from(req.file.buffer).toString('base64');
    const dataURI = `data:${req.file.mimetype};base64,${b64}`;

    // Upload to Cloudinary
    const result = await cloudinary.uploader.upload(dataURI, {
      folder: 'kirana_store',
      resource_type: 'image'
    });

    res.json({ url: result.secure_url });
  } catch (err) {
    console.error('Cloudinary upload error:', err);
    res.status(500).json({ error: err.message || 'Cloudinary upload failed' });
  }
});

// Master Count
app.get('/api/master/count', async (req, res) => {
  try {
    const count = await MasterProduct.countDocuments();
    res.json({ count });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Master Search
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
        { 'Indian Sub-Category': regex }
      ]
    }).limit(24);

    res.json(results);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Upload Master Excel File to MongoDB
app.post('/api/master/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (rawData.length === 0) return res.status(400).json({ error: 'File is empty' });

    const getVal = (row, possibleKeys) => {
      const rowKeys = Object.keys(row);
      for (const pKey of possibleKeys) {
        const matched = rowKeys.find(k => k.trim().toLowerCase() === pKey.toLowerCase());
        if (matched !== undefined && row[matched] !== undefined) return String(row[matched]).trim();
      }
      return '';
    };

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
      'Indian Sub-Category': getVal(row, ['Indian Sub-Category', 'IndianSubCategory'])
    }));

    await MasterProduct.deleteMany({});
    
    const CHUNK_SIZE = 2000;
    for (let i = 0; i < normalized.length; i += CHUNK_SIZE) {
      await MasterProduct.insertMany(normalized.slice(i, i + CHUNK_SIZE), { ordered: false });
    }

    const totalCount = await MasterProduct.countDocuments();
    res.json({ message: 'Success', count: totalCount });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 5000;
mongoose.connect(process.env.MONGODB_URI)
  .then(() => {
    console.log('Connected to MongoDB Atlas');
    app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
  })
  .catch(err => console.error('MongoDB Connection Failed:', err));