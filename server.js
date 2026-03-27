require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const QRCode = require('qrcode');
const sharp = require("sharp");

const app = express();
app.use(express.json());
app.use(express.static('public'));
app.use('/images', express.static('images'));
app.use('/generated', express.static('generated'));

// MongoDB Connection
mongoose.connect(process.env.MONGODB_URI)
  .then(() => console.log('✅ MongoDB connected'))
  .catch(err => console.error('❌ MongoDB error:', err));

// Schema
const StudentSchema = new mongoose.Schema({
  name: { type: String, required: true },
  schoolName: { type: String, required: true },
  rollNo: String,
  class: String,
  email: String,
  phone: String,
  address: String,
  uploadedAt: { type: Date, default: Date.now },
  cardGenerated: { type: Boolean, default: false },
  cardPath: String
});

const Student = mongoose.model('Student', StudentSchema);

// Multer setup
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'uploads/'),
  filename: (req, file, cb) => cb(null, Date.now() + path.extname(file.originalname))
});
const upload = multer({ storage });

// Ensure directories exist
['uploads', 'generated', 'images'].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir);
});

// POST: Upload Excel
app.post('/api/upload', upload.single('excel'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const workbook = xlsx.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    if (!data.length) return res.status(400).json({ error: 'Excel file is empty' });

    // Normalize column names (case-insensitive)
    const normalized = data.map(row => {
      const obj = {};
      Object.keys(row).forEach(k => { obj[k.toLowerCase().trim()] = row[k]; });
      return {
        name: obj['name'] || obj['student name'] || obj['studentname'] || '',
        schoolName: obj['school'] || obj['school name'] || obj['schoolname'] || obj['institution'] || '',
        rollNo: obj['roll no'] || obj['rollno'] || obj['roll'] || obj['roll number'] || '',
        class: obj['class'] || obj['grade'] || obj['std'] || '',
        email: obj['email'] || obj['email id'] || '',
        phone: obj['phone'] || obj['mobile'] || obj['contact'] || '',
        address: obj['address'] || obj['city'] || ''
      };
    }).filter(s => s.name && s.schoolName);

    const inserted = await Student.insertMany(normalized);
    fs.unlinkSync(req.file.path);

    res.json({
      success: true,
      message: `${inserted.length} students uploaded successfully`,
      count: inserted.length,
      students: inserted
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// GET: All Students
app.get('/api/students', async (req, res) => {
  try {
    const students = await Student.find().sort({ uploadedAt: -1 });
    res.json(students);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET: Student by ID (for QR scan)
app.get('/api/student/:id', async (req, res) => {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ error: 'Student not found' });
    res.json(student);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

function escapeSvgText(value = '') {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// POST: Generate Card for a student
app.post('/api/generate-card/:id', async (req, res) => {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ error: 'Student not found' });

    // Find template image
    const imageFiles = fs.readdirSync('images').filter(f =>
      ['.jpg', '.jpeg', '.png'].includes(path.extname(f).toLowerCase())
    );
    if (!imageFiles.length) {
      return res.status(400).json({ error: 'No template image found in /images folder' });
    }

    const templatePath = path.join('images', imageFiles[0]);
    const templateSharp = sharp(templatePath);
    const meta = await templateSharp.metadata();

    const width = meta.width || 1200;
    const height = meta.height || 800;
    const cx = Math.floor(width / 2);
    const cy = Math.floor(height / 2);

    // Generate QR
    const baseUrl = process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`;
    const qrData = `${baseUrl}/student.html?id=${student._id}`;

    const qrSize = Math.floor(Math.min(width, height) * 0.22);
    const qrBuffer = await QRCode.toBuffer(qrData, {
      width: qrSize,
      margin: 1,
      color: { dark: '#000000', light: '#ffffff' }
    });

    // Text sizes
    const nameFontSize = Math.floor(width * 0.055);
    const schoolFontSize = Math.floor(width * 0.035);
    const labelFontSize = Math.floor(width * 0.022);

    const safeName = escapeSvgText((student.name || '').toUpperCase());
    const safeSchool = escapeSvgText(student.schoolName || '');

    // SVG text overlay (centered)
    const textSvg = Buffer.from(`
      <svg width="${width}" height="${height}">
        <style>
          .name { fill:#1a1a2e; font-size:${nameFontSize}px; font-weight:700; font-family:Arial, sans-serif; }
          .school { fill:#2d4a7a; font-size:${schoolFontSize}px; font-weight:500; font-family:Arial, sans-serif; }
          .label { fill:#555555; font-size:${labelFontSize}px; font-weight:400; font-family:Arial, sans-serif; }
        </style>

        <text x="${cx}" y="${Math.round(cy - qrSize * 0.75)}" text-anchor="middle" dominant-baseline="middle" class="name">${safeName}</text>
        <text x="${cx}" y="${Math.round(cy - qrSize * 0.4)}" text-anchor="middle" dominant-baseline="middle" class="school">${safeSchool}</text>
        <text x="${cx}" y="${Math.round(cy + qrSize * 0.05 + qrSize + 20)}" text-anchor="middle" dominant-baseline="middle" class="label">Scan for Details</text>
      </svg>
    `);

    // QR position (centered)
    const qrX = Math.round(cx - qrSize / 2);
    const qrY = Math.round(cy + qrSize * 0.05);

    // Save
    const filename = `card_${student._id}.png`;
    const outputPath = path.join('generated', filename);

    await sharp(templatePath)
      .composite([
        { input: textSvg, top: 0, left: 0 },
        { input: qrBuffer, top: qrY, left: qrX }
      ])
      .png()
      .toFile(outputPath);

    await Student.findByIdAndUpdate(student._id, {
      cardGenerated: true,
      cardPath: filename
    });

    res.json({ success: true, cardPath: `/generated/${filename}` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// POST: Generate Cards for ALL students
app.post('/api/generate-all', async (req, res) => {
  try {
    const students = await Student.find();
    const results = [];
    for (const student of students) {
      try {
        const r = await fetch(`http://localhost:${process.env.PORT || 3000}/api/generate-card/${student._id}`, { method: 'POST' });
        results.push({ id: student._id, success: r.ok });
      } catch (e) {
        results.push({ id: student._id, success: false });
      }
    }
    res.json({ success: true, results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// DELETE: Clear all students
app.delete('/api/students', async (req, res) => {
  try {
    await Student.deleteMany({});
    res.json({ success: true, message: 'All students deleted' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Server running at http://localhost:${PORT}`));