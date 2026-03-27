require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const QRCode = require('qrcode');
const sharp = require("sharp");
const archiver = require('archiver');

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
  email: { type: String, trim: true, lowercase: true, unique: true, sparse: true },
  phone: { type: String, trim: true, unique: true, sparse: true },
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

function normalizeEmail(v = '') {
  return String(v).trim().toLowerCase();
}

function normalizePhone(v = '') {
  return String(v).replace(/\D/g, '');
}

// POST: Upload Excel
app.post('/api/upload', upload.single('excel'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const workbook = xlsx.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    if (!data.length) return res.status(400).json({ error: 'Excel file is empty' });

    const normalized = data.map(row => {
      const obj = {};
      Object.keys(row).forEach(k => { obj[k.toLowerCase().trim()] = row[k]; });

      const email = normalizeEmail(obj['email'] || obj['email id'] || '');
      const phone = normalizePhone(obj['phone'] || obj['mobile'] || obj['contact'] || '');

      return {
        name: (obj['name'] || obj['student name'] || obj['studentname'] || '').toString().trim(),
        schoolName: (obj['school'] || obj['school name'] || obj['schoolname'] || obj['institution'] || '').toString().trim(),
        rollNo: obj['roll no'] || obj['rollno'] || obj['roll'] || obj['roll number'] || '',
        class: obj['class'] || obj['grade'] || obj['std'] || '',
        email: email || undefined,
        phone: phone || undefined,
        address: obj['address'] || obj['city'] || ''
      };
    }).filter(s => s.name && s.schoolName);

    // Remove duplicates inside uploaded file
    const seenEmails = new Set();
    const seenPhones = new Set();
    const uniqueFromFile = [];
    const duplicates = [];

    for (const s of normalized) {
      const e = s.email || '';
      const p = s.phone || '';

      if ((e && seenEmails.has(e)) || (p && seenPhones.has(p))) {
        duplicates.push({ reason: 'duplicate_in_file', email: e || null, phone: p || null, name: s.name });
        continue;
      }
      if (e) seenEmails.add(e);
      if (p) seenPhones.add(p);
      uniqueFromFile.push(s);
    }

    // Remove records already present in DB
    const emails = [...new Set(uniqueFromFile.map(s => s.email).filter(Boolean))];
    const phones = [...new Set(uniqueFromFile.map(s => s.phone).filter(Boolean))];

    const or = [];
    if (emails.length) or.push({ email: { $in: emails } });
    if (phones.length) or.push({ phone: { $in: phones } });

    const existing = or.length ? await Student.find({ $or: or }, { email: 1, phone: 1 }) : [];
    const existingEmails = new Set(existing.map(x => x.email).filter(Boolean));
    const existingPhones = new Set(existing.map(x => x.phone).filter(Boolean));

    const finalInsert = [];
    for (const s of uniqueFromFile) {
      if ((s.email && existingEmails.has(s.email)) || (s.phone && existingPhones.has(s.phone))) {
        duplicates.push({
          reason: 'already_in_database',
          email: s.email || null,
          phone: s.phone || null,
          name: s.name
        });
        continue;
      }
      finalInsert.push(s);
    }

    const inserted = finalInsert.length ? await Student.insertMany(finalInsert, { ordered: false }) : [];

    fs.unlinkSync(req.file.path);

    res.json({
      success: true,
      message: `${inserted.length} students uploaded successfully`,
      insertedCount: inserted.length,
      duplicateCount: duplicates.length,
      duplicates
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

function sanitizeFileName(value = '') {
  return String(value).replace(/[^a-z0-9_-]+/gi, '_').replace(/^_+|_+$/g, '');
}

async function generateCardForStudent(student) {
  // 4x6 inch card at 300 DPI
  const width = 1200;  // 4 inches * 300 DPI
  const height = 1800; // 6 inches * 300 DPI

  const baseUrl = process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`;
  const qrData = `${baseUrl}/student.html?id=${student._id}`;

  // Generate QR Code
  const qrSize = 280;
  const qrBuffer = await QRCode.toBuffer(qrData, {
    width: qrSize,
    margin: 1,
    color: { dark: '#1F2937', light: '#FFFFFF' }
  });

  // Escape student data for SVG
  const name = escapeSvgText((student.name || '').toUpperCase());
  const school = escapeSvgText(student.schoolName || '');
  const rollNo = escapeSvgText(student.rollNo || '');
  const className = escapeSvgText(student.class || '');
  const studentId = escapeSvgText(student._id.toString().substring(0, 12).toUpperCase());

  // Create professional card design
 const cardSvg = Buffer.from(`
    <svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="accentGrad" x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" style="stop-color:#FBBF24;stop-opacity:1" />
          <stop offset="100%" style="stop-color:#F59E0B;stop-opacity:1" />
        </linearGradient>
        <filter id="shadow">
          <feDropShadow dx="0" dy="4" stdDeviation="10" flood-opacity="0.12"/>
        </filter>
      </defs>

      <!-- Fully transparent background (nothing printed outside white box) -->
      <rect width="${width}" height="${height}" fill="none"/>

      <!-- WHITE INNER BOX ONLY — matches the center white area on the template -->
      <rect x="60" y="520" width="${width - 120}" height="${height - 660}"
            fill="#FFFFFF"
            rx="30"/>

      <!-- Participant Name Label -->
      <text x="${width / 2}" y="700"
            font-family="Arial, sans-serif"
            font-size="32"
            font-weight="600"
            fill="#6B7280"
            text-anchor="middle"
            letter-spacing="3">PARTICIPANT NAME</text>

      <!-- Student Name -->
      <text x="${width / 2}" y="800"
            font-family="Arial, sans-serif"
            font-size="80"
            font-weight="900"
            fill="#111827"
            text-anchor="middle"
            letter-spacing="1">${name}</text>

      <!-- Gold divider line -->
      <line x1="200" y1="860" x2="${width - 200}" y2="860"
            stroke="#FBBF24"
            stroke-width="5"/>

      <!-- School -->
      <text x="${width / 2}" y="960"
            font-family="Arial, sans-serif"
            font-size="40"
            font-weight="700"
            fill="#374151"
            text-anchor="middle">${school}</text>

      ${rollNo ? `
        <text x="200" y="1080"
              font-family="Arial, sans-serif"
              font-size="30"
              font-weight="600"
              fill="#6B7280">Roll No:</text>
        <text x="430" y="1080"
              font-family="Arial, sans-serif"
              font-size="30"
              font-weight="700"
              fill="#111827">${rollNo}</text>
      ` : ''}

      ${className ? `
        <text x="${width - 620}" y="1080"
              font-family="Arial, sans-serif"
              font-size="30"
              font-weight="600"
              fill="#6B7280">Class:</text>
        <text x="${width - 390}" y="1080"
              font-family="Arial, sans-serif"
              font-size="30"
              font-weight="700"
              fill="#111827">${className}</text>
      ` : ''}

      <!-- ID badge -->
      <rect x="320" y="1150" width="560" height="70"
            fill="#F3F4F6"
            rx="10"/>
      <text x="${width / 2}" y="1195"
            font-family="Courier New, monospace"
            font-size="34"
            font-weight="700"
            fill="#374151"
            text-anchor="middle"
            letter-spacing="4">ID: ${studentId}</text>

      <!-- Scan label -->
      <text x="${width / 2}" y="1300"
            font-family="Arial, sans-serif"
            font-size="28"
            font-weight="600"
            fill="#6B7280"
            text-anchor="middle"
            letter-spacing="1">SCAN FOR DETAILS</text>
    </svg>
  `);

  // Position QR code
  const qrLeft = Math.round((width - qrSize) / 2);
  const qrTop = 1360;

  const filename = `card_${student._id}.png`;
  const outputPath = path.join('generated', filename);

  // Compose the final card
await sharp({
  create: {
    width: width,
    height: height,
    channels: 4,
    background: { r: 255, g: 255, b: 255, alpha: 0 } // ← alpha: 0 = transparent
  }
})
  .composite([
    { input: cardSvg, top: 0, left: 0 },
    { input: qrBuffer, top: qrTop, left: qrLeft }
  ])
  .png({ quality: 100 })
  .toFile(outputPath);

  await Student.findByIdAndUpdate(student._id, {
    cardGenerated: true,
    cardPath: filename
  });

  return filename;
}

// POST: Generate Card for a student
app.post('/api/generate-card/:id', async (req, res) => {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ error: 'Student not found' });

    const filename = await generateCardForStudent(student);
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
        await generateCardForStudent(student);
        results.push({ id: student._id, success: true });
      } catch (e) {
        results.push({ id: student._id, success: false });
      }
    }

    res.json({ success: true, results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET: Download single student card
app.get('/api/download-card/:id', async (req, res) => {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ error: 'Student not found' });

    let filename = student.cardPath;
    const hasFile = filename && fs.existsSync(path.join('generated', filename));
    if (!hasFile) filename = await generateCardForStudent(student);

    const filePath = path.join('generated', filename);
    const pretty = sanitizeFileName(student.name || 'student');
    return res.download(filePath, `${pretty}_card.png`);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

// GET: Download all cards as ZIP
app.get('/api/download-all-cards', async (req, res) => {
  try {
    const students = await Student.find();
    if (!students.length) return res.status(404).json({ error: 'No students found' });

    for (const student of students) {
      const fileExists = student.cardPath && fs.existsSync(path.join('generated', student.cardPath));
      if (!fileExists) await generateCardForStudent(student);
    }

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="all-student-cards.zip"');

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('error', err => { throw err; });
    archive.pipe(res);

    for (const student of students) {
      const absPath = path.join('generated', student.cardPath);
      const pretty = sanitizeFileName(student.name || 'student');
      archive.file(absPath, { name: `${pretty}_${student._id}.png` });
    }

    await archive.finalize();
  } catch (err) {
    if (!res.headersSent) res.status(500).json({ error: err.message });
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