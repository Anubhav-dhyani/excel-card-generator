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
  // keep digits only; adjust if you want country-code format
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

    // 1) Remove duplicates inside uploaded file itself
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

    // 2) Remove records already present in DB by email/phone
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

function getTemplatePath() {
  const imageFiles = fs.readdirSync('images').filter(f =>
    ['.jpg', '.jpeg', '.png'].includes(path.extname(f).toLowerCase())
  );
  if (!imageFiles.length) throw new Error('No template image found in /images folder');
  return path.join('images', imageFiles[0]);
}

async function generateCardForStudent(student) {
  const templatePath = getTemplatePath();
  const templateSharp = sharp(templatePath);
  const meta = await templateSharp.metadata();

  const width = meta.width || 1200;
  const height = meta.height || 800;
  const cx = Math.floor(width / 2);
  const cy = Math.floor(height / 2);

  const baseUrl = process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`;
  const qrData = `${baseUrl}/student.html?id=${student._id}`;

  const qrSize = Math.floor(Math.min(width, height) * 0.22);
  const qrBuffer = await QRCode.toBuffer(qrData, {
    width: qrSize,
    margin: 1,
    color: { dark: '#000000', light: '#ffffff' }
  });

  const nameFontSize = Math.floor(width * 0.055);
  const schoolFontSize = Math.floor(width * 0.035);
  const labelFontSize = Math.floor(width * 0.022);

  const safeName = escapeSvgText((student.name || '').toUpperCase());
  const safeSchool = escapeSvgText(student.schoolName || '');

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

  const qrX = Math.round(cx - qrSize / 2);
  const qrY = Math.round(cy + qrSize * 0.05);

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