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
  // 60x80mm at 300 DPI = 708 x 945 px
  const width = 708;
  const height = 945;

  const baseUrl = process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`;
  const qrData = `${baseUrl}/student.html?id=${student._id}`;

  // QR code size
  const qrSize = 200;
  const qrBuffer = await QRCode.toBuffer(qrData, {
    width: qrSize,
    margin: 1,
    color: { dark: '#1F2937', light: '#FFFFFF' }
  });

  // Escape student data
  const name = escapeSvgText((student.name || '').toUpperCase());
  const school = escapeSvgText(student.schoolName || '');
  const rollNo = escapeSvgText(student.rollNo || '');
  const className = escapeSvgText(student.class || '');
  const studentId = escapeSvgText(student._id.toString().substring(0, 12).toUpperCase());

  // Layout constants (all coordinates relative to 708x945 white card)
 const cardSvg = Buffer.from(`
  <svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="topBar" x1="0%" y1="0%" x2="100%" y2="0%">
        <stop offset="0%" style="stop-color:#3B27A1;stop-opacity:1" />
        <stop offset="100%" style="stop-color:#4F3DC7;stop-opacity:1" />
      </linearGradient>
    </defs>

    <!-- White background -->
    <rect width="${width}" height="${height}" fill="#FFFFFF" rx="20"/>

    <!-- Top accent bar -->
    <rect x="0" y="0" width="${width}" height="18" fill="url(#topBar)"/>

    <!-- PARTICIPANT -->
    <text x="${width / 2}" y="95"
          font-family="Arial, sans-serif"
          font-size="28"
          font-weight="700"
          fill="#9CA3AF"
          text-anchor="middle"
          letter-spacing="4">PARTICIPANT</text>

    <!-- Name (REDUCED) -->
    <text x="${width / 2}" y="180"
          font-family="Arial, sans-serif"
          font-size="60"
          font-weight="900"
          fill="#111827"
          text-anchor="middle"
          letter-spacing="1.5">${name}</text>

    <!-- Gold divider -->
    <rect x="70" y="210" width="${width - 140}" height="7" fill="#FBBF24" rx="4"/>

    <!-- School (INCREASED) -->
    <text x="${width / 2}" y="290"
          font-family="Arial, sans-serif"
          font-size="42"
          font-weight="800"
          fill="#374151"
          text-anchor="middle">${school}</text>

    <!-- Roll + Class -->
    ${rollNo ? `
      <text x="90" y="370"
            font-family="Arial, sans-serif"
            font-size="28"
            fill="#6B7280">Roll No:</text>
      <text x="260" y="370"
            font-family="Arial, sans-serif"
            font-size="28"
            font-weight="700"
            fill="#111827">${rollNo}</text>
    ` : ''}

    ${className ? `
      <text x="${width - 300}" y="370"
            font-family="Arial, sans-serif"
            font-size="28"
            fill="#6B7280">Class:</text>
      <text x="${width - 160}" y="370"
            font-family="Arial, sans-serif"
            font-size="28"
            font-weight="700"
            fill="#111827">${className}</text>
    ` : ''}

    <!-- Divider -->
    <line x1="70" y1="410" x2="${width - 70}" y2="410"
          stroke="#E5E7EB" stroke-width="3"/>

    <!-- ID Badge -->
    <rect x="${(width - 480) / 2}" y="440" width="480" height="70"
          fill="#F3F4F6" rx="12"/>
    <text x="${width / 2}" y="485"
          font-family="Courier New, monospace"
          font-size="30"
          font-weight="700"
          fill="#374151"
          text-anchor="middle"
          letter-spacing="4">ID: ${studentId}</text>

    <!-- Scan label -->
    <text x="${width / 2}" y="600"
          font-family="Arial, sans-serif"
          font-size="26"
          font-weight="700"
          fill="#9CA3AF"
          text-anchor="middle"
          letter-spacing="3">SCAN FOR DETAILS</text>

    <!-- Bottom accent -->
    <rect x="0" y="${height - 18}" width="${width}" height="18" fill="url(#topBar)"/>
  </svg>
`);

  // QR code centered horizontally, below the scan label


// Center QR
const qrLeft = Math.round((width - qrSize) / 2);
const qrTop = 620;

const filename = `card_${student._id}.png`;
const outputPath = path.join('generated', filename);

await sharp({
  create: {
    width,
    height,
    channels: 4,
    background: { r: 255, g: 255, b: 255, alpha: 255 }
  }
})
  .composite([
    { input: cardSvg, top: 0, left: 0 },
    { input: qrBuffer, top: qrTop, left: qrLeft }
  ])
  .png({
    quality: 100,
    density: 300
  })
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

// POST: Add single student + optional card generation
app.post('/api/student', async (req, res) => {
  try {
    const {
      name,
      schoolName,
      rollNo = '',
      class: className = '',
      email = '',
      phone = '',
      address = '',
      generateCard = true
    } = req.body || {};

    if (!name || !schoolName) {
      return res.status(400).json({ error: 'name and schoolName are required' });
    }

    const nEmail = normalizeEmail(email);
    const nPhone = normalizePhone(phone);

    // duplicate check by email OR phone
    const or = [];
    if (nEmail) or.push({ email: nEmail });
    if (nPhone) or.push({ phone: nPhone });

    if (or.length) {
      const existing = await Student.findOne({ $or: or });
      if (existing) {
        return res.status(409).json({
          success: false,
          error: 'Student already exists with same email or phone'
        });
      }
    }

    const student = await Student.create({
      name: String(name).trim(),
      schoolName: String(schoolName).trim(),
      rollNo: String(rollNo).trim(),
      class: String(className).trim(),
      email: nEmail || undefined,
      phone: nPhone || undefined,
      address: String(address).trim()
    });

    let cardPath = null;
    if (generateCard) {
      const filename = await generateCardForStudent(student);
      cardPath = `/generated/${filename}`;
    }

    return res.status(201).json({
      success: true,
      message: 'Student created successfully',
      student,
      cardPath
    });
  } catch (err) {
    if (err && err.code === 11000) {
      return res.status(409).json({ success: false, error: 'Duplicate email or phone' });
    }
    return res.status(500).json({ success: false, error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Server running at http://localhost:${PORT}`));