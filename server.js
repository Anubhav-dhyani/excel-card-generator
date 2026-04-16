require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
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
  schoolName: { type: String, default: '' },
  rollNo: String,
  class: String,
  email: { type: String, trim: true, lowercase: true, unique: true, sparse: true },
  phone: { type: String, trim: true, unique: true, sparse: true },
  address: String,
  uploadedAt: { type: Date, default: Date.now },
  passToken: { type: String, index: true },
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

const QR_SECRET = process.env.QR_SECRET;
if (!QR_SECRET) {
  console.warn('⚠️ QR_SECRET is not set. Set QR_SECRET in .env for signed/verified QR codes.');
}

function base64urlEncode(buf) {
  return Buffer.from(buf)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function base64urlDecode(input) {
  const str = String(input || '');
  if (!/^[A-Za-z0-9_-]+$/.test(str)) {
    throw new Error('Invalid token encoding');
  }
  let b64 = str.replace(/-/g, '+').replace(/_/g, '/');
  while (b64.length % 4 !== 0) b64 += '=';
  return Buffer.from(b64, 'base64');
}

function issuePassToken(studentId) {
  if (!QR_SECRET) throw new Error('QR_SECRET is not configured');
  const payload = {
    sid: String(studentId),
    iat: Math.floor(Date.now() / 1000)
  };
  const payloadB64 = base64urlEncode(JSON.stringify(payload));
  const sig = crypto.createHmac('sha256', QR_SECRET).update(payloadB64).digest();
  return `${payloadB64}.${base64urlEncode(sig)}`;
}

function verifyPassToken(token) {
  if (!QR_SECRET) throw new Error('QR_SECRET is not configured');
  const raw = String(token || '');
  if (!raw || raw.length > 2048) throw new Error('Invalid token');

  const parts = raw.split('.');
  if (parts.length !== 2) throw new Error('Invalid token');

  const [payloadB64, sigB64] = parts;
  const sig = base64urlDecode(sigB64);
  const expected = crypto.createHmac('sha256', QR_SECRET).update(payloadB64).digest();

  if (sig.length !== expected.length || !crypto.timingSafeEqual(sig, expected)) {
    throw new Error('Invalid token signature');
  }

  const payloadJson = base64urlDecode(payloadB64).toString('utf8');
  const payload = JSON.parse(payloadJson);
  if (!payload || typeof payload.sid !== 'string') throw new Error('Invalid token payload');

  const maxAgeDays = Number(process.env.PASS_TOKEN_MAX_AGE_DAYS || 3650);
  const now = Math.floor(Date.now() / 1000);
  if (payload.iat && typeof payload.iat === 'number') {
    if (payload.iat > now + 300) throw new Error('Invalid token timestamp');
    const maxAgeSec = Math.max(1, maxAgeDays) * 24 * 60 * 60;
    if (now - payload.iat > maxAgeSec) throw new Error('Token expired');
  }

  return payload;
}

function buildVerificationUrl(token) {
  const baseUrl = process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`;
  return `${baseUrl}/student.html?token=${encodeURIComponent(token)}`;
}

// GET: Verified student details via signed token
app.get('/api/verify', async (req, res) => {
  try {
    const token = String(req.query.token || '');
    const { sid } = verifyPassToken(token);

    const student = await Student.findById(sid);
    if (!student) return res.status(404).json({ verified: false, error: 'Student not found' });

    return res.json({ verified: true, student });
  } catch (err) {
    return res.status(400).json({ verified: false, error: 'Invalid or expired token' });
  }
});

// GET: QR code image for a signed token
app.get('/api/qr', async (req, res) => {
  try {
    const token = String(req.query.token || '');
    verifyPassToken(token);

    const qrData = buildVerificationUrl(token);
    const qrBuffer = await QRCode.toBuffer(qrData, {
      width: 768,
      margin: 0,
      color: { dark: '#111111', light: '#FFFFFF' }
    });

    res.setHeader('Content-Type', 'image/png');
    res.setHeader('Cache-Control', 'no-store');
    return res.send(qrBuffer);
  } catch (err) {
    return res.status(400).json({ error: 'Invalid token' });
  }
});

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
    }).filter(s => s.name);

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

function wrapTextIntoLines(value = '', maxCharsPerLine = 14, maxLines = 3) {
  const words = String(value).trim().split(/\s+/).filter(Boolean);
  if (!words.length) return [''];

  const lines = [];
  let currentLine = '';

  for (const word of words) {
    const candidate = currentLine ? `${currentLine} ${word}` : word;
    if (candidate.length <= maxCharsPerLine) {
      currentLine = candidate;
      continue;
    }

    if (currentLine) {
      lines.push(currentLine);
      currentLine = word;
    } else {
      lines.push(word);
      currentLine = '';
    }

    if (lines.length === maxLines - 1) {
      break;
    }
  }

  const remainingWords = [];
  if (currentLine) remainingWords.push(currentLine);

  const usedWords = lines.join(' ').split(/\s+/).filter(Boolean).length + (currentLine ? currentLine.split(/\s+/).length : 0);
  const leftovers = words.slice(usedWords);
  if (leftovers.length) remainingWords.push(leftovers.join(' '));

  if (remainingWords.length) {
    const finalLine = remainingWords.join(' ').trim();
    lines.push(finalLine);
  }

  if (lines.length > maxLines) {
    lines.length = maxLines;
  }

  if (lines.length === maxLines && words.join(' ') !== lines.join(' ')) {
    lines[maxLines - 1] = `${lines[maxLines - 1].slice(0, Math.max(0, maxCharsPerLine - 3)).trim()}...`;
  }

  return lines;
}

async function generateEntryPassForStudent(student) {
  const templatePath = path.join(__dirname, 'images', 'templete.webp');

  const token = student.passToken || issuePassToken(student._id.toString());
  const qrData = buildVerificationUrl(token);

  const meta = await sharp(templatePath).metadata();
  const width = meta.width || 1600;
  const height = meta.height || 607;

  const rawName = (student.name || '').toUpperCase();
  const nameLines = wrapTextIntoLines(rawName, 12, 2).map(escapeSvgText);
  const longestNameLine = nameLines.reduce((max, line) => Math.max(max, line.length), 0);
  const nameFontSize = longestNameLine > 10 ? 28 : nameLines.length > 1 ? 32 : 36;
  const nameLineHeight = nameFontSize + 6;
  const nameX = 206;
  const nameY = 276;

  const qrBoxLeft = 56;
  const qrBoxTop = 356;
  const qrBoxSize = 288;
  const qrPadding = 4;
  const qrLeft = qrBoxLeft + qrPadding;
  const qrTop = qrBoxTop + qrPadding;
  const qrSize = qrBoxSize - (qrPadding * 2);
  const qrBuffer = await QRCode.toBuffer(qrData, {
    width: qrSize,
    margin: 0,
    color: { dark: '#111111', light: '#FFFFFF' }
  });

  const overlaySvg = Buffer.from(`
    <svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
      <rect x="${qrBoxLeft}" y="${qrBoxTop}" width="${qrBoxSize}" height="${qrBoxSize}" fill="#FFFFFF"/>
      <rect x="${qrBoxLeft}" y="${qrBoxTop}" width="${qrBoxSize}" height="${qrBoxSize}" fill="none" stroke="#111111" stroke-width="4"/>

      <text x="${nameX}" y="${nameY}"
            font-family="Arial, sans-serif"
            font-size="${nameFontSize}"
            font-weight="900"
            fill="#111111"
            text-anchor="middle"
            letter-spacing="0.5">${nameLines
              .map((line, index) => `<tspan x=\"${nameX}\" dy=\"${index === 0 ? 0 : nameLineHeight}\">${line}</tspan>`)
              .join('')}</text>
    </svg>
  `);

  const filename = `pass_${student._id}.png`;
  const outputPath = path.join('generated', filename);

  await sharp(templatePath)
    .composite([
      { input: overlaySvg, top: 0, left: 0 },
      { input: qrBuffer, top: qrTop, left: qrLeft }
    ])
    .png({ quality: 100 })
    .toFile(outputPath);

  await Student.findByIdAndUpdate(student._id, {
    passToken: token,
    cardGenerated: true,
    cardPath: filename
  });

  return { filename, token };
}

// POST: Generate Entry Pass for a student (alias: /api/generate-card/:id)
async function generatePassById(req, res) {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ error: 'Student not found' });

    const { filename, token } = await generateEntryPassForStudent(student);
    return res.json({
      success: true,
      passPath: `/generated/${filename}`,
      token,
      verifyUrl: buildVerificationUrl(token)
    });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}

app.post('/api/generate-pass/:id', generatePassById);
app.post('/api/generate-card/:id', generatePassById);

// POST: Generate Entry Passes for ALL students
app.post('/api/generate-all', async (req, res) => {
  try {
    const students = await Student.find();
    const results = [];

    for (const student of students) {
      try {
        await generateEntryPassForStudent(student);
        results.push({ id: student._id, name: student.name, success: true });
      } catch (e) {
        results.push({ id: student._id, name: student.name, success: false, error: e.message });
      }
    }

    return res.json({ success: true, results });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

// GET: Download single entry pass (alias: /api/download-card/:id)
async function downloadPassById(req, res) {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ error: 'Student not found' });

    let filename = student.cardPath;
    const hasFile = filename && fs.existsSync(path.join('generated', filename));
    if (!hasFile) {
      filename = (await generateEntryPassForStudent(student)).filename;
    }

    const filePath = path.join('generated', filename);
    const pretty = sanitizeFileName(student.name || 'student');
    return res.download(filePath, `${pretty}_entry-pass.png`);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}

app.get('/api/download-pass/:id', downloadPassById);
app.get('/api/download-card/:id', downloadPassById);

// GET: Download all entry passes as ZIP (alias: /api/download-all-cards)
async function downloadAllPasses(req, res) {
  try {
    const students = await Student.find();
    if (!students.length) return res.status(404).json({ error: 'No students found' });

    for (const student of students) {
      const fileExists = student.cardPath && fs.existsSync(path.join('generated', student.cardPath));
      if (!fileExists) await generateEntryPassForStudent(student);
    }

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="all-entry-passes.zip"');

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('error', err => {
      throw err;
    });
    archive.pipe(res);

    for (const student of students) {
      const absPath = path.join('generated', student.cardPath);
      const pretty = sanitizeFileName(student.name || 'student');
      archive.file(absPath, { name: `${pretty}_${student._id}_entry-pass.png` });
    }

    await archive.finalize();
  } catch (err) {
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
}

app.get('/api/download-all-passes', downloadAllPasses);
app.get('/api/download-all-cards', downloadAllPasses);

// DELETE: Clear all students
app.delete('/api/students', async (req, res) => {
  try {
    await Student.deleteMany({});
    res.json({ success: true, message: 'All students deleted' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST: Add single student + optional entry pass generation
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
      generatePass,
      generateCard
    } = req.body || {};

    const shouldGeneratePass =
      typeof generatePass === 'boolean'
        ? generatePass
        : typeof generateCard === 'boolean'
          ? generateCard
          : true;

    if (!name) {
      return res.status(400).json({ error: 'name is required' });
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
      schoolName: String(schoolName || '').trim(),
      rollNo: String(rollNo).trim(),
      class: String(className).trim(),
      email: nEmail || undefined,
      phone: nPhone || undefined,
      address: String(address).trim()
    });

    let passPath = null;
    let token = null;
    let warning = null;

    if (shouldGeneratePass) {
      try {
        const result = await generateEntryPassForStudent(student);
        passPath = `/generated/${result.filename}`;
        token = result.token;
      } catch (passErr) {
        warning = `Student saved, but entry pass generation failed: ${passErr.message}`;
        console.error('Entry pass generation failed for new student:', passErr);
      }
    }

    return res.status(201).json({
      success: true,
      message: warning ? 'Student created, but entry pass generation failed' : 'Student created successfully',
      student,
      passPath,
      token,
      verifyUrl: token ? buildVerificationUrl(token) : null,
      warning
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
