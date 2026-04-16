const fs = require('fs');
let content = fs.readFileSync('server.js', 'utf8');
content = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

// Fix single download - use explicit headers
const oldSingle = `async function downloadPassById(req, res) {
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
    return res.download(filePath, \`\${pretty}_entry-pass.png\`);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}`;

const newSingle = `async function downloadPassById(req, res) {
  try {
    const student = await Student.findById(req.params.id);
    if (!student) return res.status(404).json({ error: 'Student not found' });

    let filename = student.cardPath;
    const hasFile = filename && fs.existsSync(path.join('generated', filename));
    if (!hasFile) {
      const result = await generateEntryPassForStudent(student);
      filename = result.filename;
    }

    const filePath = path.resolve('generated', filename);
    const pretty = sanitizeFileName(student.name || 'student') || 'student';
    const downloadName = \`\${pretty}_entry-pass.png\`;

    res.setHeader('Content-Type', 'image/png');
    res.setHeader('Content-Disposition', \`attachment; filename="\${downloadName}"\`);
    return res.sendFile(filePath);
  } catch (err) {
    console.error('Download error:', err);
    return res.status(500).json({ error: err.message });
  }
}`;

if (content.includes(oldSingle)) {
  content = content.replace(oldSingle, newSingle);
  console.log('OK: single download fixed');
} else {
  console.log('ERROR: single download block not found');
}

fs.writeFileSync('server.js', content, 'utf8');
console.log('DONE');
