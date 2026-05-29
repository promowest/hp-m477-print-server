const express = require('express');
const cors = require('cors');
const multer = require('multer');
const http = require('http');
const { execSync, exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
app.use(cors());
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

const PRINTER_HOST = process.env.PRINTER_HOST;
const PRINTER_PORT = parseInt(process.env.PRINTER_PORT || '631');
const PRINTER_PATH = process.env.PRINTER_PATH || '/ipp/print';

// ─────────────────────────────────────────────
// Tipuri de fișiere suportate
// ─────────────────────────────────────────────
const SUPPORTED_TYPES = {
  // Trimise direct la imprimantă (HP M477fdn suportă nativ)
  'application/pdf':  { direct: true,  ext: 'pdf'  },
  'image/jpeg':       { direct: true,  ext: 'jpg'  },
  'image/png':        { direct: true,  ext: 'png'  },

  // Convertite în PDF via LibreOffice
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': { direct: false, ext: 'docx' },
  'application/msword':                                                        { direct: false, ext: 'doc'  },
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':        { direct: false, ext: 'xlsx' },
  'application/vnd.ms-excel':                                                  { direct: false, ext: 'xls'  },
  'application/vnd.openxmlformats-officedocument.presentationml.presentation':{ direct: false, ext: 'pptx' },
  'application/vnd.ms-powerpoint':                                             { direct: false, ext: 'ppt'  },
  'text/plain':       { direct: false, ext: 'txt'  },
  'text/csv':         { direct: false, ext: 'csv'  },
};

// ─────────────────────────────────────────────
// Utilitar: fix DPI pentru JPEG
// ─────────────────────────────────────────────
function setJPEGDPI(buffer, dpi) {
  if (buffer[0] === 0xFF && buffer[1] === 0xD8 &&
      buffer[2] === 0xFF && buffer[3] === 0xE0) {
    const modified = Buffer.from(buffer);
    modified[11] = 1;
    modified[12] = (dpi >> 8) & 0xFF;
    modified[13] = dpi & 0xFF;
    modified[14] = (dpi >> 8) & 0xFF;
    modified[15] = dpi & 0xFF;
    return modified;
  }
  return buffer;
}

// ─────────────────────────────────────────────
// Conversie fișiere Office/text → PDF via LibreOffice
// ─────────────────────────────────────────────
function convertToPDF(buffer, ext) {
  return new Promise((resolve, reject) => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'print-'));
    const inputFile = path.join(tmpDir, `input.${ext}`);
    const outputFile = path.join(tmpDir, 'input.pdf');

    try {
      fs.writeFileSync(inputFile, buffer);

      exec(
        `soffice --headless --convert-to pdf --outdir "${tmpDir}" "${inputFile}"`,
        { timeout: 60000, env: { ...process.env, HOME: tmpDir } },
        (err, stdout, stderr) => {
          if (err) {
            cleanup(tmpDir);
            return reject(new Error(`LibreOffice error: ${stderr || err.message}`));
          }
          if (!fs.existsSync(outputFile)) {
            cleanup(tmpDir);
            return reject(new Error('PDF output not found after conversion'));
          }
          const pdfBuffer = fs.readFileSync(outputFile);
          cleanup(tmpDir);
          resolve(pdfBuffer);
        }
      );
    } catch (e) {
      cleanup(tmpDir);
      reject(e);
    }
  });
}

function cleanup(dir) {
  try { execSync(`rm -rf "${dir}"`); } catch (_) {}
}

// ─────────────────────────────────────────────
// Construire pachet IPP și trimitere la imprimantă
// ─────────────────────────────────────────────
function sendIPP(buffer, mimeType, res) {
  const printerUri = `ipp://${PRINTER_HOST}:${PRINTER_PORT}${PRINTER_PATH}`;

  const writeAttr = (tag, name, value) => {
    const n = Buffer.from(name, 'utf8');
    const v = Buffer.from(value, 'utf8');
    const b = Buffer.allocUnsafe(1 + 2 + n.length + 2 + v.length);
    let o = 0;
    b.writeUInt8(tag, o++);
    b.writeUInt16BE(n.length, o); o += 2;
    n.copy(b, o); o += n.length;
    b.writeUInt16BE(v.length, o); o += 2;
    v.copy(b, o);
    return b;
  };

  const hdr = Buffer.alloc(8);
  hdr.writeUInt8(2, 0);
  hdr.writeUInt8(0, 1);
  hdr.writeUInt16BE(0x0002, 2);
  hdr.writeInt32BE(1, 4);

  const attrs = Buffer.concat([
    writeAttr(0x47, 'attributes-charset', 'utf-8'),
    writeAttr(0x48, 'attributes-natural-language', 'en'),
    writeAttr(0x45, 'printer-uri', printerUri),
    writeAttr(0x42, 'requesting-user-name', 'MobileApp'),
    writeAttr(0x42, 'job-name', `PrintJob-${Date.now()}`),
    writeAttr(0x49, 'document-format', mimeType),
  ]);

  const ippBody = Buffer.concat([hdr, Buffer.from([0x01]), attrs, Buffer.from([0x03]), buffer]);

  const options = {
    hostname: PRINTER_HOST,
    port: PRINTER_PORT,
    path: PRINTER_PATH,
    method: 'POST',
    headers: {
      'Content-Type': 'application/ipp',
      'Content-Length': ippBody.length,
    },
  };

  let responded = false;
  const request = http.request(options, (response) => {
    const chunks = [];
    response.on('data', c => chunks.push(c));
    response.on('end', () => {
      const body = Buffer.concat(chunks);
      if (body.length < 4) {
        if (!responded) { responded = true; res.status(500).json({ error: 'Răspuns IPP invalid (prea scurt)' }); }
        return;
      }
      const ippStatus = body.readUInt16BE(2);
      console.log(`IPP status: 0x${ippStatus.toString(16)}`);
      if (!responded) {
        responded = true;
        if (ippStatus === 0x0000) {
          res.json({ success: true, message: 'Documentul a fost trimis la imprimantă.' });
        } else {
          res.status(500).json({ error: `IPP error: 0x${ippStatus.toString(16)}` });
        }
      }
    });
  });

  request.on('error', (err) => {
    console.error('IPP connection error:', err.message);
    if (!responded) {
      responded = true;
      res.status(500).json({ error: `Nu s-a putut conecta la imprimantă: ${err.message}` });
    }
  });

  request.setTimeout(30000, () => {
    request.destroy();
    if (!responded) {
      responded = true;
      res.status(504).json({ error: 'Timeout: imprimanta nu a răspuns în 30 secunde' });
    }
  });

  request.write(ippBody);
  request.end();
}

// ─────────────────────────────────────────────
// Endpoint status
// ─────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({
    status: 'HP M477fdn Print Server online',
    printer: `${PRINTER_HOST}:${PRINTER_PORT}${PRINTER_PATH}`,
    supportedTypes: Object.keys(SUPPORTED_TYPES),
  });
});

// ─────────────────────────────────────────────
// Endpoint principal: POST /print
// ─────────────────────────────────────────────
app.post('/print', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Niciun fișier primit' });
  if (!PRINTER_HOST) return res.status(500).json({ error: 'PRINTER_HOST nu este configurat pe server' });

  const mimeType = req.file.mimetype || 'application/octet-stream';
  const fileSize = req.file.size;
  console.log(`Fișier primit: ${req.file.originalname}, ${fileSize} bytes, tip: ${mimeType}`);

  const typeInfo = SUPPORTED_TYPES[mimeType];
  if (!typeInfo) {
    return res.status(415).json({
      error: `Tip de fișier nesupportat: ${mimeType}`,
      supported: Object.keys(SUPPORTED_TYPES),
    });
  }

  try {
    let printBuffer = req.file.buffer;
    let printMime = mimeType;

    if (!typeInfo.direct) {
      // Conversie în PDF
      console.log(`Convertesc ${typeInfo.ext} → PDF via LibreOffice...`);
      printBuffer = await convertToPDF(req.file.buffer, typeInfo.ext);
      printMime = 'application/pdf';
      console.log(`Conversie reușită: ${printBuffer.length} bytes PDF`);
    } else if (mimeType === 'image/jpeg') {
      printBuffer = setJPEGDPI(printBuffer, 300);
    }

    sendIPP(printBuffer, printMime, res);
  } catch (err) {
    console.error('Eroare procesare:', err.message);
    res.status(500).json({ error: `Eroare la procesarea fișierului: ${err.message}` });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`HP M477fdn Print Server pornit pe portul ${PORT}`);
  console.log(`Printer: ${PRINTER_HOST}:${PRINTER_PORT}${PRINTER_PATH}`);
});
