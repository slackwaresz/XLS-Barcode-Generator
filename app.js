const express = require('express');
const multer = require('multer');
const xlsxPopulate = require('xlsx-populate');
const bwipjs = require('bwip-js');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const maxHeaderSize = 64 * 1024;
const app = express({ 'settings': { 'maxHttpHeaderSize': maxHeaderSize } });
const port = 3000;

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.use(cors());

app.use('/barcode_images', express.static('barcode_images'));

const barcodeImageDir = './barcode_images';
if (!fs.existsSync(barcodeImageDir)) {
  fs.mkdirSync(barcodeImageDir);
}

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

app.post('/upload', upload.single('uploadedFile'), async (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file was uploaded.');
  }

  const buffer = req.file.buffer;
  const barcodeType = req.query.barcodeType;

  try {
    const workbook = await xlsxPopulate.fromDataAsync(buffer);
    const sheet = workbook.sheet(0);

    const barcodes = [];

    for (let rowNumber = 2; rowNumber <= sheet.usedRange().endCell().rowNumber(); rowNumber++) {
      const materialCode = sheet.cell(`A${rowNumber}`).value();
      const materialName = sheet.cell(`B${rowNumber}`).value();
      let barcodeText = sheet.cell(`C${rowNumber}`).value();

      if (typeof barcodeText !== 'string') {
        barcodeText = String(barcodeText);
      }

      if (barcodeType === 'EAN-13') {
        if (barcodeText.length < 12) {
            return res.status(400).send(`编码长度必须为至少12位 (行号: ${rowNumber})`);
        }

        if (barcodeText) {
          const validatedBarcode = validateAndFixEAN13(barcodeText);
          const pngPath = await generateAndSaveBarcode(validatedBarcode);
          const base64Image = await convertToBase64(pngPath);
          barcodes.push({ materialCode, materialName, barcode: validatedBarcode, png: base64Image });
        }
      } else if (barcodeType === 'code128B') {
        const pngPath = await generateAndSaveCode128BBarcode(barcodeText);
        const base64Image = await convertToBase64(pngPath);
        barcodes.push({ materialCode, materialName, barcode: barcodeText, png: base64Image });
      }
    }

    res.json(barcodes);
  } catch (err) {
    console.error(err);
    res.status(500).send('Error processing the file.');
  }
});

function validateAndFixEAN13(barcode) {
  const cleanedBarcode = barcode.replace(/\D/g, '');

  if (cleanedBarcode.length === 12) {
    return cleanedBarcode;
  } else if (cleanedBarcode.length === 13) {
    const truncatedBarcode = cleanedBarcode.substring(0, 12);
    return truncatedBarcode + calculateEAN13CheckDigit(truncatedBarcode);
  } else {
    throw new Error('EAN-13 data must be 12 or 13 digits.');
  }
}

function calculateEAN13CheckDigit(data) {
  if (data.length !== 12) {
    throw new Error('EAN-13 data must be 12 digits to calculate the check digit.');
  }

  const digits = data.split('').map(Number);

  const oddSum = digits
    .filter((_, index) => index % 2 === 0)
    .reduce((acc, curr) => acc + curr, 0);

  const evenSum = digits
    .filter((_, index) => index % 2 !== 0)
    .reduce((acc, curr) => acc + curr, 0);

  const totalSum = evenSum * 3 + oddSum;
  const checkDigit = (10 - (totalSum % 10)) % 10;

  return String(checkDigit);
}

async function generateAndSaveBarcode(barcodeText) {
  return new Promise((resolve, reject) => {
    bwipjs.toBuffer({
      bcid: 'ean13',
      text: barcodeText,
      scale: 3,
      includetext: true,
    }, (err, pngBuffer) => {
      if (err) {
        reject(err);
      } else {
        const timestamp = Date.now();
        const pngPath = path.join(barcodeImageDir, `barcode_${timestamp}.png`);
        fs.writeFileSync(pngPath, pngBuffer);
        resolve(pngPath);
      }
    });
  });
}

async function generateAndSaveCode128BBarcode(barcodeText) {
  return new Promise((resolve, reject) => {
    bwipjs.toBuffer({
      bcid: 'code128',
      text: barcodeText,
      scale: 3,
      includetext: true,
      textxalign: 'center',
    }, (err, pngBuffer) => {
      if (err) {
        reject(err);
      } else {
        const timestamp = Date.now();
        const pngPath = path.join(barcodeImageDir, `barcode_${timestamp}.png`);
        fs.writeFileSync(pngPath, pngBuffer);
        resolve(pngPath);
      }
    });
  });
}

async function convertToBase64(imagePath) {
  return new Promise((resolve, reject) => {
    fs.readFile(imagePath, (err, data) => {
      if (err) {
        reject(err);
      } else {
        const base64Image = data.toString('base64');
        resolve(base64Image);
      }
    });
  })
}

app.post('/download-json', (req, res) => {
  let barcodes = req.body.barcodes;
  if (!barcodes) {
    return res.status(400).send('No barcodes data provided for download.');
  }

  const jsonContent = JSON.stringify({ barcodes }, null, 2);

  res.setHeader('Content-Type', 'application/json');
  res.setHeader('Content-Disposition', 'attachment; filename=barcodes.json');
  res.send(jsonContent);
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
