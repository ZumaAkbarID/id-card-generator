const mysql = require('mysql2');
const path = require('path');
const fs = require('fs');
const sharp = require('sharp');
const QRCode = require('qrcode');
const dotenv = require('dotenv');
dotenv.config();

// Fungsi untuk menulis log ke file
function logToFile(message) {
  const logFilePath = path.join(__dirname, 'process.log');
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] ${message}\n`;

  // Tulis log ke file
  fs.appendFileSync(logFilePath, logMessage, 'utf8');
}

async function fetchAsistenAndCreateDir() {
  logToFile("CONNECTING TO DATABASE...");

  // Koneksi ke database MySQL
  const connection = mysql.createConnection({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASS,
    database: process.env.DB_NAME
  });

  connection.query(`
    UR QUERY
  `, async (err, results) => {
    if (err) {
      logToFile(`Error saat menjalankan query: ${err}`);
      connection.end();
      return;
    }
    logToFile("CONNECTED TO DATABASE");

    // Menentukan path direktori output
    const now = new Date();
    const formattedDate = now.toISOString().replace(/:/g, '-'); // Format menjadi 'YYYY-MM-DDTHH-MM-SS'
    const outputDir = path.join(__dirname, `../output/${formattedDate}`);

    // Membuat direktori utama jika belum ada
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Path direktori gambar
    const imageDir = path.join(__dirname, '/image/');

    // Proses setiap hasil query
    for (const row of results) {
      // Membuat sub-folder berdasarkan NPM
      const asistenDir = path.join(outputDir, row.npm);
      if (!fs.existsSync(asistenDir)) {
        fs.mkdirSync(asistenDir, { recursive: true });
      }

      // Path gambar yang akan dihasilkan
      const imagePathWebp = path.join(imageDir, `${row.npm}.webp`);
      const imagePathJpg = path.join(imageDir, `${row.npm}.jpg`);
      const finalImagePath = path.join(asistenDir, 'image.jpg');

      // Path QR code
      const qrCodePath = path.join(asistenDir, 'qrcode.png');

      // Mengonversi gambar dari webp ke jpg jika file gambar ada
      if (fs.existsSync(imagePathWebp)) {
        try {
          await sharp(imagePathWebp).toFile(finalImagePath);
          logToFile(`Gambar berhasil dikonversi ke jpg: ${finalImagePath}`);
        } catch (err) {
          logToFile(`Error saat mengonversi gambar ${imagePathWebp} ke jpg: ${err}`);
        }
      } else if (fs.existsSync(imagePathJpg)) {
        // Jika file sudah jpg, cukup salin ke folder output
        fs.copyFileSync(imagePathJpg, finalImagePath);
        logToFile(`Gambar jpg berhasil disalin: ${finalImagePath}`);
      } else {
        logToFile(`Gambar tidak ditemukan untuk NPM: ${row.npm}`);
      }

      // Generate QR code
      const qrCodeURL = `${process.env.VALIDATE_QR_URL}${row.npm}`;
      try {
        await QRCode.toFile(qrCodePath, qrCodeURL);
        logToFile(`QR code berhasil dibuat: ${qrCodePath}`);
      } catch (err) {
        logToFile(`Error saat membuat QR code untuk ${qrCodeURL}: ${err}`);
      }
    }

    connection.end();
    logToFile("DATABASE CONNECTION CLOSED");
  });
}

fetchAsistenAndCreateDir().catch(err => logToFile(`Error: ${err}`));
