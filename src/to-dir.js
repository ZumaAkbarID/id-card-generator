const mysql = require('mysql2');
const path = require('path');
const fs = require('fs');
const sharp = require('sharp');
const QRCode = require('qrcode');
const dotenv = require('dotenv');
dotenv.config();

async function fetchAsistenAndCreateDir() {
  console.info("CONNECTING TO DATABASE...");

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
      console.error('Error saat menjalankan query:', err);
      connection.end();
      return;
    }
    console.info("CONNECTED");

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
          console.log(`Gambar berhasil dikonversi ke jpg: ${finalImagePath}`);
        } catch (err) {
          console.error(`Error saat mengonversi gambar ${imagePathWebp} ke jpg:`, err);
        }
      } else if (fs.existsSync(imagePathJpg)) {
        // Jika file sudah jpg, cukup salin ke folder output
        fs.copyFileSync(imagePathJpg, finalImagePath);
        console.log(`Gambar jpg berhasil disalin: ${finalImagePath}`);
      } else {
        console.log(`Gambar tidak ditemukan untuk NPM: ${row.npm}`);
      }

      // Generate QR code
      const qrCodeURL = `${process.env.VALIDATE_QR_URL}${row.npm}`;
      try {
        await QRCode.toFile(qrCodePath, qrCodeURL);
        console.log(`QR code berhasil dibuat: ${qrCodePath}`);
      } catch (err) {
        console.error(`Error saat membuat QR code untuk ${qrCodeURL}:`, err);
      }
    }

    connection.end();
  });
}

fetchAsistenAndCreateDir().catch(err => console.error('Error:', err));
