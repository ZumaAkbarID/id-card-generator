const mysql = require('mysql2');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const sharp = require('sharp');
const QRCode = require('qrcode');
const dotenv = require('dotenv');
dotenv.config();

async function fetchAsistenAndCreateExcel() {
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

    console.info("CREATING WORKBOOK");
    // Buat workbook dan sheet Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Asisten');

    // Tambahkan header
    sheet.columns = [
      { header: 'NPM', key: 'npm', width: 30 },
      { header: 'Nama', key: 'nama', width: 30 },
      { header: 'Foto', key: 'foto', width: 30 },
      { header: 'QR Code', key: 'qrcode', width: 30 }
    ];

    // Path direktori gambar
    const imageDir = path.join(__dirname, '/image/');
    const tempImageDir = path.join(__dirname, '/temp_images/');
    const tempQRCodeDir = path.join(__dirname, '/temp_qrcodes/');

    // Membuat direktori sementara jika belum ada
    if (!fs.existsSync(tempImageDir)) {
      fs.mkdirSync(tempImageDir, { recursive: true });
    }
    if (!fs.existsSync(tempQRCodeDir)) {
      fs.mkdirSync(tempQRCodeDir, { recursive: true });
    }

    // Tambahkan data dari query ke sheet
    for (const [index, row] of results.entries()) {
      // Menentukan path gambar berdasarkan npm
      const imagePathWebp = path.join(imageDir, `${row.npm}.webp`);
      const imagePathJpg = path.join(imageDir, `${row.npm}.jpg`);
      const tempImagePath = path.join(tempImageDir, `${row.npm}.jpg`);

      // Menentukan path QR code
      const qrCodePath = path.join(tempQRCodeDir, `${row.npm}.png`);

      // Menambahkan baris ke worksheet
      const rowIndex = index + 2; // Karena header ada di baris pertama
      let imageAdded = false;

      // Menambahkan data NPM dan Nama ke worksheet
      sheet.getCell(`A${rowIndex}`).value = row.npm;
      sheet.getCell(`B${rowIndex}`).value = row.name;

      // Mengonversi gambar dari webp ke jpg jika file gambar ada
      if (fs.existsSync(imagePathWebp)) {
        try {
          await sharp(imagePathWebp).toFile(tempImagePath);

          // Menambahkan gambar ke worksheet
          const imageId = workbook.addImage({
            filename: tempImagePath,
            extension: 'jpg'
          });

          sheet.addImage(imageId, {
            tl: { col: 2, row: rowIndex - 1 }, // Kolom dan baris untuk posisi gambar
            ext: { width: 400, height: 600 } // Ukuran gambar
          });

          // Menyesuaikan tinggi baris dengan tinggi gambar
          const metadata = await sharp(tempImagePath).metadata();
          sheet.getRow(rowIndex).height = metadata.height ? metadata.height / 6 : 20; // Adjust divisor as needed

          imageAdded = true;
        } catch (err) {
          console.error(`Error saat mengonversi gambar ${imagePathWebp} ke ${tempImagePath}:`, err);
        }
      }

      // Jika gambar tidak ditemukan atau tidak bisa dikonversi, periksa format jpg
      if (!imageAdded && fs.existsSync(imagePathJpg)) {
        try {
          // Menambahkan gambar ke worksheet
          const imageId = workbook.addImage({
            filename: imagePathJpg,
            extension: 'jpg'
          });

          sheet.addImage(imageId, {
            tl: { col: 2, row: rowIndex - 1 }, // Kolom dan baris untuk posisi gambar
            ext: { width: 400, height: 600 } // Ukuran gambar
          });

          // Menyesuaikan tinggi baris dengan tinggi gambar
          const metadata = await sharp(imagePathJpg).metadata();
          sheet.getRow(rowIndex).height = metadata.height ? metadata.height / 6 : 20; // Adjust divisor as needed

          imageAdded = true;
        } catch (err) {
          console.error(`Error saat memasukkan gambar ${imagePathJpg} ke Excel:`, err);
        }
      }

      // Jika gambar tidak ditemukan, tambahkan teks "Foto belum diatur"
      if (!imageAdded) {
        sheet.getCell(`C${rowIndex}`).value = 'Foto belum diatur';
      } else {
        sheet.getCell(`C${rowIndex}`).value = ''; // Kosongkan jika gambar berhasil ditambahkan
      }

      // Generate QR code
      const qrCodeURL = `${process.env.VALIDATE_QR_URL}${row.npm}`;
      try {
        await QRCode.toFile(qrCodePath, qrCodeURL);

        // Menambahkan QR code ke worksheet
        const qrCodeId = workbook.addImage({
          filename: qrCodePath,
          extension: 'png'
        });

        sheet.addImage(qrCodeId, {
          tl: { col: 3, row: rowIndex - 1 }, // Kolom dan baris untuk posisi QR code
          ext: { width: 400, height: 400 } // Ukuran QR code
        });

      } catch (err) {
        console.error(`Error saat membuat QR code untuk ${qrCodeURL}:`, err);
      }
    }

    // Menentukan path direktori dan file
    const outputDir = path.join(__dirname, '../output');
    const now = new Date();
    const formattedDate = now.toISOString().replace(/:/g, '-'); // Format menjadi 'YYYY-MM-DDTHH-MM-SS'
    const fileName = `asisten-${formattedDate}.xlsx`;
    const filePath = path.join(outputDir, fileName);

    // Membuat direktori jika belum ada
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Simpan workbook ke file
    try {
      await workbook.xlsx.writeFile(filePath);
      console.log(`File Excel berhasil dibuat! Nama file: ${fileName}`);
    } catch (err) {
      console.error('Error saat menyimpan file Excel:', err);
    } finally {
      connection.end();

      // Menghapus gambar sementara
      fs.readdir(tempImageDir, (err, files) => {
        if (err) {
          console.error('Error saat membaca direktori sementara:', err);
          return;
        }
        files.forEach(file => fs.unlinkSync(path.join(tempImageDir, file)));
        fs.rmdirSync(tempImageDir);
      });

      // Menghapus QR code sementara
      fs.readdir(tempQRCodeDir, (err, files) => {
        if (err) {
          console.error('Error saat membaca direktori QR code sementara:', err);
          return;
        }
        files.forEach(file => fs.unlinkSync(path.join(tempQRCodeDir, file)));
        fs.rmdirSync(tempQRCodeDir);
      });
    }
  });
}

fetchAsistenAndCreateExcel().catch(err => console.error('Error:', err));
