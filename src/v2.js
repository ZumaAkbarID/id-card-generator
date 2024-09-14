const mysql = require('mysql2');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const sharp = require('sharp');
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
      { header: 'Foto', key: 'foto', width: 30 }
    ];

    // Path direktori gambar
    const imageDir = path.join(__dirname, '/image/');
    const tempImageDir = path.join(__dirname, '/temp_images/');

    // Membuat direktori sementara jika belum ada
    if (!fs.existsSync(tempImageDir)) {
      fs.mkdirSync(tempImageDir, { recursive: true });
    }

    // Tambahkan data dari query ke sheet
    for (const row of results) {
      // Menentukan path gambar berdasarkan npm
      const imagePathWebp = path.join(imageDir, `${row.npm}.webp`);
      const imagePathJpg = path.join(imageDir, `${row.npm}.jpg`);
      const tempImagePath = path.join(tempImageDir, `${row.npm}.jpg`);

      // Menambahkan baris ke worksheet
      const rowIndex = results.indexOf(row) + 2; // Karena header ada di baris pertama
      let imageAdded = false;

      let imageWidth = 100;
      let imageHeight = 100;

      // Mengonversi gambar dari webp ke jpg jika file gambar ada
      if (fs.existsSync(imagePathWebp)) {
        try {
          await sharp(imagePathWebp).toFile(tempImagePath);

          // Menentukan ukuran gambar
          const metadata = await sharp(tempImagePath).metadata();
          imageWidth = metadata.width;
          imageHeight = metadata.height;

          // Menambahkan gambar ke worksheet
          const imageId = workbook.addImage({
            filename: tempImagePath,
            extension: 'jpg'
          });

          sheet.addImage(imageId, {
            tl: { col: 2, row: rowIndex - 1 }, // Kolom dan baris untuk posisi gambar
            ext: { width: imageWidth, height: imageHeight } // Ukuran gambar
          });

          imageAdded = true;
        } catch (err) {
          console.error(`Error saat mengonversi gambar ${imagePathWebp} ke ${tempImagePath}:`, err);
        }
      }

      // Jika gambar tidak ditemukan atau tidak bisa dikonversi, periksa format jpg
      if (!imageAdded && fs.existsSync(imagePathJpg)) {
        try {
          // Menentukan ukuran gambar
          const metadata = await sharp(imagePathJpg).metadata();
          imageWidth = metadata.width;
          imageHeight = metadata.height;

          // Menambahkan gambar ke worksheet
          const imageId = workbook.addImage({
            filename: imagePathJpg,
            extension: 'jpg'
          });

          sheet.addImage(imageId, {
            tl: { col: 2, row: rowIndex - 1 }, // Kolom dan baris untuk posisi gambar
            ext: { width: imageWidth, height: imageHeight } // Ukuran gambar
          });

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

      // Menyesuaikan tinggi baris dengan tinggi gambar
      sheet.getRow(rowIndex).height = imageHeight ? imageHeight / 6 : 20; // Adjust divisor as needed
    }

    // Menentukan path direktori dan file
    const outputDir = path.join(__dirname, '../output');
    const now = new Date();
    const formattedDate = now.toISOString().replace(/:/g, '-'); // Format menjadi 'YYYY-MM-DDTHH-MM-SS'
    const fileName = `file-${formattedDate}.xlsx`;
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
    }
  });
}

fetchAsistenAndCreateExcel().catch(err => console.error('Error:', err));
