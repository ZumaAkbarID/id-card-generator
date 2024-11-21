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

  const query = `
    SELECT
      a.name,
      a.npm,
      r.role,
      r.deskripsi,
      a.mhs_foto
    FROM
      asistens a
      JOIN role_manage rm ON a.id = rm.id_asisten
      JOIN role r ON rm.id_role = r.id
    WHERE
      rm.id_role BETWEEN 21 AND 48
    ORDER BY
      rm.id_role DESC
  `;

  connection.query(query, async (err, results) => {
    if (err) {
      console.error('Error saat menjalankan query:', err);
      connection.end();
      return;
    }
    console.info("QUERY SUCCESS");

    console.info("CREATING WORKBOOK");
    // Buat workbook dan sheet Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Asisten');

    // Tambahkan header
    sheet.columns = [
      { header: 'No', key: 'no', width: 10 },
      { header: 'Nama', key: 'name', width: 30 },
      { header: 'NPM', key: 'npm', width: 20 },
      { header: 'Role', key: 'role', width: 30 },
      { header: 'Deskripsi', key: 'deskripsi', width: 50 }
    ];

    // Path direktori output
    const outputDir = path.join(__dirname, '../output');
    const photoDir = path.join(outputDir, 'photos');
    const qrDir = path.join(outputDir, 'qrs');

    const QR_URL = process.env.VALIDATE_QR_FA_URL;

    // Membuat direktori jika belum ada
    [outputDir, photoDir, qrDir].forEach(dir => {
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
    });

    console.info(`PROCESSING ${results.length} DATA`);
    let no = 1;
    for (const row of results) {
      const photoPath = path.join(photoDir, `${no}.jpg`);
      const qrPath = path.join(qrDir, `${no}.jpg`);

      // Proses foto
      if (row.mhs_foto) {
        const inputPhotoPath1 = path.join(__dirname, 'image', row.npm + '.jpg');
        const inputPhotoPath2 = path.join(__dirname, 'image', row.npm + '.webp');
        let inputPhotoPath = "";

        if (fs.existsSync(inputPhotoPath1)) {
          inputPhotoPath = inputPhotoPath1;
        } else if (fs.existsSync(inputPhotoPath2)) {
          inputPhotoPath = inputPhotoPath2;
        }

        if (inputPhotoPath !== "") {
          const ext = path.extname(inputPhotoPath).toLowerCase();
          if (ext === '.webp') {
            // Konversi dari .webp ke .jpg
            try {
              await sharp(inputPhotoPath).toFile(photoPath);
            } catch (err) {
              console.error(`Gagal mengonversi ${inputPhotoPath}:`, err);
            }
          } else if (ext === '.jpg') {
            fs.copyFileSync(inputPhotoPath, photoPath);
          }
        } else {
          console.warn(`File foto tidak ditemukan: ${inputPhotoPath}`);
        }
      }

      // Generate QR Code
      try {
        const qrContent = `${QR_URL}${row.npm}`;
        await QRCode.toFile(qrPath, qrContent);
      } catch (err) {
        console.error(`Gagal membuat QR untuk ${row.npm}:`, err);
      }

      // Tambahkan data ke Excel
      sheet.addRow({
        no,
        name: row.name,
        npm: row.npm,
        role: row.role,
        deskripsi: row.deskripsi
      });

      no++;
    }

    // Menentukan nama file Excel
    const now = new Date();
    const formattedDate = now.toISOString().replace(/:/g, '-');
    const fileName = `Asisten-${formattedDate}.xlsx`;
    const filePath = path.join(outputDir, fileName);

    console.info("WRITING EXCEL FILE");
    try {
      await workbook.xlsx.writeFile(filePath);
      console.log(`File Excel berhasil dibuat! Nama file: ${fileName}`);
    } catch (err) {
      console.error('Error saat menyimpan file Excel:', err);
    } finally {
      connection.end();
    }
  });
}

fetchAsistenAndCreateExcel().catch(err => console.error('Error:', err));