// =================================================================================
// PENGATURAN GLOBAL (GANTI DENGAN MILIK ANDA)
// =================================================================================
const SS_ID = '19LHveCO9M72lAio8mH5JkLNE2tvVfdoVjPDRyvVB-Gc'; // GANTI DENGAN ID SPREADSHEET UNTUK AMBIL DATA
const FOLDER_ID_RAPOR = "1OP_ggUCve6fm-v9o6KCED_NXsmPUPmsH"; // GANTI DENGAN ID FOLDER UNTUK CETAK RAPOR
const FOLDER_ID_IDENTITAS = "1lKJ1nFUUeuPhrvDtEe6qgzmU239EGiy-"; // GANTI DENGAN ID FOLDER UNTUK IDENTITAS
const TEMPLATE_FILE_NAME = "RaporTemplate";
const TEMPLATE_FILE_NAME_IDENTITAS = "IdentitasTemplate";
// =================================================================================
// GANTI ID DI BAWAH INI DENGAN ID FILE DARI GOOGLE DRIVE ANDA
const LOGO_ID = "1aVP8Es4udsnaFxY4CgC0JkhVR5gR30Pp";
const STEMPEL_ID = "1v06V1q98ZAhpUb3dcNFPNfQuxy9lvCmx";
const TTD_ID = "14SVxFYcBZU9SrqJlzYiiPsBV-PGnGcmM";
const STEMPELID_ID = "1Bo77ubGEi1S0ZYWiiiBv-6DCYGkh0bOL";
// =================================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Nilai & Rapor Madrasah');
}

/**
 * Mengambil data unik untuk mengisi dropdown filter.
 */
function getFilterData() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const guruSheet = ss.getSheetByName('Guru');

  if (!siswaSheet) {
    throw new Error("Sheet 'Siswa' tidak ditemukan!");
  }

  const dataSiswa = siswaSheet.getDataRange().getValues();
  const headersSiswa = dataSiswa[0];
  const colIndexKelas = headersSiswa.indexOf('Kelas');

  const kelasList = (dataSiswa.length > 1 && colIndexKelas !== -1)
    ? [...new Set(dataSiswa.slice(1).map(row => row[colIndexKelas]).filter(String))].sort()
    : [];

  const guruList = (guruSheet && guruSheet.getLastRow() > 1)
    ? [...new Set(guruSheet.getRange(2, 1, guruSheet.getLastRow() - 1, 1).getValues().flat().filter(String))]
    : [];

  return {
    kelasList: kelasList,
    semesterList: ['Ganjil', 'Genap'],
    tahunPelajaranList: ['2025/2026'],
    guruList: guruList
  };
}

/**
 * Mengambil data siswa, materi, dan nilai yang sudah ada berdasarkan filter.
 */
function getDataSiswaMateriAndNilai(kelas, semester, tahun, guru) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const materiSheet = ss.getSheetByName('Materi');
  const nilaiSheet = ss.getSheetByName(`Nilai ${kelas}`);

  if (!siswaSheet || !materiSheet) {
    throw new Error("Sheet 'Siswa' atau 'Materi' tidak ditemukan!");
  }

  const siswaDataAll = siswaSheet.getDataRange().getValues();
  const siswaHeaders = siswaDataAll[0];
  const siswaDataRows = siswaDataAll.slice(1);
  const siswaColIndices = {
    kelas: siswaHeaders.indexOf('Kelas'),
    noAbsen: siswaHeaders.indexOf('No. Absen'),
    namaSiswa: siswaHeaders.indexOf('Nama Siswa')
  };

  if (Object.values(siswaColIndices).some(index => index === -1)) {
    throw new Error("Kolom yang diperlukan ('Kelas', 'No. Absen', atau 'Nama Siswa') tidak ditemukan di sheet 'Siswa'.");
  }

  const siswaData = siswaDataRows
    .filter(row => String(row[siswaColIndices.kelas]).trim() === String(kelas).trim());

  const siswaList = siswaData.map(row => ({
    noAbsen: row[siswaColIndices.noAbsen],
    nama: row[siswaColIndices.namaSiswa]
  }));

  const materiData = materiSheet.getRange(2, 1, materiSheet.getLastRow() - 1, materiSheet.getLastColumn()).getValues();
  const foundMateri = materiData.find(r => String(r[0]).trim() === String(kelas).trim() && String(r[1]).trim() === String(semester).trim());
  const materiList = foundMateri ? foundMateri.slice(2).filter(m => m && m !== '-') : [];

  let nilaiMap = {};
  if (nilaiSheet && nilaiSheet.getLastRow() > 1) {
    const nilaiValues = nilaiSheet.getDataRange().getValues();
    const headers = nilaiValues[0];
    const dataRows = nilaiValues.slice(1);

    const colIndices = {
      semester: headers.indexOf('Semester'),
      tahun: headers.indexOf('Tahun Pelajaran'),
      noAbsen: headers.indexOf('No. Absen'),
      namaSiswa: headers.indexOf('Nama Siswa')
    };

    dataRows.forEach(row => {
      if (String(row[colIndices.semester]).trim() === String(semester).trim() && String(row[colIndices.tahun]).trim() === String(tahun).trim()) {
        const key = `${String(row[colIndices.noAbsen]).trim()}-${String(row[colIndices.namaSiswa]).trim()}`;
        if (!nilaiMap[key]) {
          nilaiMap[key] = {};
        }

        materiList.forEach(materi => {
          if (!nilaiMap[key][materi]) {
            nilaiMap[key][materi] = {};
          }
          const aspek = ['Kitab', 'Syafahi', 'Tulis'];
          aspek.forEach(as => {
            const nilai = row[headers.indexOf(`${materi} (${as})`)];
            nilaiMap[key][materi][as] = (nilai !== undefined && nilai !== '') ? nilai : null;
          });
        });
      }
    });
  }

  const finalData = siswaList.map(siswa => {
    const key = `${String(siswa.noAbsen).trim()}-${String(siswa.nama).trim()}`;
    return {
      noAbsen: siswa.noAbsen,
      nama: siswa.nama,
      nilai: nilaiMap[key] || {}
    };
  });

  return { siswa: finalData, materi: materiList };
}

/**
 * Menyimpan nilai materi dengan pendekatan batch update.
 */
function simpanSemuaNilai(kelas, semester, tahunPelajaran, guru, nilaiData) {
  if (!nilaiData || !Array.isArray(nilaiData) || nilaiData.length === 0) {
    return 'Data nilai tidak ditemukan atau kosong.';
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetName = `Nilai ${kelas}`;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const materiToUpdate = Object.keys(nilaiData[0].nilai)[0];
  const aspekNilai = ['Kitab', 'Syafahi', 'Tulis'];

  let headers = [];
  if (sheet.getLastRow() > 0) {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  if (!headers || headers.length === 0 || headers[0] === "") {
    headers = ["Guru", "Kelas", "Semester", "Tahun Pelajaran", "No. Absen", "Nama Siswa"];
  }
  const newMateriHeaders = aspekNilai.map(aspek => `${materiToUpdate} (${aspek})`);
  const finalHeaders = [...new Set([...headers, ...newMateriHeaders])];

  if (finalHeaders.length > headers.length) {
    sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
    headers = finalHeaders;
  }
  const headerMap = {};
  headers.forEach((h, i) => headerMap[h] = i);

  const dataRange = sheet.getDataRange();
  const existingData = dataRange.getValues();
  const existingDataHeaders = existingData.shift();

  const dataToUpdate = existingData.filter(row =>
    String(row[headerMap["Kelas"]]).trim() !== String(kelas).trim() ||
    String(row[headerMap["Semester"]]).trim() !== String(semester).trim() ||
    String(row[headerMap["Tahun Pelajaran"]]).trim() !== String(tahunPelajaran).trim()
  );

  const rowsToModify = existingData.filter(row =>
    String(row[headerMap["Kelas"]]).trim() === String(kelas).trim() &&
    String(row[headerMap["Semester"]]).trim() === String(semester).trim() &&
    String(row[headerMap["Tahun Pelajaran"]]).trim() === String(tahunPelajaran).trim()
  );

  const siswaMap = {};
  nilaiData.forEach(siswa => {
    const key = `${String(siswa.noAbsen).trim()}-${String(siswa.nama).trim()}`;
    siswaMap[key] = siswa;
  });

  const modifiedRows = rowsToModify.map(row => {
    const key = `${String(row[headerMap["No. Absen"]]).trim()}-${String(row[headerMap["Nama Siswa"]]).trim()}`;
    const siswaNilai = siswaMap[key];

    if (siswaNilai) {
      const nilaiMateri = siswaNilai.nilai[materiToUpdate] || {};

      row[headerMap["Guru"]] = guru;

      aspekNilai.forEach(aspek => {
        const header = `${materiToUpdate} (${aspek})`;
        const colIndex = headerMap[header];
        if (colIndex !== undefined) {
          row[colIndex] = nilaiMateri[aspek] || '';
        }
      });

      delete siswaMap[key];
      return row;
    }
    return row;
  });

  const newRows = Object.values(siswaMap).map(siswa => {
    const newRow = new Array(headers.length).fill('');
    const nilaiSiswa = siswa.nilai[materiToUpdate] || {};

    newRow[headerMap["Guru"]] = guru;
    newRow[headerMap["Kelas"]] = kelas;
    newRow[headerMap["Semester"]] = semester;
    newRow[headerMap["Tahun Pelajaran"]] = tahunPelajaran;
    newRow[headerMap["No. Absen"]] = siswa.noAbsen;
    newRow[headerMap["Nama Siswa"]] = siswa.nama;

    aspekNilai.forEach(aspek => {
      const header = `${materiToUpdate} (${aspek})`;
      const colIndex = headerMap[header];
      if (colIndex !== undefined) {
        newRow[colIndex] = nilaiSiswa[aspek] || '';
      }
    });

    return newRow;
  });

  const allNewData = dataToUpdate.concat(modifiedRows).concat(newRows);

  if (allNewData.length > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, allNewData.length, headers.length).setValues(allNewData);
  }

  const updatedCount = modifiedRows.length;
  const newCount = newRows.length;
  return `Berhasil menyimpan nilai. ${updatedCount} data siswa diperbarui, ${newCount} data siswa baru ditambahkan.`;
}

/**
 * Mengambil nama wali kelas berdasarkan nama kelas.
 */
function getWaliKelas(kelas) {
  const waliKelasMap = {
    'SHIFIR A 1': 'M. Hafidz',
    'SHIFIR A 2': 'Choirul Anam',
    'SHIFIR B 1': 'Masykur',
    'SHIFIR B 2': 'Rudi Hartanto',
    'I A': 'Muhammad Munib, S.Ag',
    'I B': 'M. Ikhwan',
    'II A': 'Nur Kolis, S.Pd',
    'II B': 'M. Faizin',
    'III': 'M. Sodiq Musthofa',
    'IV': 'Abdul Basyir',
    'V': 'Ahmad Sugeng',
    'VI': 'Moch. Sofyan Assauri, S.Ag',
    // Tambahkan wali kelas lainnya di sini
  };
  return waliKelasMap[kelas] || '_________________________';
}

function getFinalGradesPreview(kelas, semester, tahun) {
  if (!kelas || !semester || !tahun) {
    throw new Error("Parameter Kelas, Semester, atau Tahun tidak tidak boleh kosong.");
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const materiSheet = ss.getSheetByName('Materi');
  const nilaiSheet = ss.getSheetByName(`Nilai ${kelas}`);

  if (!siswaSheet || !materiSheet || !nilaiSheet) {
    throw new Error("Satu atau lebih sheet tidak ditemukan.");
  }

  const siswaDataAll = siswaSheet.getDataRange().getValues();
  const siswaHeaders = siswaDataAll[0];
  const siswaDataRows = siswaDataAll.slice(1);
  const siswaColIndices = {
    noAbsen: siswaHeaders.indexOf('No. Absen'),
    namaSiswa: siswaHeaders.indexOf('Nama Siswa'),
    kelas: siswaHeaders.indexOf('Kelas')
  };

  const siswaData = siswaDataRows
    .filter(row => String(row[siswaColIndices.kelas]).trim() === String(kelas).trim());

  const materiData = materiSheet.getRange(2, 1, materiSheet.getLastRow() - 1, materiSheet.getLastColumn()).getValues();
  const foundMateri = materiData.find(r => String(r[0]).trim() === String(kelas).trim() && String(r[1]).trim() === String(semester).trim());
  const materiList = foundMateri ? foundMateri.slice(2).filter(m => m && m !== '-') : [];

  const nilaiValues = nilaiSheet.getDataRange().getValues();
  const nilaiHeaders = nilaiValues[0];
  const dataRows = nilaiValues.slice(1);
  const nilaiColIndices = {};
  materiList.forEach(materi => {
    nilaiColIndices[materi] = {
      kitab: nilaiHeaders.indexOf(`${materi} (Kitab)`),
      syafahi: nilaiHeaders.indexOf(`${materi} (Syafahi)`),
      tulis: nilaiHeaders.indexOf(`${materi} (Tulis)`)
    };
  });

  const previewData = siswaData.map(siswa => {
    const noAbsen = siswa[siswaColIndices.noAbsen];
    const namaSiswa = siswa[siswaColIndices.namaSiswa];

    const siswaNilaiRows = dataRows.filter(r =>
      String(r[nilaiHeaders.indexOf('No. Absen')]).trim() === String(noAbsen).trim() &&
      String(r[nilaiHeaders.indexOf('Semester')]).trim() === String(semester).trim() &&
      String(r[nilaiHeaders.indexOf('Tahun Pelajaran')]).trim() === String(tahun).trim()
    );

    const nilaiAkhirSiswa = {
      noAbsen,
      namaSiswa
    };

    if (siswaNilaiRows.length > 0) {
      materiList.forEach(materi => {
        const kitabVal = siswaNilaiRows[0][nilaiColIndices[materi].kitab];
        const syafahiVal = siswaNilaiRows[0][nilaiColIndices[materi].syafahi];
        const tulisVal = siswaNilaiRows[0][nilaiColIndices[materi].tulis];

        const kitabNum = typeof kitabVal === 'number' ? kitabVal : 0;
        const syafahiNum = typeof syafahiVal === 'number' ? syafahiVal : 0;
        const tulisNum = typeof tulisVal === 'number' ? tulisVal : 0;

        const nilaiAkhir = Math.round(
          (kitabNum * 0.50) + (syafahiNum * 0.1667) + (tulisNum * 0.3333)
        );
        nilaiAkhirSiswa[materi] = nilaiAkhir;
      });
    } else {
      materiList.forEach(materi => {
        nilaiAkhirSiswa[materi] = 'N/A';
      });
    }

    return nilaiAkhirSiswa;
  });

  const previewHeaders = ['No. Absen', 'Nama Siswa', ...materiList];
  return {
    headers: previewHeaders,
    data: previewData
  };
}

/**
 * Membuat rapor siswa dalam format PDF dan menyimpannya di Google Drive.
 */
function getStudentsInClass(kelas) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const dataSheet = ss.getSheetByName('Siswa');

  if (!dataSheet) {
    throw new Error('Sheet "Siswa" tidak ditemukan.');
  }

  const data = dataSheet.getDataRange().getValues();
  const header = data[0];
  const kelasIndex = header.indexOf('Kelas');
  const namaSiswaIndex = header.indexOf('Nama Siswa');

  if (kelasIndex === -1 || namaSiswaIndex === -1) {
    throw new Error('Kolom "Kelas" atau "Nama Siswa" tidak ditemukan di sheet Siswa.');
  }

  const filteredSiswa = data.slice(1).filter(row => row[kelasIndex] == kelas);

  return filteredSiswa.map(row => {
    return {
      namaSiswa: row[namaSiswaIndex]
    };
  });
}


// =================================================================================
// FUNGSI UTAMA UNTUK MEMBUAT RAPOR
// =================================================================================
function createRaporPDF(kelas, semester, tahun, siswaList = null) {
  if (!kelas || !semester || !tahun) {
    throw new Error("Parameter Kelas, Semester, atau Tahun tidak boleh kosong.");
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const materiSheet = ss.getSheetByName('Materi');
  const nilaiSheet = ss.getSheetByName(`Nilai ${kelas}`);
  const ketidakhadiranSheet = ss.getSheetByName('Siswa');

  if (!siswaSheet || !materiSheet || !nilaiSheet || !ketidakhadiranSheet) {
    throw new Error("Satu atau lebih sheet tidak ditemukan. Pastikan sheet 'Siswa', 'Materi', dan 'Nilai " + kelas + "' ada.");
  }

  const siswaDataAll = siswaSheet.getDataRange().getValues();
  const siswaHeaders = siswaDataAll[0];
  const siswaDataRows = siswaDataAll.slice(1);
  const siswaColIndices = {
    kelas: siswaHeaders.indexOf('Kelas'),
    noAbsen: siswaHeaders.indexOf('No. Absen'),
    namaSiswa: siswaHeaders.indexOf('Nama Siswa'),
    sakit: siswaHeaders.indexOf('Sakit'),
    izin: siswaHeaders.indexOf('Izin'),
    alpha: siswaHeaders.indexOf('Alpha')
  };

  if (Object.values(siswaColIndices).some(index => index === -1)) {
    throw new Error("Kolom yang diperlukan ('Kelas', 'No. Absen', 'Nama Siswa', 'Sakit', 'Izin', atau 'Alpha') tidak ditemukan di sheet 'Siswa'.");
  }

  let filteredSiswaData = siswaDataRows
    .filter(row => String(row[siswaColIndices.kelas]).trim() === String(kelas).trim());

  if (siswaList && siswaList.length > 0) {
    filteredSiswaData = filteredSiswaData.filter(row => siswaList.includes(row[siswaColIndices.namaSiswa]));
  }

  const materiData = materiSheet.getRange(2, 1, materiSheet.getLastRow() - 1, materiSheet.getLastColumn()).getValues();
  const foundMateri = materiData.find(r => String(r[0]).trim() === String(kelas).trim() && String(r[1]).trim() === String(semester).trim());
  const materiList = foundMateri ? foundMateri.slice(2).filter(m => m && m !== '-') : [];

  const nilaiValues = nilaiSheet.getDataRange().getValues();
  const headers = nilaiValues[0];
  const dataRows = nilaiValues.slice(1);

  const urls = [];
  const namaWaliKelas = getWaliKelas(kelas);

  const parentFolder = DriveApp.getFolderById(FOLDER_ID_RAPOR);
  let folderKelas;
  const existingFolders = parentFolder.getFoldersByName(kelas);

  if (existingFolders.hasNext()) {
    folderKelas = existingFolders.next();
  } else {
    folderKelas = parentFolder.createFolder(kelas);
  }

  filteredSiswaData.forEach(siswa => {
    const noAbsen = siswa[siswaColIndices.noAbsen];
    const namaSiswa = siswa[siswaColIndices.namaSiswa];

    let raportContent = '';
    let nilaiRataRata = 0;
    let totalNilai = 0;
    let nomorUrut = 1;

    const siswaNilaiRows = dataRows.filter(r =>
      String(r[headers.indexOf('No. Absen')]).trim() === String(noAbsen).trim() &&
      String(r[headers.indexOf('Semester')]).trim() === String(semester).trim() &&
      String(r[headers.indexOf('Tahun Pelajaran')]).trim() === String(tahun).trim()
    );

    if (siswaNilaiRows.length > 0) {
      materiList.forEach(materi => {
        const kitabVal = siswaNilaiRows[0][headers.indexOf(`${materi} (Kitab)`)];
        const syafahiVal = siswaNilaiRows[0][headers.indexOf(`${materi} (Syafahi)`)];
        const tulisVal = siswaNilaiRows[0][headers.indexOf(`${materi} (Tulis)`)];

        const nilaiKitab = (kitabVal !== null && typeof kitabVal === 'number') ? kitabVal : 0;
        const nilaiSyafahi = (syafahiVal !== null && typeof syafahiVal === 'number') ? syafahiVal : 0;
        const nilaiTulis = (tulisVal !== null && typeof tulisVal === 'number') ? tulisVal : 0;

        const nilaiAkhir = Math.round(
          (nilaiKitab * 0.50) + (nilaiSyafahi * 0.1667) + (nilaiTulis * 0.3333)
        );
        totalNilai += nilaiAkhir;

        raportContent += `
          <tr>
            <td style="text-align: center;">${nomorUrut}</td>
            <td>${materi}</td>
            <td style="text-align: center;">${nilaiAkhir}</td>
            <td>${angkaToLatin(nilaiAkhir)}</td>
            <td>${predikatNilai(nilaiAkhir)}</td>
          </tr>
        `;
        nomorUrut++;
      });
      nilaiRataRata = (totalNilai / materiList.length).toFixed(2);
    } else {
      raportContent = `<tr><td colspan="4">Data nilai tidak ditemukan untuk siswa ini.</td></tr>`;
    }

    const sakit = siswa[siswaColIndices.sakit] || 0;
    const izin = siswa[siswaColIndices.izin] || 0;
    const alpha = siswa[siswaColIndices.alpha] || 0;

    const barcodeData = `RAPOR DIGITAL\nMadrasah Diniyah Tarbiyatul Falah\nNama Siswa: ${namaSiswa}\nSemester: ${semester}\nTahun Pelajaran: ${tahun}`;
    const barcodeUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodeURIComponent(barcodeData)}`;

    const template = HtmlService.createTemplateFromFile(TEMPLATE_FILE_NAME);

    template.namaSiswa = namaSiswa;
    template.kelas = kelas;
    template.semester = semester;
    template.tahun = tahun;
    template.raportContent = raportContent;
    template.nilaiRataRata = nilaiRataRata;
    template.sakit = sakit;
    template.izin = izin;
    template.alpha = alpha;
    template.waliKelas = namaWaliKelas;
    template.barcodeUrl = barcodeUrl;

    // --- KODE BARU UNTUK GAMBAR ---
    template.logo = getGambarAsBase64(LOGO_ID);
    template.stempel = getGambarAsBase64(STEMPEL_ID);
    template.ttd = getGambarAsBase64(TTD_ID);
    // --- AKHIR KODE BARU ---

    const htmlOutput = template.evaluate();
    const pdfBlob = htmlOutput.getAs('application/pdf');

    const file = folderKelas.createFile(pdfBlob).setName(`${kelas}_${noAbsen}_${namaSiswa}.pdf`);
    urls.push(file.getUrl());
  });

  return urls;
}

// FUNGSI PENDUKUNG UNTUK MENGUBAH GAMBAR MENJADI BASE64
function getGambarAsBase64(fileId) {
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return `data:${mimeType};base64,${base64Data}`;
  } catch (e) {
    Logger.log("Gagal mengambil gambar dengan ID " + fileId + ": " + e.message);
    return "";
  }
}

// Fungsi untuk mencetak rapor siswa tunggal
function createSingleRaporPDF(kelas, semester, tahun, namaSiswa) {
  return createRaporPDF(kelas, semester, tahun, [namaSiswa]);
}

// Fungsi untuk mencetak rapor semua siswa
function createAllRaporPDF(kelas, semester, tahun) {
  return createRaporPDF(kelas, semester, tahun, null);
}

// Fungsi untuk mengambil gambar stempel dari Google Drive
function getStempelAsBase64() {
  // GANTI INI dengan ID file gambar dari Google Drive Anda
  const fileId = "1p6sPXTk9PfyNmbkBdI4hmjRCWMw9FLXn"; 
  const imageUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;

  try {
    const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
    const base64Data = Utilities.base64Encode(imageBlob.getBytes());
    const mimeType = imageBlob.getContentType();
    return `data:${mimeType};base64,${base64Data}`;
  } catch (e) {
    Logger.log("Gagal mengambil gambar: " + e.message);
    return "";
  }
}

// Fungsi bantu untuk mendapatkan nama wali kelas (sesuaikan jika perlu)
function predikatNilai(nilai) {
  if (nilai >= 86) return 'Sangat Baik';
  if (nilai >= 80) return 'Baik';
  if (nilai >= 75) return 'Sedang';
  if (nilai >= 70) return 'Cukup';
  return 'Kurang';
}

function angkaToLatin(angka) {
  const bilangan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan'];
  if (typeof angka !== 'number') return '';
  const angkaBulat = Math.floor(angka);
  let hasilTerbilang = '';
  if (angkaBulat < 10) {
    hasilTerbilang = bilangan[angkaBulat];
  } else if (angkaBulat >= 10 && angkaBulat <= 19) {
    if (angkaBulat === 10) hasilTerbilang = 'sepuluh';
    else if (angkaBulat === 11) hasilTerbilang = 'sebelas';
    else hasilTerbilang = bilangan[angkaBulat % 10] + ' belas';
  } else if (angkaBulat >= 20 && angkaBulat <= 99) {
    if (angkaBulat % 10 === 0) hasilTerbilang = bilangan[Math.floor(angkaBulat / 10)] + ' puluh';
    else hasilTerbilang = bilangan[Math.floor(angkaBulat / 10)] + ' puluh ' + bilangan[angkaBulat % 10];
  } else if (angkaBulat === 100) {
    hasilTerbilang = 'seratus';
  }
  if (angkaBulat > 100 && angkaBulat < 200) hasilTerbilang = 'seratus ' + angkaToLatin(angkaBulat % 100);
  if (angkaBulat >= 200 && angkaBulat <= 999) hasilTerbilang = bilangan[Math.floor(angkaBulat / 100)] + ' ratus ' + angkaToLatin(angkaBulat % 100);
  if (hasilTerbilang.length > 0) {
    return hasilTerbilang.charAt(0).toUpperCase() + hasilTerbilang.slice(1);
  } else {
    return hasilTerbilang;
  }
}

/**
 * Mengambil data ketidakhadiran dari sheet "Absen" berdasarkan kelas.
 */
function getKetidakhadiran(kelas) {
  const spreadsheet = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = spreadsheet.getSheetByName('Absen'); // Perubahan di sini!

  if (!siswaSheet) {
    throw new Error("Sheet 'Absen' tidak ditemukan. Pastikan nama sheet Anda sudah benar.");
  }

  const data = siswaSheet.getDataRange().getValues();
  const header = data[0];

  const colIndices = {
    kelas: header.indexOf('Kelas'),
    nama: header.indexOf('Nama Siswa'),
    sakit: header.indexOf('Sakit'),
    izin: header.indexOf('Izin'),
    alpha: header.indexOf('Alpha')
  };

  if (Object.values(colIndices).some(index => index === -1)) {
    throw new Error("Kolom yang diperlukan tidak ditemukan di sheet 'Absen'. Pastikan ada kolom 'Kelas', 'Nama Siswa', 'Sakit', 'Izin', dan 'Alpha'.");
  }

  const siswaFilteredByKelas = data
    .slice(1)
    .filter(row => String(row[colIndices.kelas]).trim() === String(kelas).trim())
    .map(row => {
      return {
        nama: row[colIndices.nama],
        sakit: row[colIndices.sakit] || 0,
        izin: row[colIndices.izin] || 0,
        alpha: row[colIndices.alpha] || 0
      };
    });

  return siswaFilteredByKelas;
}

/**
 * Menyimpan data ketidakhadiran ke sheet "Absen" dengan batch update.
 */
function simpanKetidakhadiran(kelas, dataArray) {
  const spreadsheet = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = spreadsheet.getSheetByName('Absen'); // Perubahan di sini!

  if (!siswaSheet) {
    throw new Error("Sheet 'Absen' tidak ditemukan. Pastikan nama sheet Anda sudah benar.");
  }

  const data = siswaSheet.getDataRange().getValues();
  const header = data[0];

  const colIndices = {
    nama: header.indexOf('Nama Siswa'),
    sakit: header.indexOf('Sakit'),
    izin: header.indexOf('Izin'),
    alpha: header.indexOf('Alpha')
  };

  if (Object.values(colIndices).some(index => index === -1)) {
    throw new Error("Kolom 'Nama Siswa', 'Sakit', 'Izin', atau 'Alpha' tidak ditemukan di sheet 'Absen'.");
  }

  const dataToUpdateMap = {};
  dataArray.forEach(item => {
    dataToUpdateMap[item.nama] = {
      sakit: item.sakit || 0,
      izin: item.izin || 0,
      alpha: item.alpha || 0
    };
  });

  const valuesToSet = data.slice(1);

  for (let i = 0; i < valuesToSet.length; i++) {
    const row = valuesToSet[i];
    const namaSiswa = row[colIndices.nama];

    if (dataToUpdateMap[namaSiswa]) {
      const updatedData = dataToUpdateMap[namaSiswa];
      row[colIndices.sakit] = updatedData.sakit;
      row[colIndices.izin] = updatedData.izin;
      row[colIndices.alpha] = updatedData.alpha;
    }
  }

  siswaSheet.getRange(2, 1, valuesToSet.length, valuesToSet[0].length).setValues(valuesToSet);

  return `Berhasil menyimpan data ketidakhadiran untuk ${dataArray.length} siswa.`;
}

/**
 * Mengambil nama wali kelas berdasarkan nama kelas.
 */
function getWaliKelas(kelas) {
  const waliKelasMap = {
    'SHIFIR A 1': 'M. Hafidz',
    'SHIFIR A 2': 'Choirul Anam',
    'SHIFIR B 1': 'Masykur',
    'SHIFIR B 2': 'Rudi Hartanto',
    'I A': 'Muhammad Munib, S.Ag',
    'I B': 'M. Ikhwan',
    'II A': 'Nur Kolis, S.Pd',
    'II B': 'M. Faizin',
    'III': 'M. Sodiq Musthofa',
    'IV': 'Abdul Basyir',
    'V': 'Ahmad Sugeng',
    'VI': 'Moch. Sofyan Assauri, S.Ag',
  };
  return waliKelasMap[kelas] || '_________________________';
}

/**
 * Membuat kartu identitas siswa dalam format PDF dari template HTML.
 */
function createIdentitasPDF(kelas, semester, tahun, siswaList = null) {
  if (!kelas || !semester || !tahun) {
    throw new Error("Parameter Kelas, Semester, atau Tahun tidak boleh kosong.");
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  if (!siswaSheet) {
    throw new Error("Sheet 'Siswa' tidak ditemukan!");
  }

  const siswaDataAll = siswaSheet.getDataRange().getValues();
  const siswaHeaders = siswaDataAll[0];
  const siswaDataRows = siswaDataAll.slice(1);
  
  const siswaColIndices = {};
  siswaHeaders.forEach((header, index) => {
    siswaColIndices[header] = index;
  });

  let filteredSiswaData = siswaDataRows.filter(row => String(row[siswaColIndices['Kelas']]).trim() === String(kelas).trim());

  if (siswaList && siswaList.length > 0) {
    filteredSiswaData = filteredSiswaData.filter(row => siswaList.includes(row[siswaColIndices['Nama Siswa']]));
  }
  
  const parentFolder = DriveApp.getFolderById(FOLDER_ID_IDENTITAS);
  let folderIdentitas;
  const existingFolders = parentFolder.getFoldersByName('Identitas');

  if (existingFolders.hasNext()) {
    folderIdentitas = existingFolders.next();
  } else {
    folderIdentitas = parentFolder.createFolder('Identitas');
  }

  const urls = [];
  
  const template = HtmlService.createTemplateFromFile(TEMPLATE_FILE_NAME_IDENTITAS); 

  // --- KODE BARU UNTUK GAMBAR ---
  template.logo = getGambarAsBase64(LOGO_ID);
  template.ttd = getGambarAsBase64(TTD_ID);
  template.stempelid = getGambarAsBase64(STEMPELID_ID);
  // --- AKHIR KODE BARU ---

  // --- KODE BARU UNTUK FOLDER KELAS ---
  let folderKelas;
  const existingKelasFolders = folderIdentitas.getFoldersByName(kelas);
  if (existingKelasFolders.hasNext()) {
    folderKelas = existingKelasFolders.next();
  } else {
    folderKelas = folderIdentitas.createFolder(kelas);
  }
  // --- AKHIR KODE BARU UNTUK FOLDER KELAS ---

  filteredSiswaData.forEach(siswa => {
    const dataSiswa = {
      nama: siswa[siswaColIndices['Nama Siswa']] || '',
      noAbsen: siswa[siswaColIndices['No. Absen']] || '',
      nis: siswa[siswaColIndices['NIS']] || '',
      jenisKelamin: siswa[siswaColIndices['Jenis Kelamin']] || '',
      ttl: siswa[siswaColIndices['Tempat/Tanggal Lahir']] || '',
      alamat: siswa[siswaColIndices['Alamat']] || '',
      namaAyah: siswa[siswaColIndices['Nama Ayah']] || '',
      namaIbu: siswa[siswaColIndices['Nama Ibu']] || '',
      pekerjaanAyah: siswa[siswaColIndices['Pekerjaan Ayah']] || '',
      pekerjaanIbu: siswa[siswaColIndices['Pekerjaan Ibu']] || '',
      alamatOrtu: siswa[siswaColIndices['Alamat Orang Tua']] || ''
    };
    
    template.data = dataSiswa;
    const html = template.evaluate().getContent();
    
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
    // Mengubah lokasi penyimpanan file ke dalam folderKelas yang baru dibuat
    const pdfFile = folderKelas.createFile(blob).setName(`${kelas}_${dataSiswa.noAbsen}_${dataSiswa.nama}_Identitas.pdf`);
    urls.push(pdfFile.getUrl());
  });
  
  return urls;
}

function getLogoAsBase64() {
  const imageUrl = "https://i.postimg.cc/nXptQKrc/IMG-20250815-WA0072.jpg";
  const imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  const base64Data = Utilities.base64Encode(imageBlob.getBytes());
  const mimeType = imageBlob.getContentType();
  return `data:${mimeType};base64,${base64Data}`;
}

