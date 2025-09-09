tolong lanjutkan ini lagi
/*******************************************************
 * Code.gs — Rapor Madrasah (UI + REST Router)
 * Kompatibel dengan Index.html kamu
 *******************************************************/

// =====================================================
// PENGATURAN GLOBAL (SESUAI KEPUNYAAN ANDA)
// =====================================================
const SS_ID = '19LHveCO9M72lAio8mH5JkLNE2tvVfdoVjPDRyvVB-Gc'; // Spreadsheet nilai & data
const FOLDER_ID_RAPOR = "1OP_ggUCve6fm-v9o6KCED_NXsmPUPmsH";     // Folder output rapor
const FOLDER_ID_IDENTITAS = "1lKJ1nFUUeuPhrvDtEe6qgzmU239EGiy-"; // Folder output identitas

const TEMPLATE_FILE_NAME = "RaporTemplate";
const TEMPLATE_FILE_NAME_IDENTITAS = "IdentitasTemplate";

// Gambar dari Drive (base64)
const LOGO_ID = "1aVP8Es4udsnaFxY4CgC0JkhVR5gR30Pp";
const STEMPEL_ID = "1v06V1q98ZAhpUb3dcNFPNfQuxy9lvCmx";
const TTD_ID = "14SVxFYcBZU9SrqJlzYiiPsBV-PGnGcmM";
const STEMPELID_ID = "1Bo77ubGEi1S0ZYWiiiBv-6DCYGkh0bOL";

// =====================================================
// ENTRY UI
// =====================================================
function doGet(e) {
  // Jika tanpa action → tampilkan UI (Index.html)
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService
      .createHtmlOutputFromFile('Index')
      .setTitle('Nilai & Rapor Madrasah');
  }
  // Jika ada action → balas JSON
  try {
    const res = routeGet(e);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: res }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const res = routePost(e, body);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: res }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================
// ROUTER (GET/POST)
// =====================================================
// Contoh pemakaian dari frontend (fetch):
// GET  .../exec?action=filters
// GET  .../exec?action=class-data&kelas=I%20A&semester=Ganjil&tahun=2025/2026&guru=Nama%20Guru
// GET  .../exec?action=preview&kelas=I%20A&semester=Ganjil&tahun=2025/2026
// GET  .../exec?action=ketidakhadiran-get&kelas=I%20A
// GET  .../exec?action=siswa-by-kelas&kelas=I%20A
//
// POST .../exec?action=save-nilai                  body: { kelas, semester, tahun, guru, nilaiData }
// POST .../exec?action=rapor-single                body: { kelas, semester, tahun, namaSiswa }
// POST .../exec?action=rapor-all                   body: { kelas, semester, tahun }
// POST .../exec?action=ketidakhadiran-save         body: { kelas, dataArray: [{nama,sakit,izin,alpha}, ...] }
// POST .../exec?action=identitas-single            body: { kelas, semester, tahun, namaSiswa }
// POST .../exec?action=identitas-all               body: { kelas, semester, tahun }

function routeGet(e) {
  const p = e.parameter || {};
  const action = (p.action || '').toLowerCase();

  if (action === 'filters') {
    return getFilterData();
  }

  if (action === 'class-data') {
    const { kelas, semester, tahun, guru } = p;
    return getDataSiswaMateriAndNilai(required(kelas, 'kelas'), required(semester, 'semester'), required(tahun, 'tahun'), guru || '');
  }

  if (action === 'preview') {
    const { kelas, semester, tahun } = p;
    return getFinalGradesPreview(required(kelas, 'kelas'), required(semester, 'semester'), required(tahun, 'tahun'));
  }

  if (action === 'ketidakhadiran-get') {
    return getKetidakhadiran(required(p.kelas, 'kelas'));
  }

  if (action === 'siswa-by-kelas') {
    return getStudentsInClass(required(p.kelas, 'kelas'));
  }

  throw new Error('Unknown GET action.');
}

function routePost(e, body) {
  const p = e.parameter || {};
  const action = (p.action || '').toLowerCase();

  if (action === 'save-nilai') {
    const { kelas, semester, tahun, guru, nilaiData } = body;
    return simpanSemuaNilai(required(kelas, 'kelas'), required(semester, 'semester'), required(tahun, 'tahun'), required(guru, 'guru'), requiredArray(nilaiData, 'nilaiData'));
  }

  if (action === 'rapor-single') {
    const { kelas, semester, tahun, namaSiswa } = body;
    return createSingleRaporPDF(required(kelas, 'kelas'), required(semester, 'semester'), required(tahun, 'tahun'), required(namaSiswa, 'namaSiswa'));
  }

  if (action === 'rapor-all') {
    const { kelas, semester, tahun } = body;
    return createAllRaporPDF(required(kelas, 'kelas'), required(semester, 'semester'), required(tahun, 'tahun'));
  }

  if (action === 'ketidakhadiran-save') {
    const { kelas, dataArray } = body;
    return simpanKetidakhadiran(required(kelas, 'kelas'), requiredArray(dataArray, 'dataArray'));
  }

  if (action === 'identitas-single') {
    const { kelas, semester, tahun, namaSiswa } = body;
    return createIdentitasPDF(required(kelas, 'kelas'), required(semester, 'semester'), required(tahun, 'tahun'), [required(namaSiswa, 'namaSiswa')]);
  }

  if (action === 'identitas-all') {
    const { kelas, semester, tahun } = body;
    return createIdentitasPDF(required(kelas, 'kelas'), required(semester, 'semester'), required(tahun, 'tahun'), null);
  }

  throw new Error('Unknown POST action.');
}

function required(val, name) {
  if (val === undefined || val === null || String(val).trim() === '') throw new Error(`Parameter '${name}' wajib diisi.`);
  return val;
}
function requiredArray(arr, name) {
  if (!Array.isArray(arr)) throw new Error(`Parameter '${name}' harus array.`);
  return arr;
}

// =====================================================
// DATA FILTER (Dropdown awal UI)
// =====================================================
function getFilterData() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const guruSheet = ss.getSheetByName('Guru');

  if (!siswaSheet) throw new Error("Sheet 'Siswa' tidak ditemukan!");

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
    kelasList,
    semesterList: ['Ganjil', 'Genap'],
    tahunPelajaranList: ['2025/2026'],
    guruList
  };
}

// =====================================================
// AMBIL DATA SISWA, MATERI, & NILAI
// =====================================================
function getDataSiswaMateriAndNilai(kelas, semester, tahun, guru) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const materiSheet = ss.getSheetByName('Materi');
  const nilaiSheet = ss.getSheetByName(`Nilai ${kelas}`);

  if (!siswaSheet || !materiSheet) throw new Error("Sheet 'Siswa' atau 'Materi' tidak ditemukan!");

  const siswaDataAll = siswaSheet.getDataRange().getValues();
  const siswaHeaders = siswaDataAll[0];
  const siswaDataRows = siswaDataAll.slice(1);
  const idx = {
    kelas: siswaHeaders.indexOf('Kelas'),
    noAbsen: siswaHeaders.indexOf('No. Absen'),
    namaSiswa: siswaHeaders.indexOf('Nama Siswa')
  };
  if (Object.values(idx).some(i => i === -1)) {
    throw new Error("Kolom 'Kelas', 'No. Absen', 'Nama Siswa' wajib ada di sheet 'Siswa'.");
  }

  const siswaList = siswaDataRows
    .filter(r => String(r[idx.kelas]).trim() === String(kelas).trim())
    .map(r => ({ noAbsen: r[idx.noAbsen], nama: r[idx.namaSiswa] }));

  const materiData = materiSheet.getRange(2, 1, Math.max(0, materiSheet.getLastRow() - 1), materiSheet.getLastColumn()).getValues();
  const foundMateri = materiData.find(r => String(r[0]).trim() === String(kelas).trim() && String(r[1]).trim() === String(semester).trim());
  const materiList = foundMateri ? foundMateri.slice(2).filter(m => m && m !== '-') : [];

  let nilaiMap = {};
  if (nilaiSheet && nilaiSheet.getLastRow() > 1) {
    const nilaiValues = nilaiSheet.getDataRange().getValues();
    const headers = nilaiValues[0];
    const dataRows = nilaiValues.slice(1);

    const col = {
      semester: headers.indexOf('Semester'),
      tahun: headers.indexOf('Tahun Pelajaran'),
      noAbsen: headers.indexOf('No. Absen'),
      namaSiswa: headers.indexOf('Nama Siswa')
    };

    dataRows.forEach(row => {
      if (String(row[col.semester]).trim() === String(semester).trim() &&
          String(row[col.tahun]).trim() === String(tahun).trim()) {

        const key = `${String(row[col.noAbsen]).trim()}-${String(row[col.namaSiswa]).trim()}`;
        if (!nilaiMap[key]) nilaiMap[key] = {};

        materiList.forEach(materi => {
          if (!nilaiMap[key][materi]) nilaiMap[key][materi] = {};
          ['Kitab','Syafahi','Tulis'].forEach(as => {
            const v = row[headers.indexOf(`${materi} (${as})`)];
            nilaiMap[key][materi][as] = (v !== undefined && v !== '') ? v : null;
          });
        });
      }
    });
  }

  const finalData = siswaList.map(s => {
    const key = `${String(s.noAbsen).trim()}-${String(s.nama).trim()}`;
    return { noAbsen: s.noAbsen, nama: s.nama, nilai: nilaiMap[key] || {} };
  });

  return { siswa: finalData, materi: materiList };
}

// =====================================================
// SIMPAN NILAI (batch upsert)
// =====================================================
function simpanSemuaNilai(kelas, semester, tahunPelajaran, guru, nilaiData) {
  if (!Array.isArray(nilaiData) || nilaiData.length === 0) return 'Data nilai kosong.';

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetName = `Nilai ${kelas}`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const materiToUpdate = Object.keys(nilaiData[0].nilai)[0];
  const aspek = ['Kitab','Syafahi','Tulis'];

  let headers = [];
  if (sheet.getLastRow() > 0) {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  if (!headers || headers.length === 0 || headers[0] === "") {
    headers = ["Guru","Kelas","Semester","Tahun Pelajaran","No. Absen","Nama Siswa"];
  }

  const newMateriHeaders = aspek.map(a => `${materiToUpdate} (${a})`);
  const finalHeaders = [...new Set([...headers, ...newMateriHeaders])];

  if (finalHeaders.length > headers.length) {
    sheet.getRange(1,1,1,finalHeaders.length).setValues([finalHeaders]);
    headers = finalHeaders;
  }

  const col = {};
  headers.forEach((h,i)=> col[h]=i);

  const existing = sheet.getDataRange().getValues();
  const existingHeaders = existing.shift() || headers;

  const stayOther = existing.filter(r =>
    String(r[col["Kelas"]]).trim() !== String(kelas).trim() ||
    String(r[col["Semester"]]).trim() !== String(semester).trim() ||
    String(r[col["Tahun Pelajaran"]]).trim() !== String(tahunPelajaran).trim()
  );

  const targetRows = existing.filter(r =>
    String(r[col["Kelas"]]).trim() === String(kelas).trim() &&
    String(r[col["Semester"]]).trim() === String(semester).trim() &&
    String(r[col["Tahun Pelajaran"]]).trim() === String(tahunPelajaran).trim()
  );

  const siswaMap = {};
  nilaiData.forEach(s => {
    const key = `${String(s.noAbsen).trim()}-${String(s.nama).trim()}`;
    siswaMap[key] = s;
  });

  const modifiedRows = targetRows.map(r => {
    const key = `${String(r[col["No. Absen"]]).trim()}-${String(r[col["Nama Siswa"]]).trim()}`;
    const sNilai = siswaMap[key];
    if (sNilai) {
      const n = sNilai.nilai[materiToUpdate] || {};
      r[col["Guru"]] = guru;
      ['Kitab','Syafahi','Tulis'].forEach(a=>{
        const hdr = `${materiToUpdate} (${a})`;
        const idx = col[hdr];
        if (idx !== undefined) r[idx] = n[a] || '';
      });
      delete siswaMap[key];
      return r;
    }
    return r;
  });

  const newRows = Object.values(siswaMap).map(s=>{
    const row = new Array(headers.length).fill('');
    const n = s.nilai[materiToUpdate] || {};
    row[col["Guru"]] = guru;
    row[col["Kelas"]] = kelas;
    row[col["Semester"]] = semester;
    row[col["Tahun Pelajaran"]] = tahunPelajaran;
    row[col["No. Absen"]] = s.noAbsen;
    row[col["Nama Siswa"]] = s.nama;
    ['Kitab','Syafahi','Tulis'].forEach(a=>{
      const hdr = `${materiToUpdate} (${a})`;
      const idx = col[hdr];
      if (idx !== undefined) row[idx] = n[a] || '';
    });
    return row;
  });

  const all = stayOther.concat(modifiedRows).concat(newRows);

  if (all.length > 0) {
    sheet.clearContents();
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(2,1,all.length,headers.length).setValues(all);
  }

  return `Berhasil menyimpan nilai. ${modifiedRows.length} diperbarui, ${newRows.length} ditambahkan.`;
}

// =====================================================
// PREVIEW NILAI AKHIR
// =====================================================
function getFinalGradesPreview(kelas, semester, tahun) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const materiSheet = ss.getSheetByName('Materi');
  const nilaiSheet = ss.getSheetByName(`Nilai ${kelas}`);

  if (!siswaSheet || !materiSheet || !nilaiSheet) throw new Error("Sheet yang dibutuhkan tidak ditemukan.");

  const siswaDataAll = siswaSheet.getDataRange().getValues();
  const hs = siswaDataAll[0];
  const rows = siswaDataAll.slice(1);
  const idx = {
    noAbsen: hs.indexOf('No. Absen'),
    namaSiswa: hs.indexOf('Nama Siswa'),
    kelas: hs.indexOf('Kelas')
  };

  const siswaData = rows.filter(r => String(r[idx.kelas]).trim() === String(kelas).trim());

  const materiData = materiSheet.getRange(2,1,Math.max(0,materiSheet.getLastRow()-1),materiSheet.getLastColumn()).getValues();
  const found = materiData.find(r => String(r[0]).trim()===String(kelas).trim() && String(r[1]).trim()===String(semester).trim());
  const materiList = found ? found.slice(2).filter(m=>m && m!=='-') : [];

  const nilaiValues = nilaiSheet.getDataRange().getValues();
  const nh = nilaiValues[0];
  const drows = nilaiValues.slice(1);

  const col = {};
  materiList.forEach(m=>{
    col[m] = {
      kitab: nh.indexOf(`${m} (Kitab)`),
      syafahi: nh.indexOf(`${m} (Syafahi)`),
      tulis: nh.indexOf(`${m} (Tulis)`)
    };
  });

  const data = siswaData.map(s=>{
    const noAbsen = s[idx.noAbsen];
    const namaSiswa = s[idx.namaSiswa];
    const r = drows.filter(rr =>
      String(rr[nh.indexOf('No. Absen')]).trim()===String(noAbsen).trim() &&
      String(rr[nh.indexOf('Semester')]).trim()===String(semester).trim() &&
      String(rr[nh.indexOf('Tahun Pelajaran')]).trim()===String(tahun).trim()
    );
    const out = { noAbsen, namaSiswa };
    if (r.length>0) {
      materiList.forEach(m=>{
        const kv = r[0][col[m].kitab];
        const sv = r[0][col[m].syafahi];
        const tv = r[0][col[m].tulis];
        const k = typeof kv==='number'?kv:0;
        const sya = typeof sv==='number'?sv:0;
        const t = typeof tv==='number'?tv:0;
        const akhir = Math.round( (k*0.50) + (sya*0.1667) + (t*0.3333) );
        out[m] = akhir;
      });
    } else {
      materiList.forEach(m=> out[m] = 'N/A');
    }
    return out;
  });

  return { headers: ['No. Absen','Nama Siswa',...materiList], data };
}

// =====================================================
// CETAK RAPOR
// =====================================================
function createRaporPDF(kelas, semester, tahun, siswaList = null) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  const materiSheet = ss.getSheetByName('Materi');
  const nilaiSheet = ss.getSheetByName(`Nilai ${kelas}`);

  if (!siswaSheet || !materiSheet || !nilaiSheet) {
    throw new Error(`Sheet 'Siswa' / 'Materi' / 'Nilai ${kelas}' tidak ditemukan.`);
  }

  const siswaDataAll = siswaSheet.getDataRange().getValues();
  const hs = siswaDataAll[0];
  const rows = siswaDataAll.slice(1);
  const i = {
    kelas: hs.indexOf('Kelas'),
    noAbsen: hs.indexOf('No. Absen'),
    namaSiswa: hs.indexOf('Nama Siswa'),
    sakit: hs.indexOf('Sakit'),
    izin: hs.indexOf('Izin'),
    alpha: hs.indexOf('Alpha')
  };
  if (Object.values(i).some(v=>v===-1)) throw new Error("Kolom wajib ('Kelas','No. Absen','Nama Siswa','Sakit','Izin','Alpha') belum lengkap di 'Siswa'.");

  let filtered = rows.filter(r => String(r[i.kelas]).trim()===String(kelas).trim());
  if (siswaList && siswaList.length>0) {
    filtered = filtered.filter(r => siswaList.includes(r[i.namaSiswa]));
  }

  const materiData = materiSheet.getRange(2,1,Math.max(0,materiSheet.getLastRow()-1),materiSheet.getLastColumn()).getValues();
  const found = materiData.find(r => String(r[0]).trim()===String(kelas).trim() && String(r[1]).trim()===String(semester).trim());
  const materiList = found ? found.slice(2).filter(m => m && m!=='-') : [];

  const nilaiValues = nilaiSheet.getDataRange().getValues();
  const nh = nilaiValues[0];
  const drows = nilaiValues.slice(1);

  const urls = [];
  const namaWaliKelas = getWaliKelas(kelas);

  const parent = DriveApp.getFolderById(FOLDER_ID_RAPOR);
  let folderKelas;
  const it = parent.getFoldersByName(kelas);
  folderKelas = it.hasNext() ? it.next() : parent.createFolder(kelas);

  filtered.forEach(s=>{
    const noAbsen = s[i.noAbsen];
    const namaSiswa = s[i.namaSiswa];

    let raportContent = '';
    let total = 0, urut = 1;

    const match = drows.filter(r =>
      String(r[nh.indexOf('No. Absen')]).trim()===String(noAbsen).trim() &&
      String(r[nh.indexOf('Semester')]).trim()===String(semester).trim() &&
      String(r[nh.indexOf('Tahun Pelajaran')]).trim()===String(tahun).trim()
    );

    if (match.length>0) {
      materiList.forEach(m=>{
        const kv = match[0][nh.indexOf(`${m} (Kitab)`)];
        const sv = match[0][nh.indexOf(`${m} (Syafahi)`)];
        const tv = match[0][nh.indexOf(`${m} (Tulis)`)];
        const k = (kv !== null && typeof kv==='number')?kv:0;
        const sya = (sv !== null && typeof sv==='number')?sv:0;
        const t = (tv !== null && typeof tv==='number')?tv:0;
        const akhir = Math.round( (k*0.50) + (sya*0.1667) + (t*0.3333) );
        total += akhir;
        raportContent += `
          <tr>
            <td style="text-align:center;">${urut}</td>
            <td>${m}</td>
            <td style="text-align:center;">${akhir}</td>
            <td>${angkaToLatin(akhir)}</td>
            <td>${predikatNilai(akhir)}</td>
          </tr>
        `;
        urut++;
      });
    } else {
      raportContent = `<tr><td colspan="5" style="text-align:center;">Data nilai belum ada.</td></tr>`;
    }

    const rata = materiList.length ? (total / materiList.length).toFixed(2) : '0.00';

    const sakit = s[i.sakit] || 0;
    const izin = s[i.izin] || 0;
    const alpha = s[i.alpha] || 0;

    const barcodeData = `RAPOR DIGITAL\nMadrasah Diniyah Tarbiyatul Falah\nNama Siswa: ${namaSiswa}\nSemester: ${semester}\nTahun Pelajaran: ${tahun}`;
    const barcodeUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodeURIComponent(barcodeData)}`;

    const tmpl = HtmlService.createTemplateFromFile(TEMPLATE_FILE_NAME);
    tmpl.namaSiswa = namaSiswa;
    tmpl.kelas = kelas;
    tmpl.semester = semester;
    tmpl.tahun = tahun;
    tmpl.raportContent = raportContent;
    tmpl.nilaiRataRata = rata;
    tmpl.sakit = sakit;
    tmpl.izin = izin;
    tmpl.alpha = alpha;
    tmpl.waliKelas = namaWaliKelas;
    tmpl.barcodeUrl = barcodeUrl;

    // gambar
    tmpl.logo = getGambarAsBase64(LOGO_ID);
    tmpl.stempel = getGambarAsBase64(STEMPEL_ID);
    tmpl.ttd = getGambarAsBase64(TTD_ID);

    const htmlOutput = tmpl.evaluate();
    const pdfBlob = htmlOutput.getAs('application/pdf');
    const file = folderKelas.createFile(pdfBlob).setName(`${kelas}_${noAbsen}_${namaSiswa}.pdf`);
    urls.push(file.getUrl());
  });

  return urls;
}

function createSingleRaporPDF(kelas, semester, tahun, namaSiswa) {
  return createRaporPDF(kelas, semester, tahun, [namaSiswa]);
}
function createAllRaporPDF(kelas, semester, tahun) {
  return createRaporPDF(kelas, semester, tahun, null);
}

// =====================================================
// KETIDAKHADIRAN (sheet: Absen)
// =====================================================
function getKetidakhadiran(kelas) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Absen');
  if (!sheet) throw new Error("Sheet 'Absen' tidak ditemukan.");

  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const idx = {
    kelas: h.indexOf('Kelas'),
    nama: h.indexOf('Nama Siswa'),
    sakit: h.indexOf('Sakit'),
    izin: h.indexOf('Izin'),
    alpha: h.indexOf('Alpha')
  };
  if (Object.values(idx).some(v=>v===-1)) throw new Error("Kolom wajib tidak lengkap di sheet 'Absen'.");

  return data.slice(1)
    .filter(r => String(r[idx.kelas]).trim()===String(kelas).trim())
    .map(r => ({
      nama: r[idx.nama],
      sakit: r[idx.sakit] || 0,
      izin: r[idx.izin] || 0,
      alpha: r[idx.alpha] || 0
    }));
}

function simpanKetidakhadiran(kelas, dataArray) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Absen');
  if (!sheet) throw new Error("Sheet 'Absen' tidak ditemukan.");

  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const idx = {
    nama: h.indexOf('Nama Siswa'),
    sakit: h.indexOf('Sakit'),
    izin: h.indexOf('Izin'),
    alpha: h.indexOf('Alpha')
  };
    if (Object.values(idx).some(v => v === -1)) throw new Error("Kolom wajib tidak lengkap di sheet 'Absen'.");

  // Buat map dari nama siswa → indeks baris
  const nameToRow = {};
  for (let i = 1; i < data.length; i++) {
    const nama = String(data[i][idx.nama]).trim();
    if (nama) nameToRow[nama] = i + 1; // baris di Sheet
  }

  dataArray.forEach(d => {
    const row = nameToRow[String(d.nama).trim()];
    if (row) {
      sheet.getRange(row, idx.sakit + 1).setValue(d.sakit || 0);
      sheet.getRange(row, idx.izin + 1).setValue(d.izin || 0);
      sheet.getRange(row, idx.alpha + 1).setValue(d.alpha || 0);
    }
  });

  return `Ketidakhadiran untuk kelas ${kelas} berhasil disimpan.`;
}

// =====================================================
// IDENTITAS SISWA (Cetak PDF identitas per siswa)
// =====================================================
function createIdentitasPDF(kelas, semester, tahun, siswaList = null) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const siswaSheet = ss.getSheetByName('Siswa');
  if (!siswaSheet) throw new Error("Sheet 'Siswa' tidak ditemukan.");

  const siswaData = siswaSheet.getDataRange().getValues();
  const hs = siswaData[0];
  const rows = siswaData.slice(1);

  const idx = {
    kelas: hs.indexOf('Kelas'),
    noAbsen: hs.indexOf('No. Absen'),
    namaSiswa: hs.indexOf('Nama Siswa'),
    tempatLahir: hs.indexOf('Tempat Lahir'),
    tanggalLahir: hs.indexOf('Tanggal Lahir'),
    alamat: hs.indexOf('Alamat'),
    namaWali: hs.indexOf('Nama Wali'),
    noHP: hs.indexOf('No HP')
  };

  if (Object.values(idx).some(v => v === -1)) {
    throw new Error("Kolom identitas wajib belum lengkap di sheet 'Siswa'.");
  }

  let filtered = rows.filter(r => String(r[idx.kelas]).trim() === String(kelas).trim());
  if (siswaList && siswaList.length > 0) {
    filtered = filtered.filter(r => siswaList.includes(r[idx.namaSiswa]));
  }

  const urls = [];
  const parent = DriveApp.getFolderById(FOLDER_ID_IDENTITAS);
  let folderKelas;
  const it = parent.getFoldersByName(kelas);
  folderKelas = it.hasNext() ? it.next() : parent.createFolder(kelas);

  filtered.forEach(s => {
    const tmpl = HtmlService.createTemplateFromFile(TEMPLATE_FILE_NAME_IDENTITAS);
    tmpl.namaSiswa = s[idx.namaSiswa];
    tmpl.kelas = kelas;
    tmpl.semester = semester;
    tmpl.tahun = tahun;
    tmpl.tempatLahir = s[idx.tempatLahir] || '-';
    tmpl.tanggalLahir = s[idx.tanggalLahir] || '-';
    tmpl.alamat = s[idx.alamat] || '-';
    tmpl.namaWali = s[idx.namaWali] || '-';
    tmpl.noHP = s[idx.noHP] || '-';
    tmpl.logo = getGambarAsBase64(LOGO_ID);
    tmpl.stempel = getGambarAsBase64(STEMPELID_ID);
    tmpl.ttd = getGambarAsBase64(TTD_ID);

    const htmlOutput = tmpl.evaluate();
    const pdfBlob = htmlOutput.getAs('application/pdf');
    const file = folderKelas.createFile(pdfBlob).setName(`Identitas_${s[idx.noAbsen]}_${s[idx.namaSiswa]}.pdf`);
    urls.push(file.getUrl());
  });

  return urls;
}

// =====================================================
// FUNGSI BANTUAN
// =====================================================
function getWaliKelas(kelas) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const guruSheet = ss.getSheetByName('Guru');
  if (!guruSheet) return '-';

  const data = guruSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === kelas) {
      return data[i][0]; // kolom pertama = nama guru
    }
  }
  return '-';
}

function getGambarAsBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const contentType = blob.getContentType();
    const encoded = Utilities.base64Encode(blob.getBytes());
    return `data:${contentType};base64,${encoded}`;
  } catch (err) {
    return '';
  }
}

function angkaToLatin(n) {
  const angka = parseInt(n, 10);
  if (isNaN(angka)) return '-';
  const huruf = ['Nol', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan'];
  return String(angka).split('').map(d => huruf[parseInt(d)]).join(' ');
}

function predikatNilai(nilai) {
  if (nilai >= 90) return 'A';
  if (nilai >= 80) return 'B';
  if (nilai >= 70) return 'C';
  return 'D';
}

