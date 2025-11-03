function loadSetting(key) {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty(key) || '';
}
function saveSetting(key, val) {
  PropertiesService.getDocumentProperties().setProperty(key, String(val||''));
  return 'OK';
}

// ================== LOGIN & SESSION ==================
function loginUser(creds) {
  const u = String(creds?.username || '').trim();
  const p = String(creds?.password || '');
  if (!u || !p) throw new Error('Username/Password wajib.');

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sh) throw new Error('Sheet Users tidak ditemukan.');
  const vals = sh.getDataRange().getValues(); // header + data

  for (let i = 1; i < vals.length; i++) {
    const [username, password, role, kelas, active] = vals[i];
    if (String(username || '').trim().toLowerCase() === u.toLowerCase()) {
      if (String(active).toLowerCase() === 'false') throw new Error('Akun non-aktif.');
      if (String(password) !== p) throw new Error('Password salah.');
      const token = Utilities.getUuid();
      const payload = { username: username, role: String(role||'TU'), kelas: String(kelas||''), ts: Date.now() };
      CacheService.getDocumentCache().put('sess_' + token, JSON.stringify(payload), 60 * 60 * 8); // 8 jam
      return { token, username, role: payload.role, kelas: payload.kelas };
    }
  }
  throw new Error('Username tidak ditemukan.');
}

function validateSession(token) {
  if (!token) return null;
  const raw = CacheService.getDocumentCache().get('sess_' + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e){ return null; }
}

function logoutToken(token) {
  if (!token) return 'OK';
  CacheService.getDocumentCache().put('sess_' + token, '', 1);
  return 'OK';
}

// ======== AUTH HELPERS (GUARD) =========
function requireAuth_(token) {
  const sess = validateSession(token);
  if (!sess) throw new Error('Unauthorized: silakan login.');
  return sess; // { username, role, kelas, ts }
}
function enforceClassForWali_(sess, kelasInput) {
  // Wali Kelas hanya boleh kelas miliknya; TU bebas
  return (sess.role === 'Wali Kelas') ? (sess.kelas || '') : (kelasInput || '');
}
function requireRoleTU_(sess) {
  if (sess.role !== 'TU') throw new Error('Forbidden: hanya TU.');
}

// =================== ROUTER ===================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('Aplikasi Absensi Siswa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function loadPageContent(pageName, token) {
  const sess = requireAuth_(token);
  try {
    if (pageName === 'Dashboard')        return HtmlService.createHtmlOutputFromFile('Dashboard').getContent();
    if (pageName === 'FormAbsensi')      return HtmlService.createHtmlOutputFromFile('FormAbsensi').getContent();
    if (pageName === 'LaporanHarian')    return HtmlService.createHtmlOutputFromFile('LaporanHarian').getContent();
    if (pageName === 'LaporanBulanan')   return HtmlService.createHtmlOutputFromFile('LaporanBulanan').getContent();
    if (pageName === 'Pelanggaran')      return HtmlService.createHtmlOutputFromFile('Pelanggaran').getContent();
    if (pageName === 'Pengaturan')       { requireRoleTU_(sess); return HtmlService.createHtmlOutputFromFile('Pengaturan').getContent(); }
    if (pageName === 'EditKehadiran')    { requireRoleTU_(sess); return HtmlService.createHtmlOutputFromFile('EditKehadiran').getContent(); }
    if (pageName === 'DataSiswa')        { requireRoleTU_(sess); return HtmlService.createHtmlOutputFromFile('DataSiswa').getContent(); }
    return '<h3>Halaman Belum Tersedia</h3>';
  } catch (e) {
    return `<h3>Error: ${e.message}</h3>`;
  }
}

// ================= DASHBOARD =================
function getDashboardData(kelasParam, token) {
  const sess = requireAuth_(token);
  const kelas = enforceClassForWali_(sess, kelasParam);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abs = ss.getSheetByName('Absensi');
  if (!abs || abs.getLastRow() < 2)
    return { today: {Alfa:0,Izin:0,Sakit:0}, thisMonth:{Alfa:0,Izin:0,Sakit:0}, topSiswa:[], topKelas:[] };

  const last = abs.getLastRow();
  const vals = abs.getRange(2, 1, last-1, 9).getValues();
  const disp = abs.getRange(2, 1, last-1, 9).getDisplayValues();

  const now = new Date();
  const todayKey = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const startMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  const today = [];
  const month = [];

  for (let i=0;i<vals.length;i++){
    const r = vals[i], d = disp[i];
    const tgl = r[1] instanceof Date ? r[1] : new Date(r[1]);
    const key = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const row = {
      id:String(r[0]||''), tanggal:tgl, tanggalStr:key, nisn:String(d[2]||''), nama:String(r[3]||''),
      kelas:String(r[4]||''), status:String(r[5]||''), ket:String(r[6]||''), petugas:String(r[8]||'')
    };

    if (kelas && kelas!=='Semua' && row.kelas !== kelas) continue;

    if (key===todayKey) today.push(row);
    if (tgl >= startMonth) month.push(row);
  }

  const count = (arr)=>arr.reduce((a,c)=>{ a[c.status]=(a[c.status]||0)+1; return a; }, {Alfa:0,Izin:0,Sakit:0});
  const todayCounts = count(today);
  const monthCounts = count(month);

  // ranking
  const byKey = (arr, k)=> {
    const m={}; arr.forEach(r=>{ const kk=r[k]; if (!m[kk]) m[kk]={ total:0, nama:r.nama, kelas:r.kelas }; m[kk].total++; });
    return Object.values(m).sort((a,b)=>b.total-a.total).slice(0,5);
  };
  const topSiswa = byKey(month, 'nisn'); // nama, kelas ikut disimpan di map
  const topKelas = byKey(month, 'kelas').map(x=>({ total:x.total, kelas:x.kelas||'' }));

  return { today: todayCounts, thisMonth: monthCounts, topSiswa, topKelas };
}

function getTrend30Hari(kelasParam, token){
  const sess = requireAuth_(token);
  const kelas = enforceClassForWali_(sess, kelasParam);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abs = ss.getSheetByName('Absensi');
  if (!abs || abs.getLastRow()<2) return [];

  const days = {};
  const now = new Date();
  for(let i=29;i>=0;i--){
    const d=new Date(now.getFullYear(), now.getMonth(), now.getDate()-i);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    days[key] = { tanggal: key, Alfa:0, Izin:0, Sakit:0 };
  }

  const vals = abs.getRange(2,1,abs.getLastRow()-1,9).getValues();
  const disp = abs.getRange(2,1,abs.getLastRow()-1,9).getDisplayValues();
  for (let i=0;i<vals.length;i++){
    const r=vals[i], d=disp[i];
    const tgl = r[1] instanceof Date ? r[1] : new Date(r[1]);
    const key = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const kls = String(r[4]||'');
    const st  = String(r[5]||'');
    if (!days[key]) continue;
    if (kelas && kelas!=='Semua' && kls!==kelas) continue;
    if (st==='Alfa'||st==='Izin'||st==='Sakit') days[key][st]++;
  }

  return Object.values(days);
}

// =============== FORM ABSENSI ===============
function getDaftarKelas(token) {
  const sess = requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Kelas');
  if (!sh || sh.getLastRow()<2) return [];
  const arr = sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().filter(Boolean);
  // Wali Kelas hanya kelasnya
  if (sess.role==='Wali Kelas') return arr.filter(x=>x===sess.kelas);
  return arr;
}

function getSiswaByKelas(namaKelas, token) {
  const sess = requireAuth_(token);
  const kelas = enforceClassForWali_(sess, namaKelas);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Siswa');
  if (!sh || sh.getLastRow()<2) return [];
  const vals = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
  const disp = sh.getRange(2,1,sh.getLastRow()-1,5).getDisplayValues();
  const out = [];
  for (let i=0;i<vals.length;i++){
    const v=vals[i], d=disp[i];
    if (kelas && v[2]!==kelas) continue; // kol C = kelas
    out.push([d[0], v[1]]); // [NISN display (nol depan aman), Nama]
  }
  return out;
}

function simpanAbsensi(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const absensiSheet = ss.getSheetByName('Absensi');

    // Paksa NISN jadi TEKS agar nol di depan tidak hilang
    // Tambahkan apostrof ' di depan -> Sheets akan menyimpan sebagai teks
    const nisnText = "'" + String(formData.nisn_siswa || '').trim();

    // ID + tanggal sebagai Date object (agar bisa difilter)
    const id_absensi = 'ABS-' + Date.now();
    const tanggal = new Date(formData.tanggal);
    const email_petugas = Session.getActiveUser().getEmail();

    const newRow = [
      id_absensi,            // A: ID
      tanggal,               // B: Tanggal (Date)
      nisnText,              // C: NISN (Text, aman nol depan)
      String(formData.nama_siswa || '').trim(), // D: Nama
      String(formData.kelas || '').trim(),      // E: Kelas
      String(formData.status || '').trim(),     // F: Status (Alfa/Izin/Sakit)
      String(formData.keterangan || '').trim(), // G: Keterangan
      '',                     // H: (mis. bukti foto; dibiarkan kosong)
      email_petugas           // I: Petugas
    ];

    absensiSheet.appendRow(newRow);

    // OPTIONAL: set number format untuk kolom tanggal -> tanggal Indonesia
    const lastRow = absensiSheet.getLastRow();
    absensiSheet.getRange(lastRow, 2).setNumberFormat('yyyy-mm-dd');

    return 'Data absensi untuk ' + formData.nama_siswa + ' berhasil disimpan!';
  } catch (e) {
    return 'Terjadi kesalahan: ' + e.toString();
  }
}

// ============= LAPORAN HARIAN =============
function getLaporanHarian(params, token){
  const sess = requireAuth_(token);
  params = params || {};
  const kelas = enforceClassForWali_(sess, params.kelas);
  const tglStr = String(params.tanggal || '').slice(0,10);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Absensi');
  if (!sh || sh.getLastRow()<2) return { tanggal:tglStr, kelas, rows:[], counts:{Alfa:0,Izin:0,Sakit:0} };

  const vals = sh.getRange(2,1,sh.getLastRow()-1,9).getValues();
  const disp = sh.getRange(2,1,sh.getLastRow()-1,9).getDisplayValues();
  const out = [];
  const counts = { Alfa:0, Izin:0, Sakit:0 };

  for (let i=0;i<vals.length;i++){
    const r=vals[i], d=disp[i];
    const tgl = r[1] instanceof Date ? r[1] : new Date(r[1]);
    const key = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (key !== tglStr) continue;
    if (kelas && kelas!=='Semua' && String(r[4]||'') !== kelas) continue;

    const row = {
      id:String(r[0]||''), tanggal:key,
      nisn:String(d[2]||''), nama:String(r[3]||''), kelas:String(r[4]||''),
      status:String(r[5]||''), keterangan:String(r[6]||''), petugas:String(r[8]||'')
    };
    out.push(row);
    if (counts[row.status] != null) counts[row.status]++;
  }

  return { tanggal:tglStr, kelas: (kelas||'Semua'), rows: out, counts };
}

function exportLaporanHarianPDF(params, token){
  const sess = requireAuth_(token);
  params = params || {};
  params.kelas = enforceClassForWali_(sess, params.kelas);

  const rep = getLaporanHarian(params, token);
  const tplId = loadSetting('TEMPLATE_LAP_HARIAN_DOC_ID');
  if (!tplId) throw new Error('Doc ID template harian kosong. Isi di Pengaturan.');

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const fileName = `Laporan Harian ${rep.tanggal}${rep.kelas && rep.kelas!=='Semua' ? ' - ' + rep.kelas : ''}`;

  const tmp = DriveApp.getFileById(tplId).makeCopy(fileName);
  const doc = DocumentApp.openById(tmp.getId());
  const body = doc.getBody();

  body.replaceText('{{TANGGAL}}', String(rep.tanggal));
  body.replaceText('{{KELAS}}',   String(rep.kelas));
  body.replaceText('{{TOTAL_ALFA}}',  String(rep.counts.Alfa));
  body.replaceText('{{TOTAL_IZIN}}',  String(rep.counts.Izin));
  body.replaceText('{{TOTAL_SAKIT}}', String(rep.counts.Sakit));
  body.replaceText('{{WAKTU_CETAK}}', nowStr);

  const anchor = body.findText('{{TABEL_DETAIL}}');
  if (anchor) {
    const el = anchor.getElement();
    el.asText().setText('');
    const headers = ['#','NISN','Nama','Kelas','Status','Keterangan','Petugas'];
    const table = body.insertTable(el.getParent().getChildIndex(el)+1, [headers]);
    rep.rows.forEach((r,i)=>{
      table.appendTableRow([ String(i+1), r.nisn, r.nama, r.kelas, r.status, r.keterangan, r.petugas ]);
    });
    table.setBorderWidth(0.5);
  }

  doc.saveAndClose();
  const pdf = DriveApp.getFileById(tmp.getId()).getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdf).setName(fileName + '.pdf');
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return pdfFile.getUrl();
}

// ========== LAPORAN BULANAN (PAGED) ==========
function getDataBulananPaged(params, token) {
  const sess = requireAuth_(token);
  params = params || {};
  const bln  = Number(params.bulan || 0);
  const thn  = Number(params.tahun || 0);
  const kelasFilter = enforceClassForWali_(sess, params.kelas);
  const nisnFilter  = String(params.nisn || '').trim();
  const offset = Number(params.offset || 0);
  const maxRows = Number(params.maxRows || 150);

  if (!bln || !thn) return { rows:[], total:0, truncated:false, nextOffset:null, summary:{Alfa:0,Izin:0,Sakit:0}, bulan:bln, tahun:thn, kelas:(kelasFilter||'Semua') };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Absensi');
  if (!sh || sh.getLastRow() < 2)
    return { rows:[], total:0, truncated:false, nextOffset:null, summary:{Alfa:0,Izin:0,Sakit:0}, bulan:bln, tahun:thn, kelas:(kelasFilter||'Semua') };

  const vals = sh.getRange(2,1, sh.getLastRow()-1, 9).getValues();
  const disp = sh.getRange(2,1, sh.getLastRow()-1, 9).getDisplayValues();

  const firstDay = new Date(thn, bln-1, 1);
  const nextMonth= new Date(thn, bln, 1);

  const all=[];
  const summary = { Alfa:0, Izin:0, Sakit:0 };
  for (let i=0;i<vals.length;i++){
    const v=vals[i], d=disp[i];
    const tgl = v[1] instanceof Date ? v[1] : new Date(v[1]);
    if (!(tgl>=firstDay && tgl<nextMonth)) continue;
    const kelas = String(v[4]||'');
    const status= String(v[5]||'');
    const nisn  = String(d[2]||'');
    if (kelasFilter && kelasFilter!=='Semua' && kelas !== kelasFilter) continue;
    if (nisnFilter && nisn !== nisnFilter) continue;

    all.push({
      id:String(v[0]||''),
      tanggal: tgl,
      nisn,
      nama: String(v[3]||''),
      kelas,
      status,
      keterangan: String(v[6]||''),
      petugas: String(v[8]||'')
    });
    if (summary[status]!=null) summary[status]++;
  }

  all.sort((a,b)=> a.tanggal - b.tanggal || String(a.nama).localeCompare(String(b.nama)));
  const total = all.length;
  const end = Math.min(offset+maxRows, total);
  const page = all.slice(offset, end).map(r=>({
    ...r,
    tanggal: Utilities.formatDate(r.tanggal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  }));

  return { rows:page, total, truncated:end<total, nextOffset: end<total ? end : null, summary, bulan:bln, tahun:thn, kelas: (kelasFilter||'Semua') };
}

// ========== EXPORT BULANAN (TEMPLATE) ==========
function exportLaporanBulananRekapPDF(params, token) {
  const sess = requireAuth_(token);
  params = params || {};
  params.kelas = enforceClassForWali_(sess, params.kelas);
  const rep = (function(){
    const paged = getDataBulananPaged({...params, offset:0, maxRows:50000}, token);
    return { bulan:paged.bulan, tahun:paged.tahun, kelas:paged.kelas, summary:paged.summary, rows:paged.rows };
  })();

  const tplId = loadSetting('TEMPLATE_LAP_BULANAN_REKAP_DOC_ID');
  if (!tplId) throw new Error('Doc ID template rekap bulanan kosong. Isi di Pengaturan.');

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const fileName = `Rekap Bulanan ${rep.bulan}-${rep.tahun}${rep.kelas && rep.kelas!=='Semua' ? ' - ' + rep.kelas : ''}`;

  const tmp = DriveApp.getFileById(tplId).makeCopy(fileName);
  const doc = DocumentApp.openById(tmp.getId());
  const body = doc.getBody();

  body.replaceText('{{BULAN}}', String(rep.bulan));
  body.replaceText('{{TAHUN}}', String(rep.tahun));
  body.replaceText('{{KELAS}}', String(rep.kelas));
  body.replaceText('{{TOTAL_ALFA}}', String(rep.summary.Alfa));
  body.replaceText('{{TOTAL_IZIN}}', String(rep.summary.Izin));
  body.replaceText('{{TOTAL_SAKIT}}', String(rep.summary.Sakit));
  body.replaceText('{{WAKTU_CETAK}}', nowStr);

  const anchor = body.findText('{{TABEL_DETAIL}}');
  if (anchor) {
    const el = anchor.getElement();
    el.asText().setText('');
    const headers = ['#','Tanggal','NISN','Nama','Kelas','Status','Keterangan','Petugas'];
    const t = body.insertTable(el.getParent().getChildIndex(el)+1, [headers]);
    rep.rows.forEach((r,i)=>{
      t.appendTableRow([ String(i+1), r.tanggal, r.nisn, r.nama, r.kelas, r.status, r.keterangan, r.petugas ]);
    });
    t.setBorderWidth(0.5);
  }

  doc.saveAndClose();
  const pdf = DriveApp.getFileById(tmp.getId()).getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdf).setName(fileName + '.pdf');
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return pdfFile.getUrl();
}

function exportLaporanBulananPerSiswaAll(params, token){
  const sess = requireAuth_(token);
  params = params || {};
  params.kelas = enforceClassForWali_(sess, params.kelas);

  const all = getDataBulananPaged({...params, offset:0, maxRows:50000}, token);
  const tplId = loadSetting('TEMPLATE_LAP_BULANAN_PERSISWA_DOC_ID');
  if (!tplId) throw new Error('Doc ID template per siswa kosong. Isi di Pengaturan.');

  // group by NISN
  const byNisn = {};
  all.rows.forEach(r=>{
    if (!byNisn[r.nisn]) byNisn[r.nisn] = [];
    byNisn[r.nisn].push(r);
  });

  const folder = DriveApp.createFolder(`Bulanan Per Siswa ${all.bulan}-${all.tahun}${all.kelas && all.kelas!=='Semua' ? ' - ' + all.kelas : ''}`);
  Object.keys(byNisn).forEach(nisn=>{
    const rows = byNisn[nisn];
    const nama = (rows[0]?.nama)||'';
    const kelas= (rows[0]?.kelas)||'';
    const fileName = `${nisn} - ${nama}`;

    const tmp = DriveApp.getFileById(tplId).makeCopy(fileName, folder);
    const doc = DocumentApp.openById(tmp.getId());
    const body = doc.getBody();

    body.replaceText('{{BULAN}}', String(all.bulan));
    body.replaceText('{{TAHUN}}', String(all.tahun));
    body.replaceText('{{KELAS}}', String(kelas));
    body.replaceText('{{NISN}}', String(nisn));
    body.replaceText('{{NAMA}}', String(nama));

    const anchor = body.findText('{{TABEL_DETAIL}}');
    if (anchor){
      const el = anchor.getElement(); el.asText().setText('');
      const headers = ['#','Tanggal','Status','Keterangan','Petugas'];
      const t = body.insertTable(el.getParent().getChildIndex(el)+1, [headers]);
      rows.forEach((r,i)=>{
        t.appendTableRow([ String(i+1), r.tanggal, r.status, r.keterangan, r.petugas ]);
      });
      t.setBorderWidth(0.5);
    }

    doc.saveAndClose();
    const pdf = DriveApp.getFileById(tmp.getId()).getAs('application/pdf');
    const pdfFile = folder.createFile(pdf).setName(fileName + '.pdf');
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  });

  return folder.getUrl();
}

function exportLaporanBulananPerSiswaOne(params, token){
  const sess = requireAuth_(token);
  params = params || {};
  params.kelas = enforceClassForWali_(sess, params.kelas);

  const all = getDataBulananPaged({...params, offset:0, maxRows:50000}, token);
  const tplId = loadSetting('TEMPLATE_LAP_BULANAN_PERSISWA_DOC_ID');
  if (!tplId) throw new Error('Doc ID template per siswa kosong. Isi di Pengaturan.');
  const rows = all.rows.filter(r=> String(r.nisn) === String(params.nisn));
  const nama = (rows[0]?.nama)||'';
  const kelas= (rows[0]?.kelas)||'';
  const fileName = `${params.nisn} - ${nama} (${all.bulan}-${all.tahun})`;

  const tmp = DriveApp.getFileById(tplId).makeCopy(fileName);
  const doc = DocumentApp.openById(tmp.getId());
  const body = doc.getBody();

  body.replaceText('{{BULAN}}', String(all.bulan));
  body.replaceText('{{TAHUN}}', String(all.tahun));
  body.replaceText('{{KELAS}}', String(kelas||all.kelas));
  body.replaceText('{{NISN}}', String(params.nisn));
  body.replaceText('{{NAMA}}', String(nama));

  const anchor = body.findText('{{TABEL_DETAIL}}');
  if (anchor){
    const el = anchor.getElement(); el.asText().setText('');
    const headers = ['#','Tanggal','Status','Keterangan','Petugas'];
    const t = body.insertTable(el.getParent().getChildIndex(el)+1, [headers]);
    rows.forEach((r,i)=>{
      t.appendTableRow([ String(i+1), r.tanggal, r.status, r.keterangan, r.petugas ]);
    });
    t.setBorderWidth(0.5);
  }

  doc.saveAndClose();
  const pdf = DriveApp.getFileById(tmp.getId()).getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdf).setName(fileName + '.pdf');
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return pdfFile.getUrl();
}

// =============== PELANGGARAN =================
function getLaporanPelanggaran(params, token){
  const sess = requireAuth_(token);
  params = params || {};
  const bulan = Number(params.bulan||0);
  const tahun = Number(params.tahun||0);
  const kelasFilter = enforceClassForWali_(sess, params.kelas);
  if (!bulan || !tahun) return { bulan, tahun, kelas:(kelasFilter||'Semua'), rows:[] };

  const data = getDataBulananPaged({ bulan, tahun, kelas: kelasFilter, offset:0, maxRows:50000 }, token).rows;
  // hitung streak berturut-turut per status
  const byNisn = {};
  data.forEach(r=>{
    if (!byNisn[r.nisn]) byNisn[r.nisn] = { nama:r.nama, kelas:r.kelas, days:{} };
    if (!byNisn[r.nisn].days[r.tanggal]) byNisn[r.nisn].days[r.tanggal] = r.status;
  });

  const rows=[];
  Object.keys(byNisn).forEach(nisn=>{
    const info = byNisn[nisn];
    const keys = Object.keys(info.days).sort(); // urut tanggal
    let sAlfa=0,sIzin=0,sSakit=0, maxA=0,maxI=0,maxS=0;
    let prev='';

    keys.forEach(k=>{
      const st = info.days[k];
      if (st==='Alfa'){ sAlfa = (prev==='Alfa'? sAlfa+1:1); maxA=Math.max(maxA,sAlfa); } else sAlfa=0;
      if (st==='Izin'){ sIzin = (prev==='Izin'? sIzin+1:1); maxI=Math.max(maxI,sIzin); } else sIzin=0;
      if (st==='Sakit'){ sSakit = (prev==='Sakit'? sSakit+1:1); maxS=Math.max(maxS,sSakit); } else sSakit=0;
      prev = st;
    });

    if (maxA>3 || maxI>5 || maxS>4) {
      rows.push({
        nisn, nama:info.nama, kelas:info.kelas,
        alfa:maxA, izin:maxI, sakit:maxS
      });
    }
  });

  return { bulan, tahun, kelas: (kelasFilter||'Semua'), rows };
}

function exportPelanggaranPDF(payload, token){
  const sess = requireAuth_(token);
  payload = payload || {};
  const kelas = enforceClassForWali_(sess, payload.kelas);

  const jenis = String(payload.jenis||'Alfa'); // Alfa|Izin|Sakit
  const tplKey = jenis==='Alfa' ? 'TEMPLATE_PELANGGARAN_ALFA_DOC_ID'
              : jenis==='Izin' ? 'TEMPLATE_PELANGGARAN_IZIN_DOC_ID'
              : 'TEMPLATE_PELANGGARAN_SAKIT_DOC_ID';
  const tplId = loadSetting(tplKey);
  if (!tplId) throw new Error('Doc ID template pelanggaran untuk ' + jenis + ' kosong. Isi di Pengaturan.');

  // ambil data bulanan (untuk siswa tsb)
  const all = getDataBulananPaged({ bulan:Number(payload.bulan), tahun:Number(payload.tahun), kelas, offset:0, maxRows:50000 }, token);
  const rows = all.rows.filter(r => String(r.nisn) === String(payload.nisn) && r.status===jenis);
  const nama = (rows[0]?.nama)||'';
  const fileName = `Pelanggaran ${jenis} - ${payload.nisn} - ${nama} (${all.bulan}-${all.tahun})`;

  const tmp = DriveApp.getFileById(tplId).makeCopy(fileName);
  const doc = DocumentApp.openById(tmp.getId());
  const body = doc.getBody();

  body.replaceText('{{BULAN}}', String(all.bulan));
  body.replaceText('{{TAHUN}}', String(all.tahun));
  body.replaceText('{{KELAS}}', String(rows[0]?.kelas || all.kelas));
  body.replaceText('{{NISN}}', String(payload.nisn));
  body.replaceText('{{NAMA}}', String(nama));
  body.replaceText('{{STREAK}}', String(payload.streak||rows.length||0));

  const anchor = body.findText('{{TABEL_DETAIL}}');
  if (anchor){
    const el = anchor.getElement(); el.asText().setText('');
    const headers = ['#','Tanggal','Keterangan','Petugas'];
    const t = body.insertTable(el.getParent().getChildIndex(el)+1, [headers]);
    rows.forEach((r,i)=>{
      t.appendTableRow([ String(i+1), r.tanggal, r.keterangan, r.petugas ]);
    });
    t.setBorderWidth(0.5);
  }

  doc.saveAndClose();
  const pdf = DriveApp.getFileById(tmp.getId()).getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdf).setName(fileName + '.pdf');
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return pdfFile.getUrl();
}

// =============== WHATSAPP ===============
function sendWhatsApp(payload, token){
  const sess = requireAuth_(token);
  const url = loadSetting('WA_API_ENDPOINT');
  const tkn = loadSetting('WA_API_TOKEN');
  if (!url || !tkn) throw new Error('WA API belum diset.');

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ to: payload.to, message: payload.message }),
    headers: { Authorization: 'Bearer ' + tkn },
    muteHttpExceptions: true
  });
  return res.getResponseCode() + ' ' + res.getContentText();
}

function getHpByNisn(nisn, token){
  const sess = requireAuth_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Siswa');
  if (!sh || sh.getLastRow()<2) return '';
  const disp = sh.getRange(2,1, sh.getLastRow()-1, 5).getDisplayValues();
  for (let i=0;i<disp.length;i++){
    if (disp[i][0] === String(nisn)) return String(disp[i][4]||''); // kol E nohp
  }
  return '';
}

// ============== EDIT KEHADIRAN (TU) ==============
function getAbsensiPage(params, token){
  const sess = requireAuth_(token); requireRoleTU_(sess);
  params = params || {};
  const kelasFilter = String(params.kelas||'');
  const offset = Number(params.offset||0);
  const maxRows = Number(params.maxRows||100);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Absensi');
  if (!sh || sh.getLastRow()<2) return { rows:[], total:0, truncated:false, nextOffset:null };

  const rng = sh.getRange(2,1, sh.getLastRow()-1, 9);
  const vals= rng.getValues();
  const disp= rng.getDisplayValues();

  const all=[];
  for (let i=0;i<vals.length;i++){
    const v=vals[i], d=disp[i];
    if (kelasFilter && String(v[4]||'') !== kelasFilter) continue;
    all.push({
      id:String(v[0]||''),
      tanggal:(v[1] instanceof Date ? v[1] : new Date(v[1])),
      nisn:String(d[2]||''), nama:String(v[3]||''), kelas:String(v[4]||''),
      status:String(v[5]||''), keterangan:String(v[6]||''), petugas:String(v[8]||'')
    });
  }
  all.sort((a,b)=> b.tanggal - a.tanggal);

  const total = all.length;
  const end = Math.min(offset+maxRows, total);
  const page = all.slice(offset, end).map(r=>({
    ...r,
    tanggal: Utilities.formatDate(r.tanggal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  }));

  return { rows: page, total, truncated:end<total, nextOffset:end<total? end : null };
}

function updateAbsensi(payload, token){
  const sess = requireAuth_(token); requireRoleTU_(sess);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Absensi');
  if (!sh || sh.getLastRow() < 2) throw new Error('Data kosong.');

  const rng = sh.getRange(2,1, sh.getLastRow()-1, 9);
  const vals= rng.getValues();

  let rowNum=-1;
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]) === String(payload.id_absensi)) {
      rowNum = i+2; break;
    }
  }
  if (rowNum<0) throw new Error('ID tidak ditemukan.');

  if (payload.tanggal) sh.getRange(rowNum,2).setValue(new Date(payload.tanggal)); // kol B
  if (payload.status) sh.getRange(rowNum,6).setValue(payload.status);            // kol F
  sh.getRange(rowNum,7).setValue(payload.keterangan || '');                      // kol G
  return 'Perubahan disimpan.';
}

// ================== CRUD SISWA (TU) ==================
/**
 * Sheet "Siswa": A=NISN, B=Nama, C=Kelas, D=Alamat, E=NoHp
 * Semua operasi memakai format text utk kol A agar nol depan aman
 */

function getSiswaPage(params, token){
  const sess = requireAuth_(token); requireRoleTU_(sess);
  params = params || {};
  const kelas = String(params.kelas||'');
  const q     = String(params.q||'').toLowerCase();
  const offset= Number(params.offset||0);
  const maxRows= Number(params.maxRows||100);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Siswa');
  if (!sh || sh.getLastRow() < 2) return { rows:[], total:0, truncated:false, nextOffset:null };

  const rng  = sh.getRange(2,1, sh.getLastRow()-1, 5);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues(); // cek NISN dengan nol depan aman

  const out=[];
  for (let i=0;i<vals.length;i++){
    const v=vals[i], d=disp[i];
    const rec = {
      nisn:String(d[0]||''), nama:String(v[1]||''), kelas:String(v[2]||''), alamat:String(v[3]||''), nohp:String(v[4]||'')
    };
    if (kelas && rec.kelas!==kelas) continue;
    if (q && !(`${rec.nisn} ${rec.nama} ${rec.kelas} ${rec.nohp}`.toLowerCase().includes(q))) continue;
    out.push(rec);
  }

  const total = out.length;
  const end = Math.min(offset+maxRows, total);
  return { rows: out.slice(offset, end), total, truncated:end<total, nextOffset:end<total? end : null };
}

function upsertSiswa(payload, token){
  const sess = requireAuth_(token); requireRoleTU_(sess);
  let originalNisn = String(payload?.originalNisn || '').trim();
  let nisn  = String(payload?.nisn || '').trim();
  const nama  = String(payload?.nama || '').trim();
  const kelas = String(payload?.kelas || '').trim();
  const alamat= String(payload?.alamat || '').trim();
  const nohp  = String(payload?.nohp || '').trim();
  if (!nisn || !/^\d+$/.test(nisn)) throw new Error('NISN wajib angka.');
  if (!nama) throw new Error('Nama wajib diisi.');
  if (!kelas) throw new Error('Kelas wajib diisi.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Siswa');
  if (!sh) throw new Error('Sheet Siswa tidak ditemukan.');

  const last = sh.getLastRow();
  if (last < 2) {
    sh.getRange(2,1).setNumberFormat('@');
    sh.appendRow(["'"+nisn, nama, kelas, alamat, nohp]);
    return 'Ditambahkan.';
  }

  const rng  = sh.getRange(2,1, last-1, 5);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();

  // cari baris target
  let targetIndex = -1;
  if (originalNisn) {
    for (let i = 0; i < disp.length; i++) if (disp[i][0] === originalNisn) { targetIndex = i; break; }
  } else {
    for (let i = 0; i < disp.length; i++) if (disp[i][0] === nisn) { targetIndex = i; break; }
  }

  // duplikasi nisn ketika ganti
  if (targetIndex >= 0 && originalNisn && originalNisn !== nisn) {
    for (let i = 0; i < disp.length; i++) {
      if (i === targetIndex) continue;
      if (disp[i][0] === nisn) throw new Error('NISN baru sudah ada di data.');
    }
  }

  if (targetIndex < 0) {
    // tambah baru
    for (let i = 0; i < disp.length; i++) if (disp[i][0] === nisn) throw new Error('NISN sudah ada.');
    const newRow = last + 1;
    sh.getRange(newRow,1).setNumberFormat('@');
    sh.getRange(newRow,1,1,5).setValues([["'"+nisn, nama, kelas, alamat, nohp]]);
    return 'Ditambahkan.';
  } else {
    // update
    const rowNum = targetIndex + 2;
    sh.getRange(rowNum,1).setNumberFormat('@');
    sh.getRange(rowNum,1,1,5).setValues([["'"+nisn, nama, kelas, alamat, nohp]]);
    return 'Diperbarui.';
  }
}

function deleteSiswaByNisn(nisn, force, token){
  const sess = requireAuth_(token); requireRoleTU_(sess);
  nisn = String(nisn || '').trim();
  if (!nisn) throw new Error('NISN kosong.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // referensi di Absensi?
  const abs = ss.getSheetByName('Absensi');
  if (abs && abs.getLastRow() >= 2 && !force) {
    const rngA = abs.getRange(2, 3, abs.getLastRow() - 1, 1).getDisplayValues(); // kol C (NISN)
    for (let i = 0; i < rngA.length; i++) if (rngA[i][0] === nisn) {
      throw new Error('Tidak bisa hapus: NISN dipakai di Absensi. Gunakan hapus dengan konfirmasi (force).');
    }
  }

  const sh = ss.getSheetByName('Siswa');
  if (!sh || sh.getLastRow() < 2) throw new Error('Data siswa kosong.');
  const rng  = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues(); // NISN
  for (let i = 0; i < rng.length; i++) {
    if (rng[i][0] === nisn) {
      sh.deleteRow(i + 2);
      return 'Dihapus.';
    }
  }
  throw new Error('NISN tidak ditemukan.');
}
