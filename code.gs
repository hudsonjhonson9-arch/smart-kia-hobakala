const SPREADSHEET_ID = '';

const SHEETS = {
  pasien: ['No RM','Tanggal','Nama Ibu','Nama Suami','NIK','Umur','Alamat/Desa','No HP'],
  anc: [
    'No RM','Tanggal Kunjungan','Nama Ibu','Nama Suami','NIK','Umur','Alamat','No HP','DJJ',
    'Gravida (G,P,A)','HPHT','HPL','UK (Minggu)','Kunjungan','BB (Kg)','TB (Cm)',
    'TD (mmHg)','Nadi (x/menit)','Suhu (C)','Respirasi (x/menit)','Lila (Cm)','Leopold',
    'HB','Syphilis','HIV','HbSAg','GolDA','GDS',
    'Keluhan','Tindakan/Terapi','Keterangan','Rujuk','Bidan Pemeriksa','Dokter Pemeriksa'
  ],
  persalinan: [
    'No RM','Tanggal Persalinan','Nama Ibu','Nama Suami','NIK','Umur','Alamat','No HP',
    'Usia Kehamilan','Jenis Persalinan','Tempat Persalinan','Waktu Persalinan',
    'Penolong Bidan','Penolong Dokter','Jenis Kelamin Bayi',
    'Berat Bayi (gram)','Panjang Bayi (cm)','Keadaan Bayi','APGAR Skor',
    'Komplikasi','Rujuk','Tindakan'
  ],
  nifas: [
    'No RM','Nama Ibu','Nama Suami','NIK','Umur','Alamat','No HP',
    'Tanggal Persalinan','Tanggal Kunjungan Nifas','Kunjungan Nifas Ke',
    'Keadaan Umum','Kesadaran','TD (mmHg)','Nadi (x/menit)','Suhu (C)','RR (menit)',
    'TFU','Kontraksi Uterus','Lochia','Perinium/Luka','ASI','Eliminasi','Nyeri','Keluhan',
    'Tindakan/Terapi','Bidan Pemeriksa','Dokter Pemeriksa','Rujuk','Edukasi','Keterangan'
  ],
  bayi: [
    'No RM','Nama Bayi','Jenis Kelamin','Tanggal Lahir','Anak Ke','Kelahiran Ke',
    'Berat Lahir (gram)','PB Saat Lahir (cm)','Nama Ibu','Nama Ayah','Alamat/Desa','No HP',
    'Keadaan Bayi','Kesadaran','Keluhan','Nadi (x/menit)','Suhu (C)','RR (x/menit)',
    'BB Saat Ini (gram)','PB Saat Ini (cm)','LK (cm)','LP (cm)','LD (cm)',
    'APGAR Skor','IMD','HB0','SHK Vena','SHK Tumit','Inj Vitamin K','Salep Mata',
    'Tindakan/Terapi','Bidan Pemeriksa','Dokter Pemeriksa','Rujuk','Keterangan Tambahan'
  ],
  kb: [
    'No RM','Nama Ibu','Nama Suami','NIK','Umur','Alamat/Desa','No HP',
    'Tanggal Pelayanan','Metode KB','Status Kepesertaan','Riwayat KB Sebelumnya','Keluhan KB',
    'Keadaan Umum','Kesadaran','TD (mmHg)','Nadi (x/menit)','Suhu (C)','RR (menit)','BB (Kg)','TB (Cm)',
    'Tindakan/Terapi','Bidan Pemeriksa','Dokter Pemeriksa','Rujuk','Edukasi','Keterangan'
  ],
  monitoring: [
    'Bulan','Tahun','Total KIA','Total KB','Data Lengkap','Data Belum Lengkap',
    'Data Ibu','Data Bayi','Keterangan'
  ]
};

const USERS_SHEET = 'Users';

function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (!action) {
      setupSheets_();
      return ContentService.createTextOutput(JSON.stringify({ok:true,message:'SMART KIA API is running'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var args = [];
    for (var i = 0; i < 20; i++) {
      var val = e.parameter['a' + i];
      if (val === undefined) break;
      try { val = JSON.parse(val); } catch(ignored) {}
      args.push(val);
    }
    var fn = {
      loginUser: function(){return loginUser(args[0],args[1]);},
      saveData: function(){return saveData(args[0],args[1],args[2]);},
      editData: function(){return editData(args[0],args[1],args[2],args[3]);},
      deleteData: function(){return deleteData(args[0],args[1],args[2]);},
      getData: function(){return getData(args[0]);},
      getDashboard: function(){return getDashboard();},
      searchPatient: function(){return searchPatient(args[0]);},
      getActivityLog: function(){return getActivityLog();},
      getMonitoringData: function(){return getMonitoringData();},
      getKelengkapanData: function(){return getKelengkapanData();},
      changePassword: function(){return changePassword(args[0],args[1],args[2]);},
      getUsers: function(){return getUsers();},
      addUser: function(){return addUser(args[0],args[1],args[2],args[3]);},
      deleteUser: function(){return deleteUser(args[0]);},
      exportToExcel: function(){return exportToExcel(args[0]);},
      exportToPDF: function(){return exportToPDF(args[0]);},
      getBarcodeData: function(){return getBarcodeData(args[0]);},
      getModuleSummary: function(){return getModuleSummary();},
      getMonthlyStats: function(){return getMonthlyStats();},
      getFullDashboard: function(){return getFullDashboard();}
    };
    if (!fn[action]) throw new Error('Action tidak dikenal: ' + action);
    var result = fn[action]();
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, message:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSS_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan.');
  return ss;
}

function setupSheets_() {
  const ss = getSS_();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['ID', 'Waktu Input'].concat(SHEETS[name]));
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,SHEETS[name].length+2)
        .setFontWeight('bold')
        .setBackground('#ffd8e5');
    } else {
      // ponytail: header row must match SHEETS order exactly, otherwise data
      // (written in SHEETS order) sits under the wrong labels. Rewrite in place.
      const needed = ['ID','Waktu Input'].concat(SHEETS[name]);
      const lastCol = sh.getLastColumn();
      const existing = sh.getRange(1,1,1,lastCol).getValues()[0];
      const aligned = lastCol === needed.length && needed.every((h,i) => String(existing[i]) === h);
      if (!aligned) {
        sh.getRange(1,1,1,needed.length).setValues([needed]);
        sh.getRange(1,1,1,needed.length).setFontWeight('bold').setBackground('#ffd8e5');
        if (lastCol > needed.length) sh.getRange(1, needed.length+1, 1, lastCol-needed.length).clearContent();
      }
    }
  });
  let log = ss.getSheetByName('Aktivitas');
  if (!log) {
    log = ss.insertSheet('Aktivitas');
    log.appendRow(['Waktu','Modul','Keterangan','User']);
    log.setFrozenRows(1);
    log.getRange(1,1,1,4).setFontWeight('bold').setBackground('#ffd8e5');
  }
  let users = ss.getSheetByName(USERS_SHEET);
  if (!users) {
    users = ss.insertSheet(USERS_SHEET);
    users.appendRow(['Username','Password','Role','Nama']);
    users.setFrozenRows(1);
    users.getRange(1,1,1,4).setFontWeight('bold').setBackground('#ffd8e5');
    users.appendRow(['admin','admin123','admin','Administrator']);
  }
}

// ===== HELPERS =====

function getUserDisplayName_(username) {
  if (!username) return 'User';
  try {
    const ss = getSS_();
    const sh = ss.getSheetByName(USERS_SHEET);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return username;
    const values = sh.getRange(2,1,lastRow-1,4).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(username).trim()) {
        return values[i][3] || values[i][0] || 'User';
      }
    }
  } catch(e) {}
  return username || 'User';
}

// ===== CRUD =====

function saveData(module, data, username) {
  setupSheets_();
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan: ' + module);
  if (typeof data === 'string') { try { data = JSON.parse(data); } catch(e) {} }
  const ss = getSS_();
  const sh = ss.getSheetByName(module);
  const headers = SHEETS[module];
  if (headers.includes('NIK') && data['NIK'] && !/^\d{16}$/.test(String(data['NIK']).trim())) {
    throw new Error('NIK harus terdiri dari 16 digit angka.');
  }
  const id = Utilities.getUuid();
  const values = headers.map(h => data[h] == null ? '' : data[h]);
  sh.appendRow([id, new Date()].concat(values));
  const log = ss.getSheetByName('Aktivitas');
  let user = getUserDisplayName_(username);
  log.appendRow([new Date(), module.toUpperCase(), 'Data baru ditambahkan', user]);
  return {ok:true, id:id, message:'Data berhasil disimpan'};
}

function editData(module, id, data, username) {
  setupSheets_();
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan: ' + module);
  if (typeof data === 'string') { try { data = JSON.parse(data); } catch(e) {} }
  const ss = getSS_();
  const sh = ss.getSheetByName(module);
  const headers = SHEETS[module];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('Data tidak ditemukan.');
  const ids = sh.getRange(2,1,lastRow-1,1).getValues();
  let rowIndex = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) { rowIndex = i + 2; break; }
  }
  if (rowIndex === -1) throw new Error('Data dengan ID tersebut tidak ditemukan.');
  const values = headers.map(h => data[h] == null ? '' : data[h]);
  sh.getRange(rowIndex, 3, 1, headers.length).setValues([values]);
  sh.getRange(rowIndex, 2, 1, 1).setValue(new Date());
  const log = ss.getSheetByName('Aktivitas');
  let user = getUserDisplayName_(username);
  log.appendRow([new Date(), module.toUpperCase(), 'Data diperbarui', user]);
  return {ok:true, message:'Data berhasil diperbarui'};
}

function deleteData(module, id, username) {
  setupSheets_();
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan: ' + module);
  const ss = getSS_();
  const sh = ss.getSheetByName(module);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('Data tidak ditemukan.');
  const ids = sh.getRange(2,1,lastRow-1,1).getValues();
  let rowIndex = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) { rowIndex = i + 2; break; }
  }
  if (rowIndex === -1) throw new Error('Data dengan ID tersebut tidak ditemukan.');
  sh.deleteRow(rowIndex);
  const log = ss.getSheetByName('Aktivitas');
  let user = getUserDisplayName_(username);
  log.appendRow([new Date(), module.toUpperCase(), 'Data dihapus', user]);
  return {ok:true, message:'Data berhasil dihapus'};
}

function getRows_(module) {
  setupSheets_();
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan: ' + module);
  const sh = getSS_().getSheetByName(module);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  const headers = ['ID','Waktu Input'].concat(SHEETS[module]);
  return values.map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i] || '');
    return obj;
  });
}

function getData(module) {
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan');
  return getRows_(module);
}

function getPatients() {
  return getRows_('pasien').slice(-100).reverse();
}

function getDashboard() {
  setupSheets_();
  const cache = {};
  function rows(m){if(!cache[m])cache[m]=getRows_(m);return cache[m];}

  const counts = {};
  Object.keys(SHEETS).forEach(m => { if (m !== 'monitoring') counts[m] = rows(m).length; });

  const kbRows = rows('kb');
  const norm = v => String(v || '').trim().toLowerCase();
  const kbBaru = kbRows.filter(r => norm(r['Status Kepesertaan']) === 'kb baru').length;
  const kbAktif = kbRows.filter(r => norm(r['Status Kepesertaan']) === 'kb aktif').length;
  const gantiCara = kbRows.filter(r => norm(r['Status Kepesertaan']) === 'ganti cara').length;

  const log = getSS_().getSheetByName('Aktivitas');
  const last = log.getLastRow();
  const activities = last > 1
    ? log.getRange(Math.max(2,last-4),1,Math.min(5,last-1),4).getDisplayValues().reverse()
    : [];

  const pasienRows = rows('pasien');
  const totalPasien = pasienRows.length;
  const totalANC = counts.anc || 0;
  const totalPersalinan = counts.persalinan || 0;
  const totalNifas = counts.nifas || 0;
  const totalKB_ = counts.kb || 0;
  const totalBayi = counts.bayi || 0;
  const totalKunjungan = totalANC + totalPersalinan + totalNifas + totalKB_;

  const now = new Date();
  const thisYear = now.getFullYear();
  const monthlyLabels = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  const monthlyIbu = Array(12).fill(0);
  const monthlyBayi = Array(12).fill(0);

  rows('anc').forEach(r => {
    const d = new Date(r['Tanggal Kunjungan'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyIbu[d.getMonth()]++;
  });
  rows('persalinan').forEach(r => {
    const d = new Date(r['Tanggal Persalinan'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyIbu[d.getMonth()]++;
  });
  rows('nifas').forEach(r => {
    const d = new Date(r['Tanggal Kunjungan Nifas'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyIbu[d.getMonth()]++;
  });
  rows('bayi').forEach(r => {
    const d = new Date(r['Tanggal Lahir'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyBayi[d.getMonth()]++;
  });

  return {
    totalPasien: totalPasien,
    totalKunjungan: totalKunjungan,
    totalIbu: totalPasien,
    totalBayi: totalBayi,
    pasienAktif: totalPasien,
    totalKB: totalKB_,
    totalANC: totalANC,
    totalPersalinan: totalPersalinan,
    totalNifas: totalNifas,
    counts: counts,
    kbBaru: kbBaru,
    kbAktif: kbAktif,
    gantiCara: gantiCara,
    lastPatients: pasienRows.slice(-5).reverse().map(p => [p['No RM']||'', p['Nama Ibu']||'', p['Umur']||'', p['Status']||'Aktif']),
    activityLog: activities,
    monthlyLabels: monthlyLabels,
    monthlyIbu: monthlyIbu,
    monthlyBayi: monthlyBayi,
    pieLabels: ['ANC','Persalinan','Nifas','Bayi','KB'],
    pieData: [totalANC, totalPersalinan, totalNifas, totalBayi, totalKB_]
  };
}

function searchPatient(query) {
  query = String(query || '').trim().toLowerCase();
  if (!query) return [];
  return getRows_('pasien').filter(p =>
    [p['No RM'],p['Nama Ibu'],p['NIK'],p['Nama Suami']]
      .some(v => String(v).toLowerCase().includes(query))
  ).slice(-30).reverse();
}

function getSheetHeaders(module) {
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan: ' + module);
  return SHEETS[module];
}

function getFullDashboard() {
  setupSheets_();
  const cache = {};
  function rows(m){if(!cache[m])cache[m]=getRows_(m);return cache[m];}

  const counts = {};
  Object.keys(SHEETS).forEach(m => { if (m !== 'monitoring') counts[m] = rows(m).length; });

  const kbRows = rows('kb');
  const norm = v => String(v || '').trim().toLowerCase();
  const kbBaru = kbRows.filter(r => norm(r['Status Kepesertaan']) === 'kb baru').length;
  const kbAktif = kbRows.filter(r => norm(r['Status Kepesertaan']) === 'kb aktif').length;
  const gantiCara = kbRows.filter(r => norm(r['Status Kepesertaan']) === 'ganti cara').length;

  const logSheet = getSS_().getSheetByName('Aktivitas');
  const last = logSheet.getLastRow();
  const activities = last > 1
    ? logSheet.getRange(Math.max(2,last-4),1,Math.min(5,last-1),4).getDisplayValues().reverse()
    : [];

  const pasienRows = rows('pasien');
  const ancRows = rows('anc');
  const persalinanRows = rows('persalinan');
  const nifasRows = rows('nifas');
  const bayiRows = rows('bayi');
  const totalPasien = pasienRows.length;
  const totalANC = counts.anc || 0;
  const totalPersalinan = counts.persalinan || 0;
  const totalNifas = counts.nifas || 0;
  const totalKB_ = counts.kb || 0;
  const totalBayi = counts.bayi || 0;
  const totalKunjungan = totalANC + totalPersalinan + totalNifas + totalKB_;

  const now = new Date();
  const thisYear = now.getFullYear();
  const monthlyLabels = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  const monthlyIbu = Array(12).fill(0);
  const monthlyBayi = Array(12).fill(0);

  rows('anc').forEach(r => {
    const d = new Date(r['Tanggal Kunjungan'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyIbu[d.getMonth()]++;
  });
  persalinanRows.forEach(r => {
    const d = new Date(r['Tanggal Persalinan'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyIbu[d.getMonth()]++;
  });
  nifasRows.forEach(r => {
    const d = new Date(r['Tanggal Kunjungan Nifas'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyIbu[d.getMonth()]++;
  });
  bayiRows.forEach(r => {
    const d = new Date(r['Tanggal Lahir'] || r['Waktu Input']);
    if (!isNaN(d) && d.getFullYear() === thisYear) monthlyBayi[d.getMonth()]++;
  });

  // kelengkapan
  const kelengkapan = pasienRows.map(p => {
    const noRM = String(p['No RM']||'').trim();
    return {
      noRM: noRM,
      nama: p['Nama Ibu']||'',
      anc: ancRows.some(a => String(a['No RM']||'').trim() === noRM),
      persalinan: persalinanRows.some(a => String(a['No RM']||'').trim() === noRM),
      nifas: nifasRows.some(a => String(a['No RM']||'').trim() === noRM),
      bayi: bayiRows.some(a => String(a['No RM']||'').trim() === noRM),
      kb: kbRows.some(a => String(a['No RM']||'').trim() === noRM)
    };
  });

  return {
    totalPasien: totalPasien,
    totalKunjungan: totalKunjungan,
    totalIbu: totalPasien,
    totalBayi: totalBayi,
    pasienAktif: totalPasien,
    totalKB: totalKB_,
    totalANC: totalANC,
    totalPersalinan: totalPersalinan,
    totalNifas: totalNifas,
    counts: counts,
    kbBaru: kbBaru,
    kbAktif: kbAktif,
    gantiCara: gantiCara,
    lastPatients: pasienRows.slice(-5).reverse().map(p => [p['No RM']||'', p['Nama Ibu']||'', p['Umur']||'', p['Status']||'Aktif']),
    activityLog: activities,
    monthlyLabels: monthlyLabels,
    monthlyIbu: monthlyIbu,
    monthlyBayi: monthlyBayi,
    pieLabels: ['ANC','Persalinan','Nifas','Bayi','KB'],
    pieData: [totalANC, totalPersalinan, totalNifas, totalBayi, totalKB_],
    kelengkapan: kelengkapan
  };
}

function getModuleSummary() {
  setupSheets_();
  const out = {};
  Object.keys(SHEETS).forEach(m => out[m] = getRows_(m).length);
  return out;
}

function getKelengkapanData() {
  setupSheets_();
  const cache = {};
  function rows(m){if(!cache[m])cache[m]=getRows_(m);return cache[m];}
  const pasienRows = rows('pasien');
  const ancRows = rows('anc');
  const persalinanRows = rows('persalinan');
  const nifasRows = rows('nifas');
  const bayiRows = rows('bayi');
  const kbRows = rows('kb');
  return pasienRows.map(p => {
    const noRM = String(p['No RM']||'').trim();
    return {
      noRM: noRM,
      nama: p['Nama Ibu']||'',
      anc: ancRows.some(a => String(a['No RM']||'').trim() === noRM),
      persalinan: persalinanRows.some(a => String(a['No RM']||'').trim() === noRM),
      nifas: nifasRows.some(a => String(a['No RM']||'').trim() === noRM),
      bayi: bayiRows.some(a => String(a['No RM']||'').trim() === noRM),
      kb: kbRows.some(a => String(a['No RM']||'').trim() === noRM)
    };
  });
}

function getMonthlyStats() {
  setupSheets_();
  const modules = ['anc','persalinan','nifas','bayi','kb'];
  const monthNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const result = {};
  const cache = {};
  function rows(m){if(!cache[m])cache[m]=getRows_(m);return cache[m];}
  const now = new Date();
  const thisYear = now.getFullYear();
  modules.forEach(m => {
    const counts = Array(12).fill(0);
    rows(m).forEach(r => {
      const d = new Date(r['Waktu Input']);
      if (!isNaN(d) && d.getFullYear() === thisYear) counts[d.getMonth()]++;
    });
    result[m] = counts;
  });
  const monthly = {};
  for (let i = 0; i < 12; i++) {
    monthly[monthNames[i]] = {
      anc: result.anc[i],
      persalinan: result.persalinan[i],
      nifas: result.nifas[i],
      bayi: result.bayi[i],
      kb: result.kb[i]
    };
  }
  return monthly;
}

// ===== AUTH =====

function loginUser(username, password) {
  setupSheets_();
  const ss = getSS_();
  const sh = ss.getSheetByName(USERS_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {ok:false, message:'Tidak ada pengguna terdaftar'};
  const values = sh.getRange(2,1,lastRow-1,4).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(username).trim() &&
        String(values[i][1]).trim() === String(password).trim()) {
      return {ok:true, username:values[i][0], role:values[i][2], nama:values[i][3]};
    }
  }
  return {ok:false, message:'Username atau password salah'};
}

function changePassword(username, oldPw, newPw) {
  setupSheets_();
  const ss = getSS_();
  const sh = ss.getSheetByName(USERS_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('Pengguna tidak ditemukan');
  const values = sh.getRange(2,1,lastRow-1,4).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(username).trim() &&
        String(values[i][1]).trim() === String(oldPw).trim()) {
      sh.getRange(i+2, 2, 1, 1).setValue(newPw);
      return {ok:true, message:'Password berhasil diubah'};
    }
  }
  throw new Error('Password lama salah');
}

function getUsers() {
  setupSheets_();
  const ss = getSS_();
  const sh = ss.getSheetByName(USERS_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2,1,lastRow-1,4).getValues();
  return values.map(r => ({username:r[0], role:r[2], nama:r[3]}));
}

function addUser(username, password, role, nama) {
  setupSheets_();
  const ss = getSS_();
  const sh = ss.getSheetByName(USERS_SHEET);
  sh.appendRow([username, password, role, nama]);
  return {ok:true, message:'Pengguna berhasil ditambahkan'};
}

function deleteUser(username) {
  setupSheets_();
  const ss = getSS_();
  const sh = ss.getSheetByName(USERS_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('Pengguna tidak ditemukan');
  const values = sh.getRange(2,1,lastRow-1,4).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(username).trim()) {
      sh.deleteRow(i+2);
      return {ok:true, message:'Pengguna berhasil dihapus'};
    }
  }
  throw new Error('Pengguna tidak ditemukan');
}

// ===== MONITORING =====

function getActivityLog() {
  setupSheets_();
  const ss = getSS_();
  const log = ss.getSheetByName('Aktivitas');
  const last = log.getLastRow();
  if (last < 2) return [];
  return log.getRange(2, 1, last - 1, 4).getDisplayValues().reverse().slice(0, 50);
}

function getMonitoringData() {
  setupSheets_();
  const cache = {};
  function rows(m){if(!cache[m])cache[m]=getRows_(m);return cache[m];}

  const allModules = ['pasien','anc','persalinan','nifas','bayi','kb'];
  const counts = {};
  allModules.forEach(m => counts[m] = rows(m).length);

  const ancRows = rows('anc');
  const pasienRows = rows('pasien');
  const bayiRows = rows('bayi');
  const persalinanRows = rows('persalinan');

  const totalKIA = counts.anc + counts.persalinan + counts.nifas + counts.bayi;
  const totalKB = counts.kb;

  let dataLengkap = 0;
  let dataBelumLengkap = 0;

  pasienRows.forEach(p => {
    const noRM = String(p['No RM']||'').trim();
    if (!noRM) return;
    const hasANC = ancRows.some(a => String(a['No RM']||'').trim() === noRM);
    const hasPersalinan = persalinanRows.some(a => String(a['No RM']||'').trim() === noRM);
    if (hasANC && hasPersalinan) dataLengkap++;
    else dataBelumLengkap++;
  });

  const monthlyStats = getMonthlyStats();
  try {
    syncMonitoring_(monthlyStats, {
      totalKIA: totalKIA, totalKB: totalKB,
      dataLengkap: dataLengkap, dataBelumLengkap: dataBelumLengkap,
      dataIbu: pasienRows.length, dataBayi: bayiRows.length
    });
  } catch (e) {}

  return {
    counts: counts,
    totalKIA: totalKIA,
    totalKB: totalKB,
    dataLengkap: dataLengkap,
    dataBelumLengkap: dataBelumLengkap,
    dataIbu: pasienRows.length,
    dataBayi: bayiRows.length,
    monthlyStats: monthlyStats
  };
}

// ponytail: monitoring sheet is a derived report, refreshed when the Monitoring page opens.
// Monthly rows = activity that month; kelengkapan columns live on the Rekap row.
function syncMonitoring_(monthlyStats, totals) {
  const sh = getSS_().getSheetByName('monitoring');
  if (!sh) return;
  const year = new Date().getFullYear();
  const rows = Object.keys(monthlyStats).map(m => {
    const s = monthlyStats[m] || {};
    const kia = (s.anc || 0) + (s.persalinan || 0) + (s.nifas || 0) + (s.bayi || 0);
    return [m, year, kia, s.kb || 0, '', '', '', '', ''];
  });
  rows.push(['Rekap', year, totals.totalKIA, totals.totalKB,
    totals.dataLengkap, totals.dataBelumLengkap, totals.dataIbu, totals.dataBayi, '']);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
  const width = 2 + SHEETS.monitoring.length;
  sh.getRange(2, 1, rows.length, width).setValues(rows.map(r => ['', new Date()].concat(r)));
}

// ===== EXPORT =====

function exportToExcel(module) {
  setupSheets_();
  if (!module) module = 'pasien';
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan');
  const ss = getSS_();
  const url = ss.getUrl();
  const sh = ss.getSheetByName(module);
  if (!sh) throw new Error('Sheet tidak ditemukan');
  const gid = sh.getSheetId();
  const exportUrl = url.replace(/\/edit$/, '') + '/export?format=xlsx&gid=' + gid;
  return {ok:true, url: exportUrl};
}

function exportToPDF(module) {
  setupSheets_();
  if (!module) module = 'pasien';
  if (!SHEETS[module]) throw new Error('Modul tidak ditemukan');
  const ss = getSS_();
  const url = ss.getUrl();
  const sh = ss.getSheetByName(module);
  if (!sh) throw new Error('Sheet tidak ditemukan');
  const gid = sh.getSheetId();
  const exportUrl = url.replace(/\/edit$/, '') + '/export?format=pdf&gid=' + gid;
  return {ok:true, url: exportUrl};
}

function getExportUrl(module) {
  const ss = getSS_();
  const url = ss.getUrl();
  const gid = ss.getSheetByName(module).getSheetId();
  return url.replace(/edit$/, '') + 'export?format=xlsx&gid=' + gid;
}

// ===== BARCODE =====

function getBarcodeData(noRM) {
  setupSheets_();
  noRM = String(noRM||'').trim();
  if (!noRM) return null;
  const pasien = getRows_('pasien').find(p => String(p['No RM']).trim() === noRM);
  if (!pasien) return null;
  return {
    noRM: noRM,
    namaIbu: pasien['Nama Ibu'] || '',
    namaSuami: pasien['Nama Suami'] || '',
    nik: pasien['NIK'] || '',
    umur: pasien['Umur'] || '',
    alamat: pasien['Alamat/Desa'] || ''
  };
}

function setupDatabase() {
  setupSheets_();
  return 'Database SMART KIA siap digunakan.';
}
