/**
 * CBT NUSANTARA - Backend (Google Apps Script Web App)
 * ------------------------------------------------------------------
 * Deploy: Extensions > Apps Script > Deploy > New deployment > Web app
 *   Execute as : Me
 *   Who has access : Anyone
 * Salin URL ".../exec" ke konstanta API_URL di index.html.
 *
 * Sheet yang dipakai (nama harus sama persis):
 *   "Siswa"            : Nama Sekolah | Nama siswa | Nomor peserta | Kelas | Password
 *   "Guru / Admin"     : Nama Sekolah | Nama Guru | Username / NIP | Password
 *   "Pengaturan Ujian" : ID Ujian | Nama Sekolah | Nama Ujian | Link Google Form |
 *                        Batas Pelanggaran | Durasi (Menit) | Status (Aktif/Nonaktif)
 *   Kolom tambahan (dibuat otomatis saat admin menyimpan ujian, boleh juga ditambah manual):
 *                        Jadwal Mulai | Jadwal Selesai | Token | Kelas   (format jadwal: yyyy-MM-dd HH:mm;
 *                        Kelas: kosong = semua kelas; contoh "11", "XI IPA", "XI IPA 1, XI IPA 2")
 *   "Log Ujian", "Soal", "Jawaban", "Gambar" : dibuat otomatis oleh script ini (untuk ujian dengan soal internal)
 *
 * Catatan:
 *  - Kolom "Nama Sekolah" di sheet Guru boleh diisi SEMUA agar guru bisa
 *    mengelola semua sekolah. Selain itu, guru hanya melihat sekolahnya sendiri.
 *  - Password boleh teks biasa, atau hash dengan awalan "sha256:" (hex SHA-256).
 *  - Link Google Form tidak pernah dikirim ke browser sebelum siswa login.
 */

const SS_ID = '1aYsogguasQCUQwNbbsHNS6ASU3calR6_XdZvxMxSHTQ';
const SH = { SISWA: 'Siswa', ADMIN: 'Guru / Admin', EXAM: 'Pengaturan Ujian', LOG: 'Log Ujian', SOAL: 'Soal', JAWAB: 'Jawaban', IMG: 'Gambar' };
const TZ = 'Asia/Jakarta', TZ_LABEL = 'WIB';   // WITA: 'Asia/Makassar', WIT: 'Asia/Jayapura'. Semua jadwal memakai zona waktu ini.
const SESSION_TTL = 6 * 60 * 60;   // detik (maksimum CacheService = 21600)
const MAX_FAIL = 5;                // gagal login sebelum diblok sementara
const FAIL_WINDOW = 300;           // detik

const LOG_HEADERS = ['ID Ujian', 'Nama Sekolah', 'Nomor Peserta', 'Nama Siswa', 'Kelas', 'Mulai',
  'Batas Selesai', 'Pelanggaran', 'Status', 'Pelanggaran Terakhir', 'Terakhir Aktif', 'Riwayat', 'Tambahan (menit)'];
const L = { ID: 0, SCHOOL: 1, NO: 2, NAME: 3, CLASS: 4, START: 5, END: 6, VIO: 7, STATUS: 8, LAST: 9, SEEN: 10, HIST: 11, EXTRA: 12 };

/* ===================== ENTRY POINTS ===================== */

function doGet(e) {
  try {
    const a = (e && e.parameter && e.parameter.action) || 'ping';
    if (a === 'ping') return out_({ ok: true, service: 'CBT NUSANTARA', time: Date.now() });
    if (a === 'schools') return out_({ ok: true, schools: schools_() });
    return out_({ ok: false, error: 'Aksi GET tidak dikenal' });
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = ACTIONS[req.action];
    if (!fn) return out_({ ok: false, error: 'Aksi tidak dikenal' });
    return out_(fn(req));
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
}

const ACTIONS = {
  loginSiswa: loginSiswa_, getExams: getExams_, startExam: startExam_, heartbeat: heartbeat_,
  violation: violation_, finish: finish_,
  loginAdmin: loginAdmin_, adminMeta: adminMeta_, adminGetExams: adminGetExams_,
  adminSaveExam: adminSaveExam_, adminDeleteExam: adminDeleteExam_,
  adminGetLogs: adminGetLogs_, adminUnlock: adminUnlock_, adminAddTime: adminAddTime_,
  adminGetSoal: adminGetSoal_, adminSaveSoal: adminSaveSoal_, adminDeleteSoal: adminDeleteSoal_, adminReorderSoal: adminReorderSoal_,
  adminUploadImage: adminUploadImage_, getImages: getImages_, saveAnswers: saveAnswers_,
  adminGetHasil: adminGetHasil_, adminGetJawaban: adminGetJawaban_, adminNilaiEssay: adminNilaiEssay_
};

/* ===================== UTIL ===================== */

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
let _ss = null;
function ss_() { return _ss || (_ss = SpreadsheetApp.openById(SS_ID)); }   // dibuka sekali per eksekusi
function norm_(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function str_(v) { return String(v === undefined || v === null ? '' : v).trim(); }
function same_(a, b) { return str_(a).toLowerCase() === str_(b).toLowerCase(); }

function table_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + name);
  const v = sh.getDataRange().getValues();
  const headers = (v[0] || []).map(norm_);
  const rows = v.slice(1).map(function (r, i) {
    const o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  }).filter(function (o) { return headers.some(function (h) { return str_(o[h]) !== ''; }); });
  return { sh: sh, headers: headers, rows: rows };
}

function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function passOk_(stored, input) {
  stored = str_(stored); input = String(input === undefined ? '' : input);
  if (!stored) return false;
  if (stored.indexOf('sha256:') === 0) return stored.slice(7).toLowerCase() === sha256_(input);
  return stored === input;
}

function throttle_(key) {
  const c = CacheService.getScriptCache(), k = 'fail_' + key;
  if (Number(c.get(k) || 0) >= MAX_FAIL) throw new Error('Terlalu banyak percobaan. Coba lagi dalam beberapa menit.');
}
function fail_(key) {
  const c = CacheService.getScriptCache(), k = 'fail_' + key;
  c.put(k, String(Number(c.get(k) || 0) + 1), FAIL_WINDOW);
}
function clearFail_(key) { CacheService.getScriptCache().remove('fail_' + key); }

function newSession_(data) {
  const t = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  CacheService.getScriptCache().put('tok_' + t, JSON.stringify(data), SESSION_TTL);
  return t;
}
function session_(token, role) {
  const raw = token && CacheService.getScriptCache().get('tok_' + token);
  if (!raw) throw new Error('Sesi berakhir. Silakan login kembali.');
  const s = JSON.parse(raw);
  if (role && s.role !== role) throw new Error('Akses ditolak.');
  return s;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}


/* Sheet "Pengaturan Ujian": cocokkan header secara longgar (mis. "Status" = "Status (Aktif/Nonaktif)") */
function canonExamHeader_(h) {
  if (/^id/.test(h)) return 'idujian';
  if (h.indexOf('sekolah') > -1 && h.indexOf('jadwal') < 0) return 'namasekolah';
  if (h.indexOf('link') > -1 || h.indexOf('form') > -1) return 'linkgoogleform';
  if (h.indexOf('batas') > -1 || h.indexOf('pelanggaran') > -1) return 'bataspelanggaran';
  if (h.indexOf('kelas') > -1) return 'kelas';
  if (h.indexOf('sumber') > -1) return 'sumbersoal';
  if (h.indexOf('opsi') > -1) return 'opsiujian';
  if (h.indexOf('mulai') > -1) return 'jadwalmulai';
  if (h.indexOf('selesai') > -1) return 'jadwalselesai';
  if (h.indexOf('token') > -1) return 'token';
  if (h.indexOf('durasi') > -1 || h.indexOf('menit') > -1 || h.indexOf('waktu') > -1) return 'durasimenit';
  if (h.indexOf('status') > -1) return 'statusaktifnonaktif';
  if (h.indexOf('ujian') > -1 || h.indexOf('nama') > -1) return 'namaujian';
  return h;
}

function examTable_() {
  const t = table_(SH.EXAM);
  const map = {};
  t.headers.forEach(function (h) { map[h] = canonExamHeader_(h); });
  t.rows = t.rows.map(function (r) {
    const o = { _row: r._row };
    t.headers.forEach(function (h) { o[map[h]] = r[h]; });
    return o;
  });
  t.headers = t.headers.map(function (h) { return map[h]; });
  return t;
}


/* Cache daftar ujian 30 dtk: pembacaan siswa tidak perlu membuka spreadsheet. Dibuang saat admin menyimpan. */
const EXAM_CACHE_TTL = 30;
function examRows_() {
  const c = CacheService.getScriptCache(), hit = c.get('exams');
  if (hit) return JSON.parse(hit);
  const rows = examTable_().rows;
  try { c.put('exams', JSON.stringify(rows), EXAM_CACHE_TTL); } catch (e) { /* terlalu besar: abaikan */ }
  return rows;
}
function bustExams_() { CacheService.getScriptCache().remove('exams'); }

/* ---- konversi baris sheet "Pengaturan Ujian" ---- */
function isActive_(o) { return str_(o.statusaktifnonaktif).toLowerCase() === 'aktif'; }
function tzOffsetMs_() {
  const z = Utilities.formatDate(new Date(), TZ, 'Z');                 // mis. +0700
  return (z.charAt(0) === '-' ? -1 : 1) * (parseInt(z.slice(1, 3), 10) * 60 + parseInt(z.slice(3, 5), 10)) * 60000;
}
/* Jadwal boleh berupa teks "yyyy-MM-dd HH:mm" (dibaca dalam TZ) atau sel tanggal bawaan Google Sheets. */
function parseSched_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') { const t = v.getTime(); return isNaN(t) ? null : t; }
  const x = String(v).trim();
  if (/Z$/.test(x)) { const t = Date.parse(x); return isNaN(t) ? null : t; }          // tanggal dari cache (ISO UTC)
  const m = x.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})/);
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) - tzOffsetMs_() : null;
}
function fmtSched_(ms) { return ms === null || ms === undefined ? '' : Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd HH:mm'); }

/* Nonaktif | Aktif (tanpa jadwal) | Terjadwal (belum buka) | Berlangsung | Selesai (jadwal tutup terlewati) */
function examState_(o, now) {
  if (!isActive_(o)) return 'Nonaktif';
  const a = parseSched_(o.jadwalmulai), b = parseSched_(o.jadwalselesai);
  if (a === null && b === null) return 'Aktif';
  if (a !== null && now < a) return 'Terjadwal';
  if (b !== null && now > b) return 'Selesai';
  return 'Berlangsung';
}
function examPublic_(o) {
  const a = parseSched_(o.jadwalmulai), b = parseSched_(o.jadwalselesai);
  return {
    id: str_(o.idujian), sekolah: str_(o.namasekolah), nama: str_(o.namaujian),
    batas: Number(o.bataspelanggaran) || 3, durasi: Number(o.durasimenit) || 60,
    status: isActive_(o) ? 'Aktif' : 'Nonaktif', state: examState_(o, Date.now()),
    mulaiMs: a, selesaiMs: b, perluToken: str_(o.token) !== '', kelas: str_(o.kelas),
    sumber: /internal/i.test(str_(o.sumbersoal)) ? 'internal' : 'form', opsiUjian: parseOpsiUjian_(o.opsiujian)            // token itu sendiri tidak pernah dikirim ke siswa
  };
}

function examFull_(o) {
  const e = examPublic_(o);
  e.link = str_(o.linkgoogleform); e.token = str_(o.token);
  e.mulai = fmtSched_(e.mulaiMs).replace(' ', 'T'); e.selesai = fmtSched_(e.selesaiMs).replace(' ', 'T');   // untuk input datetime-local
  return e;
}

function normalizeFormUrl_(raw) {
  let u = str_(raw);
  const m = u.match(/src\s*=\s*["']([^"']+)["']/i);       // jika yang ditempel kode <iframe>
  if (m) u = m[1];
  u = u.replace(/&amp;/g, '&');
  if (!/^https:\/\/(docs\.google\.com\/forms\/|forms\.gle\/)/i.test(u))
    throw new Error('Link harus berupa Google Form (https://docs.google.com/forms/... atau https://forms.gle/...)');
  if (/\/viewform/i.test(u) && !/[?&]embedded=true/i.test(u)) u += (u.indexOf('?') > -1 ? '&' : '?') + 'embedded=true';
  return u;
}

/* Kelas: "11", "XI", "11 IPA", "XI IPA 1" dianggap setara; boleh banyak, dipisah koma. Kosong = semua kelas. */
const ROMAN_ = { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6, VII: 7, VIII: 8, IX: 9, X: 10, XI: 11, XII: 12, XIII: 13 };
function kelasTokens_(v) {
  const t = String(v === undefined || v === null ? '' : v).toUpperCase()
    .replace(/(\d)([A-Z])/g, '$1 $2').replace(/([A-Z])(\d)/g, '$1 $2')       // "11A" -> "11 A"
    .replace(/[^A-Z0-9]+/g, ' ').trim().split(' ').filter(Boolean);
  if (t.length && ROMAN_[t[0]]) t[0] = String(ROMAN_[t[0]]);                  // "XI" -> "11"
  return t;
}
function kelasMatch_(examKelas, siswaKelas) {
  const entries = String(examKelas === undefined || examKelas === null ? '' : examKelas).split(/[,;\n]+/).map(kelasTokens_).filter(function (t) { return t.length; });
  if (!entries.length) return true;
  const sk = kelasTokens_(siswaKelas);
  return entries.some(function (en) { return en.length <= sk.length && en.every(function (tok, i) { return tok === sk[i]; }); });
}

/* Daftar sekolah dan kelas dari sheet Siswa (cache 2 menit) */
function siswaMeta_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('siswaMeta');
  if (hit) return JSON.parse(hit);
  const sc = {}, kl = {};
  table_(SH.SISWA).rows.forEach(function (r) {
    const sk = str_(r.namasekolah); if (!sk) return;
    sc[sk] = 1;
    const k = str_(r.kelas); if (k) { (kl[sk] = kl[sk] || {})[k] = 1; }
  });
  const meta = { schools: Object.keys(sc).sort(), kelas: {} };
  Object.keys(kl).forEach(function (sk) { meta.kelas[sk] = Object.keys(kl[sk]).sort(); });
  try { cache.put('siswaMeta', JSON.stringify(meta), 120); } catch (e) { /* abaikan */ }
  return meta;
}
function schools_() { return siswaMeta_().schools; }

/* ===================== LOG UJIAN ===================== */

function logSheet_() {
  let sh = ss_().getSheetByName(SH.LOG);
  if (!sh) {
    sh = ss_().insertSheet(SH.LOG);
    sh.appendRow(LOG_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold').setBackground('#E0E7FF');
  } else {
    const c = CacheService.getScriptCache();
    if (!c.get('logHdrOk')) {                                              // sheet lama: tambahkan kolom yang baru
      LOG_HEADERS.forEach(function (h, i) { if (!str_(sh.getRange(1, i + 1).getValue())) sh.getRange(1, i + 1).setValue(h); });
      c.put('logHdrOk', '1', 21600);
    }
  }
  return sh;
}

function findLog_(sh, idUjian, no) {
  const c = CacheService.getScriptCache(), key = 'row_' + idUjian + '_' + no, w = LOG_HEADERS.length;
  const hint = Number(c.get(key)) || 0;
  if (hint > 1 && hint <= sh.getLastRow()) {
    const vals = sh.getRange(hint, 1, 1, w).getValues()[0];
    if (str_(vals[L.ID]) === idUjian && str_(vals[L.NO]) === no) return { row: hint, vals: vals };
  }
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++)
    if (str_(v[i][L.ID]) === idUjian && str_(v[i][L.NO]) === no) { c.put(key, String(i + 1), 21600); return { row: i + 1, vals: v[i] }; }
  return null;
}

function stamp_() { return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM HH:mm:ss'); }
function effEnd_(vals, selesaiMs) {
  let end = new Date(vals[L.END]).getTime();
  const extra = Number(vals[L.EXTRA]) || 0;
  if (!extra && selesaiMs !== null && selesaiMs !== undefined && selesaiMs < end) end = selesaiMs;   // jadwal tutup menghentikan peserta; peserta dengan tambahan waktu dikecualikan
  return end;
}
function logState_(vals, selesaiMs) {
  return {
    status: str_(vals[L.STATUS]), violations: Number(vals[L.VIO]) || 0,
    remainingSec: Math.max(0, Math.floor((effEnd_(vals, selesaiMs) - Date.now()) / 1000))
  };
}

/* ===================== SISWA ===================== */

function loginSiswa_(req) {
  const sekolah = str_(req.sekolah), no = str_(req.nomor), key = 's_' + no;
  throttle_(key);
  const s = table_(SH.SISWA).rows.find(function (r) {
    return same_(r.namasekolah, sekolah) && str_(r.nomorpeserta) === no;
  });
  if (!s || !passOk_(s.password, req.password)) { fail_(key); throw new Error('Nomor peserta atau password salah.'); }
  clearFail_(key);
  const user = { role: 'siswa', nama: str_(s.namasiswa), kelas: str_(s.kelas), sekolah: str_(s.namasekolah), nomor: no };
  return { ok: true, token: newSession_(user), user: user, exams: examsForStudent_(user), serverNow: Date.now(), tz: TZ, tzLabel: TZ_LABEL };
}

function examsForStudent_(u) {
  const c = CacheService.getScriptCache(), out = [];
  examRows_().forEach(function (r) {
    if (!same_(r.namasekolah, u.sekolah) || !isActive_(r) || !kelasMatch_(r.kelas, u.kelas)) return;
    const e = examPublic_(r);
    if (e.state === 'Selesai') {                                              // jadwal tutup: tampil hanya bagi peserta yang diberi tambahan waktu
      if (c.get('ext_' + e.id + '_' + u.nomor)) {
        const lg = findLog_(logSheet_(), e.id, u.nomor);
        if (lg && str_(lg.vals[L.STATUS]) === 'Sedang Mengerjakan' && logState_(lg.vals, e.selesaiMs).remainingSec > 0) { e.lanjut = true; e.state = 'Berlangsung'; out.push(e); }
      }
      return;
    }
    out.push(e);
  });
  return out;
}

function getExams_(req) {
  const u = session_(req.token, 'siswa');
  return { ok: true, exams: examsForStudent_(u), serverNow: Date.now(), tz: TZ, tzLabel: TZ_LABEL };
}

function startExam_(req) {
  const u = session_(req.token, 'siswa');
  return withLock_(function () {
    const row = examRows_().find(function (r) { return str_(r.idujian) === str_(req.idUjian); });
    if (!row || !same_(row.namasekolah, u.sekolah)) throw new Error('Ujian tidak ditemukan.');
    const ex = examFull_(row), sh = logSheet_();
    if (ex.state === 'Nonaktif') throw new Error('Ujian sedang dinonaktifkan oleh admin.');
    if (ex.state === 'Terjadwal') throw new Error('Ujian belum dibuka. Jadwal mulai: ' + fmtSched_(ex.mulaiMs) + ' ' + TZ_LABEL + '.');
    let lg = findLog_(sh, ex.id, u.nomor);
    const now = new Date();
    if (ex.sumber === 'internal' && !soalList_(ex.id).length) throw new Error('Ujian belum memiliki soal. Hubungi guru.');
    if (!lg) {                                                                // sesi baru: cek jadwal tutup dan token
      if (ex.state === 'Selesai') throw new Error('Waktu ujian sudah ditutup.');
      if (!kelasMatch_(ex.kelas, u.kelas)) throw new Error('Ujian ini bukan untuk kelas Anda.');
      if (ex.token) {
        if (!str_(req.tokenUjian)) throw new Error('Token ujian diperlukan.');   // kosong: bukan tebakan, tidak dihitung
        const tk = 'tok_' + u.nomor;
        throttle_(tk);
        if (!same_(ex.token, req.tokenUjian)) { fail_(tk); throw new Error('Token ujian salah.'); }
        clearFail_(tk);
      }
      const cap = ex.selesaiMs === null ? Infinity : ex.selesaiMs;
      const end = new Date(Math.min(now.getTime() + ex.durasi * 60000, cap));
      sh.appendRow([ex.id, u.sekolah, u.nomor, u.nama, u.kelas, now, end, 0, 'Sedang Mengerjakan', '', now, '']);
      sh.getRange(sh.getLastRow(), L.NO + 1).setNumberFormat('@').setValue(u.nomor);   // simpan sebagai teks
      lg = findLog_(sh, ex.id, u.nomor);
    }                                                                         // melanjutkan sesi (refresh/buka kunci): tanpa token lagi
    const st = logState_(lg.vals, ex.selesaiMs);
    let status = st.status;
    if (status === 'Sedang Mengerjakan' && st.remainingSec <= 0) {
      status = 'Selesai (Waktu Habis)';
      sh.getRange(lg.row, L.STATUS + 1).setValue(status);
    }
    sh.getRange(lg.row, L.SEEN + 1).setValue(now);
    const res = { ok: true, status: status, violations: st.violations, remainingSec: st.remainingSec,
      batas: ex.batas, exam: { id: ex.id, nama: ex.nama, sekolah: ex.sekolah, durasi: ex.durasi } };
    if (status === 'Sedang Mengerjakan') {
      if (ex.sumber === 'internal') {                                           // soal tampil di aplikasi: kirim soal TANPA kunci + jawaban tersimpan
        const a = ensureAns_(ansSheet_(), ex.id, u);
        res.mode = 'internal'; res.opsiUjian = ex.opsiUjian;
        res.soal = soalUntukSiswa_(soalList_(ex.id), u.nomor + '|' + ex.id, ex.opsiUjian);
        res.jawaban = parseJson_(a.vals[J.JAWAB], {}); res.ragu = parseJson_(a.vals[J.RAGU], []);
      } else res.link = normalizeFormUrl_(ex.link);                              // link hanya untuk sesi aktif
    }
    return res;
  });
}

function heartbeat_(req) {
  const u = session_(req.token, 'siswa');
  const row = examRows_().find(function (r) { return str_(r.idujian) === str_(req.idUjian); });
  const sh = logSheet_(), lg = findLog_(sh, str_(req.idUjian), u.nomor);
  if (!row || !lg) throw new Error('Sesi ujian tidak ditemukan.');
  const ex = examPublic_(row), vio = Number(lg.vals[L.VIO]) || 0, status = str_(lg.vals[L.STATUS]);
  const st = logState_(lg.vals, ex.selesaiMs), client = Number(req.clientViolations) || 0;
  if (status === 'Sedang Mengerjakan' && (client > vio || st.remainingSec <= 0)) return heartbeatWrite_(req, u);   // perlu menulis: pakai lock
  sh.getRange(lg.row, L.SEEN + 1).setValue(new Date());
  return { ok: true, status: status, violations: vio, remainingSec: st.remainingSec, batas: ex.batas, examActive: ex.status === 'Aktif' };
}

function heartbeatWrite_(req, u) {
  return withLock_(function () {
    const row = examRows_().find(function (r) { return str_(r.idujian) === str_(req.idUjian); });
    const sh = logSheet_(), lg = findLog_(sh, str_(req.idUjian), u.nomor);
    if (!row || !lg) throw new Error('Sesi ujian tidak ditemukan.');
    const ex = examPublic_(row);
    let vio = Number(lg.vals[L.VIO]) || 0, status = str_(lg.vals[L.STATUS]);
    const client = Number(req.clientViolations) || 0;           // sinkron jika sempat offline
    if (status === 'Sedang Mengerjakan' && client > vio) {
      vio = client;
      sh.getRange(lg.row, L.VIO + 1).setValue(vio);
      if (vio >= ex.batas) { status = 'Terkunci'; sh.getRange(lg.row, L.STATUS + 1).setValue(status); }
    }
    const st = logState_(lg.vals, ex.selesaiMs);
    if (status === 'Sedang Mengerjakan' && st.remainingSec <= 0) {
      status = 'Selesai (Waktu Habis)'; sh.getRange(lg.row, L.STATUS + 1).setValue(status);
    }
    sh.getRange(lg.row, L.SEEN + 1).setValue(new Date());
    return { ok: true, status: status, violations: vio, remainingSec: st.remainingSec,
      batas: ex.batas, examActive: ex.status === 'Aktif' };
  });
}

function violation_(req) {
  const u = session_(req.token, 'siswa');
  return withLock_(function () {
    const row = examRows_().find(function (r) { return str_(r.idujian) === str_(req.idUjian); });
    const sh = logSheet_(), lg = findLog_(sh, str_(req.idUjian), u.nomor);
    if (!row || !lg) throw new Error('Sesi ujian tidak ditemukan.');
    const ex = examPublic_(row);
    let status = str_(lg.vals[L.STATUS]), vio = Number(lg.vals[L.VIO]) || 0;
    if (status !== 'Sedang Mengerjakan') return { ok: true, status: status, violations: vio, batas: ex.batas };
    const client = Number(req.clientViolations) || 0;
    const next = client > 0 ? Math.max(vio, client) : vio + 1;     // idempoten: diulang tidak menghitung dobel
    if (next === vio) return { ok: true, status: status, violations: vio, batas: ex.batas };
    vio = next;
    if (vio >= ex.batas) status = 'Terkunci';
    const type = str_(req.type).slice(0, 80);
    const hist = (str_(lg.vals[L.HIST]) + '\n' + stamp_() + ' - ' + type).trim().slice(-4000);
    sh.getRange(lg.row, L.VIO + 1).setValue(vio);
    sh.getRange(lg.row, L.STATUS + 1).setValue(status);
    sh.getRange(lg.row, L.LAST + 1).setValue(type);
    sh.getRange(lg.row, L.SEEN + 1).setValue(new Date());
    sh.getRange(lg.row, L.HIST + 1).setValue(hist);
    return { ok: true, status: status, violations: vio, batas: ex.batas };
  });
}

function finish_(req) {
  const u = session_(req.token, 'siswa');
  return withLock_(function () {
    const idUjian = str_(req.idUjian);
    const sh = logSheet_(), lg = findLog_(sh, idUjian, u.nomor);
    if (!lg) throw new Error('Sesi ujian tidak ditemukan.');
    const row = examRows_().find(function (r) { return str_(r.idujian) === idUjian; });
    const ex = row ? examFull_(row) : null, internal = !!ex && ex.sumber === 'internal';
    let status = str_(lg.vals[L.STATUS]);
    const running = status === 'Sedang Mengerjakan';
    const grace = /Waktu Habis/.test(status) && internal && (Date.now() - effEnd_(lg.vals, ex.selesaiMs) <= 60000);   // jawaban terakhir saat waktu habis
    if (internal && req.jawaban && (running || grace)) simpanJawaban_(idUjian, u, req.jawaban, req.ragu, true);
    if (running) {
      status = req.reason === 'timeout' ? 'Selesai (Waktu Habis)' : 'Selesai';
      sh.getRange(lg.row, L.STATUS + 1).setValue(status);
    }
    sh.getRange(lg.row, L.SEEN + 1).setValue(new Date());
    const res = { ok: true, status: status };
    if (internal && ex.opsiUjian.tampilNilai) res.hasil = ringkasHasil_(idUjian, u.nomor);
    return res;
  });
}

/* ===================== ADMIN / GURU ===================== */

function loginAdmin_(req) {
  const un = str_(req.username), key = 'a_' + un;
  throttle_(key);
  const a = table_(SH.ADMIN).rows.find(function (r) { return same_(r.usernamenip, un); });
  if (!a || !passOk_(a.password, req.password)) { fail_(key); throw new Error('Username/NIP atau password salah.'); }
  clearFail_(key);
  const scope = /^(semua|\*|all)$/i.test(str_(a.namasekolah)) ? '*' : str_(a.namasekolah);
  const user = { role: 'admin', nama: str_(a.namaguru), username: un, sekolah: scope };
  return { ok: true, token: newSession_(user), user: user };
}

function allowed_(u, sekolah) { return u.sekolah === '*' || same_(u.sekolah, sekolah); }

function schoolsFor_(u) {
  const set = {};
  schools_().concat(examRows_().map(function (r) { return str_(r.namasekolah); })).forEach(function (x) { if (x) set[x] = 1; });
  const list = Object.keys(set).sort().filter(function (x) { return allowed_(u, x); });
  if (u.sekolah !== '*' && list.indexOf(u.sekolah) < 0) list.push(u.sekolah);
  return list;
}
function adminList_(u, rows) {
  const cnt = soalCount_();
  return rows.filter(function (r) { return allowed_(u, r.namasekolah); }).map(function (r) { const e = examFull_(r); e.jumlahSoal = cnt[e.id] || 0; return e; });
}

function adminMeta_(req) {
  const u = session_(req.token, 'admin');
  return { ok: true, schools: schoolsFor_(u) };
}

function adminGetExams_(req) {
  const u = session_(req.token, 'admin');
  const km = siswaMeta_().kelas, kp = {};
  Object.keys(km).forEach(function (sk) { if (allowed_(u, sk)) kp[sk] = km[sk]; });
  return { ok: true, exams: adminList_(u, examRows_()), schools: schoolsFor_(u), kelasPerSekolah: kp, serverNow: Date.now(), tz: TZ, tzLabel: TZ_LABEL };
}

function adminSaveExam_(req) {
  const u = session_(req.token, 'admin');
  const e = req.exam || {};
  const sekolah = str_(e.sekolah), nama = str_(e.nama);
  if (!sekolah || !nama) throw new Error('Nama sekolah dan nama ujian wajib diisi.');
  if (!allowed_(u, sekolah)) throw new Error('Anda tidak berhak mengelola sekolah ini.');
  let id = str_(e.id);
  if (id && !/^[A-Za-z0-9_-]{3,40}$/.test(id)) throw new Error('ID ujian tidak valid.');
  const batas = Math.max(1, parseInt(e.batas, 10) || 3), durasi = Math.max(1, parseInt(e.durasi, 10) || 60);
  const internal = /^internal$/i.test(str_(e.sumber));
  const link = internal ? '' : normalizeFormUrl_(e.link);                   // soal internal tidak memakai Google Form
  const opsiUjian = parseOpsiUjian_(e.opsiUjian);
  const status = /^aktif$/i.test(str_(e.status)) ? 'Aktif' : 'Nonaktif';
  const mulai = str_(e.mulai) ? parseSched_(str_(e.mulai)) : null, selesai = str_(e.selesai) ? parseSched_(str_(e.selesai)) : null;
  if ((str_(e.mulai) && mulai === null) || (str_(e.selesai) && selesai === null)) throw new Error('Format jadwal tidak valid.');
  if (mulai !== null && selesai !== null && selesai <= mulai) throw new Error('Jadwal selesai harus setelah jadwal mulai.');
  const token = str_(e.token).toUpperCase().replace(/\s+/g, '');
  if (token && !/^[A-Z0-9]{3,12}$/.test(token)) throw new Error('Token harus 3 sampai 12 karakter, hanya huruf dan angka.');
  const kelas = str_(e.kelas).replace(/\s+/g, ' ');
  if (kelas.length > 200) throw new Error('Daftar kelas terlalu panjang (maksimal 200 karakter).');
  return withLock_(function () {
    const t = examTable_();                                   // baca segar (bukan cache) sebelum menulis
    [['jadwalmulai', 'Jadwal Mulai'], ['jadwalselesai', 'Jadwal Selesai'], ['token', 'Token'], ['kelas', 'Kelas'], ['sumbersoal', 'Sumber Soal'], ['opsiujian', 'Opsi Ujian']].forEach(function (c) {
      if (t.headers.indexOf(c[0]) < 0) { t.sh.getRange(1, t.headers.length + 1).setValue(c[1]); t.headers.push(c[0]); }   // tambah kolom bila belum ada
    });
    const existing = id ? t.rows.find(function (r) { return str_(r.idujian) === id; }) : null;
    if (existing && !allowed_(u, existing.namasekolah)) throw new Error('Akses ditolak.');
    if (!id) id = 'UJ-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    const values = { idujian: id, namasekolah: sekolah, namaujian: nama, linkgoogleform: link,
      bataspelanggaran: batas, durasimenit: durasi, statusaktifnonaktif: status,
      jadwalmulai: fmtSched_(mulai), jadwalselesai: fmtSched_(selesai), token: token, kelas: kelas,
      sumbersoal: internal ? 'Soal Internal' : 'Google Form', opsiujian: JSON.stringify(opsiUjian) };
    const line = t.headers.map(function (h) { return values[h] !== undefined ? values[h] : ''; });
    const target = existing ? existing._row : t.sh.getLastRow() + 1;
    t.headers.forEach(function (h, i) {                       // jadwal dan token disimpan sebagai teks (tanpa konversi otomatis Sheets)
      if (h === 'jadwalmulai' || h === 'jadwalselesai' || h === 'token' || h === 'kelas' || h === 'sumbersoal' || h === 'opsiujian') t.sh.getRange(target, i + 1).setNumberFormat('@');
    });
    t.sh.getRange(target, 1, 1, line.length).setValues([line]);
    const rec = Object.assign({ _row: target }, values);
    const rows = existing ? t.rows.map(function (r) { return r._row === target ? rec : r; }) : t.rows.concat([rec]);
    bustExams_();
    return { ok: true, id: id, exams: adminList_(u, rows) };    // klien tidak perlu memuat ulang
  });
}

function adminDeleteExam_(req) {
  const u = session_(req.token, 'admin');
  return withLock_(function () {
    const t = examTable_();
    const r = t.rows.find(function (x) { return str_(x.idujian) === str_(req.id); });
    let rows = t.rows;
    if (r) {
      if (!allowed_(u, r.namasekolah)) throw new Error('Akses ditolak.');
      t.sh.deleteRow(r._row);
      rows = t.rows.filter(function (x) { return x._row !== r._row; });
      bustExams_();
    }
    return { ok: true, exams: adminList_(u, rows) };            // idempoten: sudah terhapus tetap sukses
  });
}

function adminGetLogs_(req) {
  const u = session_(req.token, 'admin');
  const v = logSheet_().getDataRange().getValues().slice(1);
  const now = Date.now(), cap = {};
  examRows_().forEach(function (r) { cap[str_(r.idujian)] = parseSched_(r.jadwalselesai); });
  const logs = v.filter(function (r) {
    return allowed_(u, r[L.SCHOOL]) && (!req.idUjian || str_(r[L.ID]) === str_(req.idUjian));
  }).map(function (r) {
    const seen = new Date(r[L.SEEN]).getTime(), status = str_(r[L.STATUS]);
    const sisa = status === 'Sedang Mengerjakan' ? Math.max(0, Math.floor((effEnd_(r, cap[str_(r[L.ID])]) - now) / 1000)) : null;
    return { idUjian: str_(r[L.ID]), sekolah: str_(r[L.SCHOOL]), nomor: str_(r[L.NO]), nama: str_(r[L.NAME]),
      kelas: str_(r[L.CLASS]), pelanggaran: Number(r[L.VIO]) || 0, status: status,
      terakhir: str_(r[L.LAST]), online: (now - seen) < 30000, seenAgo: Math.round((now - seen) / 1000),
      sisaSec: sisa, tambahan: Number(r[L.EXTRA]) || 0 };
  });
  return { ok: true, logs: logs };
}

/* Tambah waktu untuk satu peserta. Bila peserta sudah selesai/waktu habis, ujian dibuka kembali. */
function adminAddTime_(req) {
  const u = session_(req.token, 'admin');
  const menit = parseInt(req.menit, 10);
  if (!(menit >= 1 && menit <= 240)) throw new Error('Tambahan waktu harus 1 sampai 240 menit.');
  const alasan = str_(req.alasan).slice(0, 80);
  return withLock_(function () {
    const sh = logSheet_(), idUjian = str_(req.idUjian), nomor = str_(req.nomor), lg = findLog_(sh, idUjian, nomor);
    if (!lg) throw new Error('Data peserta tidak ditemukan.');
    if (!allowed_(u, lg.vals[L.SCHOOL])) throw new Error('Akses ditolak.');
    const row = examRows_().find(function (r) { return str_(r.idujian) === idUjian; });
    if (!row) throw new Error('Ujian tidak ditemukan.');
    const ex = examPublic_(row);
    const newEnd = Math.max(effEnd_(lg.vals, ex.selesaiMs), Date.now()) + menit * 60000;
    const extra = (Number(lg.vals[L.EXTRA]) || 0) + menit;
    let status = str_(lg.vals[L.STATUS]);
    const dibuka = /^Selesai/.test(status);
    if (dibuka) status = 'Sedang Mengerjakan';
    const hist = (str_(lg.vals[L.HIST]) + '\n' + stamp_() + ' - tambah waktu ' + menit + ' menit oleh ' + u.nama + (alasan ? ' (' + alasan + ')' : '')).trim().slice(-4000);
    sh.getRange(lg.row, L.END + 1).setValue(new Date(newEnd));
    sh.getRange(lg.row, L.STATUS + 1).setValue(status);
    sh.getRange(lg.row, L.HIST + 1).setValue(hist);
    sh.getRange(lg.row, L.EXTRA + 1).setValue(extra);
    CacheService.getScriptCache().put('ext_' + idUjian + '_' + nomor, '1', 21600);   // agar ujian tampil lagi di daftar peserta walau jadwal sudah tutup
    let peringatan = '';
    if (status === 'Terkunci') peringatan = 'Peserta masih terkunci. Buka kunci agar bisa melanjutkan.';
    else if (ex.state === 'Nonaktif') peringatan = 'Ujian sedang Nonaktif. Aktifkan agar peserta bisa melanjutkan.';
    return { ok: true, status: status, tambahan: extra, dibuka: dibuka, peringatan: peringatan,
      sisaSec: Math.max(0, Math.floor((newEnd - Date.now()) / 1000)) };
  });
}

function adminUnlock_(req) {
  const u = session_(req.token, 'admin');
  return withLock_(function () {
    const sh = logSheet_(), lg = findLog_(sh, str_(req.idUjian), str_(req.nomor));
    if (!lg) throw new Error('Data tidak ditemukan.');
    if (!allowed_(u, lg.vals[L.SCHOOL])) throw new Error('Akses ditolak.');
    const hist = (str_(lg.vals[L.HIST]) + '\n' + stamp_() + ' - dibuka oleh ' + u.nama).trim().slice(-4000);
    sh.getRange(lg.row, L.VIO + 1).setValue(0);
    sh.getRange(lg.row, L.STATUS + 1).setValue('Sedang Mengerjakan');
    sh.getRange(lg.row, L.HIST + 1).setValue(hist);
    return { ok: true };
  });
}

/* ===== INTI SOAL INTERNAL: fungsi murni (sama persis di server dan mode demo) ===== */
function parseJson_(v, d) { try { const x = JSON.parse(String(v)); return x === undefined || x === null ? d : x; } catch (e) { return d; } }
function hashStr_(s) { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return h >>> 0; }
function rng_(seed) {
  let a = seed >>> 0;
  return function () { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
}
function shuffle_(arr, seedStr) {                                   // acak tetap: peserta yang sama selalu mendapat urutan yang sama
  const a = arr.slice(), r = rng_(hashStr_(seedStr));
  for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(r() * (i + 1)), t = a[i]; a[i] = a[j]; a[j] = t; }
  return a;
}
function parseOpsiUjian_(v) {
  const o = (v && typeof v === 'object') ? v : parseJson_(v, {});
  return { acakSoal: !!o.acakSoal, acakOpsi: !!o.acakOpsi, tampilNilai: o.tampilNilai === undefined ? true : !!o.tampilNilai };
}
const IMG_RE_ = /^IMG-[A-Za-z0-9]{4,30}$/;

/* Validasi soal dari guru. Mengembalikan soal bersih atau melempar pesan galat yang jelas. */
function cleanSoal_(s) {
  const jenis = str_(s.jenis).toUpperCase();
  if (['PG', 'PGK', 'BS', 'ESSAY'].indexOf(jenis) < 0) throw new Error('Jenis soal tidak valid.');
  const teks = String(s.teks === undefined || s.teks === null ? '' : s.teks).slice(0, 4000), gambar = str_(s.gambar);
  if (gambar && !IMG_RE_.test(gambar)) throw new Error('Gambar tidak valid.');
  if (!teks.trim() && !gambar) throw new Error('Isi pertanyaan atau tambahkan gambar.');
  let bobot = Number(s.bobot);
  if (!(bobot > 0 && bobot <= 100)) throw new Error('Bobot harus lebih dari 0 dan paling banyak 100.');
  bobot = Math.round(bobot * 100) / 100;
  let opsi = [], kunci = null, aturan = '';
  if (jenis === 'ESSAY') kunci = String(s.kunci === undefined || s.kunci === null ? '' : s.kunci).slice(0, 2000);
  else {
    const src = Array.isArray(s.opsi) ? s.opsi : [], bs = jenis === 'BS', seen = {};
    if (src.length < (bs ? 1 : 2) || src.length > (bs ? 10 : 8)) throw new Error(bs ? 'Jumlah pernyataan Benar/Salah harus 1 sampai 10.' : 'Jumlah pilihan jawaban harus 2 sampai 8.');
    opsi = src.map(function (o) {
      const id = str_(o.id), t = String(o.t === undefined || o.t === null ? '' : o.t).slice(0, 1000), g = str_(o.g);
      if (!(bs ? /^s([1-9]|10)$/ : /^[a-h]$/).test(id) || seen[id]) throw new Error('ID pilihan tidak valid.');
      seen[id] = 1;
      if (g && !IMG_RE_.test(g)) throw new Error('Gambar pilihan tidak valid.');
      if (!t.trim() && !g) throw new Error(bs ? 'Setiap pernyataan harus berisi teks atau gambar.' : 'Setiap pilihan harus berisi teks atau gambar.');
      return { id: id, t: t, g: g };
    });
    if (jenis === 'PG') {
      if (!seen[str_(s.kunci)]) throw new Error('Pilih satu kunci jawaban.');
      kunci = str_(s.kunci);
    } else if (jenis === 'PGK') {
      const k = (Array.isArray(s.kunci) ? s.kunci : []).map(str_).filter(function (x, i, a) { return a.indexOf(x) === i; });
      if (!k.length || k.some(function (x) { return !seen[x]; })) throw new Error('Pilih minimal satu kunci jawaban.');
      kunci = k.sort(); aturan = s.aturan === 'sebagian' ? 'sebagian' : 'semua';
    } else {
      const k = s.kunci && typeof s.kunci === 'object' ? s.kunci : {};
      kunci = {};
      opsi.forEach(function (o) { if (k[o.id] !== 'B' && k[o.id] !== 'S') throw new Error('Tentukan Benar atau Salah untuk setiap pernyataan.'); kunci[o.id] = k[o.id]; });
      aturan = s.aturan === 'semua' ? 'semua' : 'per';
    }
  }
  return { jenis: jenis, teks: teks, gambar: gambar, opsi: opsi, kunci: kunci, bobot: bobot, aturan: aturan };
}

/* Jawaban dari peserta dibersihkan menurut jenis soal; yang tidak valid dibuang. */
function cleanAnswers_(soal, raw) {
  const out = {}; raw = raw && typeof raw === 'object' ? raw : {};
  soal.forEach(function (q) {
    const a = raw[q.id]; if (a === undefined || a === null) return;
    const ids = (q.opsi || []).map(function (o) { return o.id; });
    if (q.jenis === 'PG') { if (ids.indexOf(String(a)) > -1) out[q.id] = String(a); }
    else if (q.jenis === 'PGK') {
      const sel = (Array.isArray(a) ? a : []).map(String).filter(function (x, i, arr) { return ids.indexOf(x) > -1 && arr.indexOf(x) === i; });
      if (sel.length) out[q.id] = sel.sort();
    } else if (q.jenis === 'BS') {
      const m = {}; ids.forEach(function (id) { if (a[id] === 'B' || a[id] === 'S') m[id] = a[id]; });
      if (Object.keys(m).length) out[q.id] = m;
    } else if (q.jenis === 'ESSAY') { const t = String(a).slice(0, 1500); if (t.trim()) out[q.id] = t; }
  });
  return out;
}

/* Penilaian satu soal otomatis (essay dinilai guru). */
function gradeSoal_(q, ans) {
  const maks = Number(q.bobot) || 0;
  if (q.jenis === 'ESSAY') return { poin: null, maks: maks, status: 'menunggu' };
  const kosong = ans === undefined || ans === null || ans === '' || (Array.isArray(ans) && !ans.length) || (typeof ans === 'object' && !Array.isArray(ans) && !Object.keys(ans).length);
  if (kosong) return { poin: 0, maks: maks, status: 'kosong' };
  let frac = 0;
  if (q.jenis === 'PG') frac = String(ans) === String(q.kunci) ? 1 : 0;
  else if (q.jenis === 'PGK') {
    const key = q.kunci || [], sel = Array.isArray(ans) ? ans : [];
    const ok = sel.filter(function (x) { return key.indexOf(x) > -1; }).length, bad = sel.length - ok;
    frac = q.aturan === 'sebagian' ? Math.max(0, (ok - bad) / key.length) : (ok === key.length && bad === 0 ? 1 : 0);
  } else if (q.jenis === 'BS') {
    const ids = (q.opsi || []).map(function (o) { return o.id; }), key = q.kunci || {};
    const ok = ids.filter(function (id) { return ans[id] && ans[id] === key[id]; }).length;
    frac = q.aturan === 'semua' ? (ok === ids.length ? 1 : 0) : ok / ids.length;
  }
  return { poin: Math.round(frac * maks * 100) / 100, maks: maks, status: frac === 1 ? 'benar' : frac > 0 ? 'sebagian' : 'salah' };
}

/* Rekap nilai satu peserta: skor otomatis + skor essay dari guru. */
function hitung_(soal, ans, essay) {
  let auto = 0, ess = 0, maks = 0, belum = 0;
  const detail = soal.map(function (q) {
    const g = gradeSoal_(q, ans[q.id]);
    maks += g.maks;
    if (q.jenis === 'ESSAY') {
      const a = String(ans[q.id] || '').trim(), sv = essay ? essay[q.id] : undefined;
      if (sv !== undefined && sv !== null && sv !== '') { const p = Math.min(g.maks, Math.max(0, Number(sv) || 0)); ess += p; return { id: q.id, poin: p, maks: g.maks, status: 'dinilai' }; }
      if (a) { belum++; return { id: q.id, poin: null, maks: g.maks, status: 'menunggu' }; }
      return { id: q.id, poin: 0, maks: g.maks, status: 'kosong' };
    }
    auto += g.poin;
    return { id: q.id, poin: g.poin, maks: g.maks, status: g.status };
  });
  const r2 = function (x) { return Math.round(x * 100) / 100; }, total = r2(auto + ess);
  return { auto: r2(auto), essay: r2(ess), total: total, maks: r2(maks), nilai: maks ? Math.round(total / maks * 1000) / 10 : 0, belum: belum, detail: detail };
}

/* Soal yang dikirim ke peserta: tanpa kunci, tanpa pedoman; urutan diacak bila diminta. */
function soalUntukSiswa_(soal, seed, op) {
  let list = soal.slice();
  if (op && op.acakSoal) list = shuffle_(list, seed + '|q');
  return list.map(function (q) {
    let opsi = q.opsi || [];
    if (op && op.acakOpsi && (q.jenis === 'PG' || q.jenis === 'PGK')) opsi = shuffle_(opsi, seed + '|o|' + q.id);
    return { id: q.id, jenis: q.jenis, teks: q.teks, gambar: q.gambar, opsi: opsi.map(function (o) { return { id: o.id, t: o.t, g: o.g }; }) };
  });
}


/* ===================== SOAL INTERNAL: PENYIMPANAN ===================== */
const SOAL_HEADERS = ['ID Soal', 'ID Ujian', 'Urutan', 'Jenis', 'Pertanyaan', 'Gambar', 'Opsi', 'Kunci', 'Bobot', 'Aturan', 'Diubah'];
const Q = { ID: 0, UJIAN: 1, URUT: 2, JENIS: 3, TEKS: 4, IMG: 5, OPSI: 6, KUNCI: 7, BOBOT: 8, ATURAN: 9, UBAH: 10 };
const JAWAB_HEADERS = ['ID Ujian', 'Nomor Peserta', 'Nama', 'Kelas', 'Sekolah', 'Jawaban', 'Ragu', 'Skor Essay', 'Diperbarui'];
const J = { UJIAN: 0, NO: 1, NAMA: 2, KELAS: 3, SEKOLAH: 4, JAWAB: 5, RAGU: 6, ESSAY: 7, UBAH: 8 };
const IMG_HEADERS = ['ID Gambar', 'Jenis', 'Bagian', 'Data', 'Dibuat'];
const IMG_CHUNK = 45000;                        // batas sel Google Sheets 50.000 karakter

/* Semua sel di sheet ini berformat teks: isi soal/jawaban yang diawali "=" tidak pernah dijalankan sebagai rumus. */
function ensureSheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#E0E7FF');
    sh.setFrozenRows(1);
  }
  return sh;
}
function cachePutBig_(key, str, ttl) {
  const n = Math.ceil(str.length / 90000), o = {};
  for (let i = 0; i < n; i++) o[key + '#' + i] = str.slice(i * 90000, (i + 1) * 90000);
  o[key + '#n'] = String(n);
  try { CacheService.getScriptCache().putAll(o, ttl); } catch (e) { /* terlalu besar: abaikan */ }
}
function cacheGetBig_(key) {
  const c = CacheService.getScriptCache(), n = Number(c.get(key + '#n'));
  if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(key + '#' + i);
  const o = c.getAll(keys); let s = '';
  for (let i = 0; i < n; i++) { const p = o[key + '#' + i]; if (p === undefined || p === null) return null; s += p; }
  return s;
}
function cacheDelBig_(key) {
  const c = CacheService.getScriptCache(), n = Number(c.get(key + '#n')) || 0, ks = [key + '#n'];
  for (let i = 0; i < n; i++) ks.push(key + '#' + i);
  c.removeAll(ks);
}

function soalList_(idUjian) {
  const ck = 'soal_' + idUjian, hit = cacheGetBig_(ck);
  if (hit) return JSON.parse(hit);
  const sh = ensureSheet_(SH.SOAL, SOAL_HEADERS);
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, SOAL_HEADERS.length).getValues() : [];
  const list = v.filter(function (r) { return str_(r[Q.UJIAN]) === idUjian; }).map(function (r) {
    return { id: str_(r[Q.ID]), jenis: str_(r[Q.JENIS]), urut: Number(r[Q.URUT]) || 0, teks: String(r[Q.TEKS] || ''), gambar: str_(r[Q.IMG]),
      opsi: parseJson_(r[Q.OPSI], []), kunci: parseJson_(r[Q.KUNCI], null), bobot: Number(r[Q.BOBOT]) || 1, aturan: str_(r[Q.ATURAN]) };
  }).sort(function (a, b) { return a.urut - b.urut; });
  cachePutBig_(ck, JSON.stringify(list), 600);
  return list;
}
function bustSoal_(idUjian) { cacheDelBig_('soal_' + idUjian); CacheService.getScriptCache().remove('soalCount'); }
function soalCount_() {
  const c = CacheService.getScriptCache(), hit = c.get('soalCount');
  if (hit) return JSON.parse(hit);
  const sh = ss_().getSheetByName(SH.SOAL), m = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, Q.UJIAN + 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { const k = str_(r[0]); if (k) m[k] = (m[k] || 0) + 1; });
  c.put('soalCount', JSON.stringify(m), 60);
  return m;
}
function examGuard_(u, idUjian) {
  const row = examRows_().find(function (r) { return str_(r.idujian) === str_(idUjian); });
  if (!row) throw new Error('Ujian tidak ditemukan.');
  if (!allowed_(u, row.namasekolah)) throw new Error('Akses ditolak.');
  return row;
}
function soalPayload_(idUjian) {
  const row = examRows_().find(function (r) { return str_(r.idujian) === idUjian; }), e = row ? examPublic_(row) : {};
  return { soal: soalList_(idUjian), ujian: { id: idUjian, nama: e.nama || '', sumber: e.sumber, opsiUjian: e.opsiUjian } };
}

/* ===================== SOAL INTERNAL: AKSI GURU ===================== */
function adminGetSoal_(req) {
  const u = session_(req.token, 'admin'); examGuard_(u, req.idUjian);
  const p = soalPayload_(str_(req.idUjian)); p.ok = true; return p;
}
function adminSaveSoal_(req) {
  const u = session_(req.token, 'admin'), idUjian = str_(req.idUjian); examGuard_(u, idUjian);
  const src = req.soal || {}, c = cleanSoal_(src);
  let id = str_(src.id);
  if (id && !/^[A-Za-z0-9_-]{3,40}$/.test(id)) throw new Error('ID soal tidak valid.');
  return withLock_(function () {
    const sh = ensureSheet_(SH.SOAL, SOAL_HEADERS), last = sh.getLastRow();
    const v = last > 1 ? sh.getRange(2, 1, last - 1, SOAL_HEADERS.length).getValues() : [];
    let idx = -1, maxUrut = 0;
    v.forEach(function (r, i) {
      if (str_(r[Q.UJIAN]) === idUjian) maxUrut = Math.max(maxUrut, Number(r[Q.URUT]) || 0);
      if (id && str_(r[Q.ID]) === id) { if (str_(r[Q.UJIAN]) !== idUjian) throw new Error('ID soal sudah dipakai.'); idx = i; }
    });
    if (!id) id = 'Q-' + Date.now().toString(36).toUpperCase() + Utilities.getUuid().slice(0, 3).toUpperCase();
    const urut = idx >= 0 ? (Number(v[idx][Q.URUT]) || maxUrut) : maxUrut + 1;
    const line = [id, idUjian, urut, c.jenis, c.teks, c.gambar, JSON.stringify(c.opsi), JSON.stringify(c.kunci), c.bobot, c.aturan, stamp_()];
    sh.getRange(idx >= 0 ? idx + 2 : last + 1, 1, 1, line.length).setValues([line]);
    bustSoal_(idUjian);
    const p = soalPayload_(idUjian); p.ok = true; p.id = id; return p;
  });
}
function adminDeleteSoal_(req) {
  const u = session_(req.token, 'admin'), idUjian = str_(req.idUjian); examGuard_(u, idUjian);
  return withLock_(function () {
    const sh = ensureSheet_(SH.SOAL, SOAL_HEADERS), last = sh.getLastRow();
    const v = last > 1 ? sh.getRange(2, 1, last - 1, SOAL_HEADERS.length).getValues() : [];
    for (let i = 0; i < v.length; i++) if (str_(v[i][Q.ID]) === str_(req.id) && str_(v[i][Q.UJIAN]) === idUjian) { sh.deleteRow(i + 2); break; }
    bustSoal_(idUjian);
    const p = soalPayload_(idUjian); p.ok = true; return p;                  // idempoten
  });
}
function adminReorderSoal_(req) {
  const u = session_(req.token, 'admin'), idUjian = str_(req.idUjian); examGuard_(u, idUjian);
  const order = (Array.isArray(req.ids) ? req.ids : []).map(str_);
  return withLock_(function () {
    const sh = ensureSheet_(SH.SOAL, SOAL_HEADERS), last = sh.getLastRow();
    if (last > 1) {
      const v = sh.getRange(2, 1, last - 1, SOAL_HEADERS.length).getValues(), urut = v.map(function (r) { return [r[Q.URUT]]; });
      const mine = []; v.forEach(function (r, i) { if (str_(r[Q.UJIAN]) === idUjian) mine.push({ i: i, id: str_(r[Q.ID]), u: Number(r[Q.URUT]) || 0 }); });
      mine.sort(function (a, b) { const x = order.indexOf(a.id), y = order.indexOf(b.id); return (x < 0 ? 1e6 + a.u : x) - (y < 0 ? 1e6 + b.u : y); });
      mine.forEach(function (m, k) { urut[m.i] = [k + 1]; });
      sh.getRange(2, Q.URUT + 1, urut.length, 1).setValues(urut);
    }
    bustSoal_(idUjian);
    const p = soalPayload_(idUjian); p.ok = true; return p;
  });
}

/* ===================== GAMBAR (disimpan di sheet "Gambar", tidak pernah dibuka untuk umum) ===================== */
function adminUploadImage_(req) {
  session_(req.token, 'admin');
  const mime = str_(req.mime).toLowerCase(), data = String(req.data || '');
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mime) < 0) throw new Error('Format gambar harus JPG, PNG, atau WebP.');
  if (!/^[A-Za-z0-9+\/=]+$/.test(data)) throw new Error('Data gambar tidak valid.');
  if (data.length > 420000) throw new Error('Gambar terlalu besar. Perkecil gambar lalu coba lagi.');
  const id = 'IMG-' + Date.now().toString(36).toUpperCase() + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
  const rows = [];
  for (let i = 0, p = 0; i < data.length; i += IMG_CHUNK, p++) rows.push([id, mime, p, data.slice(i, i + IMG_CHUNK), stamp_()]);
  withLock_(function () {
    const sh = ensureSheet_(SH.IMG, IMG_HEADERS);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, IMG_HEADERS.length).setValues(rows);
  });
  return { ok: true, id: id };
}
function imageData_(id) {
  const ck = 'img_' + id, hit = cacheGetBig_(ck);
  if (hit) return JSON.parse(hit);
  const sh = ss_().getSheetByName(SH.IMG);
  if (!sh || sh.getLastRow() < 2) return null;
  const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues(), at = [];
  col.forEach(function (r, i) { if (str_(r[0]) === id) at.push(i + 2); });
  if (!at.length) return null;
  const first = at[0], vals = sh.getRange(first, 1, at[at.length - 1] - first + 1, 4).getValues()
    .filter(function (r) { return str_(r[0]) === id; }).sort(function (a, b) { return Number(a[2]) - Number(b[2]); });
  const obj = { mime: str_(vals[0][1]), data: vals.map(function (r) { return String(r[3]); }).join('') };
  cachePutBig_(ck, JSON.stringify(obj), 21600);
  return obj;
}
function imageIdsOf_(soal) {
  const m = {};
  soal.forEach(function (q) { if (q.gambar) m[q.gambar] = 1; (q.opsi || []).forEach(function (o) { if (o.g) m[o.g] = 1; }); });
  return m;
}
/* Guru: gambar apa pun. Peserta: hanya gambar milik ujian yang sudah ia mulai. */
function getImages_(req) {
  const raw = req.token && CacheService.getScriptCache().get('tok_' + req.token);
  if (!raw) throw new Error('Sesi berakhir. Silakan login kembali.');
  const s = JSON.parse(raw);
  const ids = (Array.isArray(req.ids) ? req.ids : []).map(str_).filter(function (x) { return IMG_RE_.test(x); }).slice(0, 8);
  let okIds = null;
  if (s.role === 'siswa') {
    const idUjian = str_(req.idUjian), row = examRows_().find(function (r) { return str_(r.idujian) === idUjian; });
    if (!row || !same_(row.namasekolah, s.sekolah) || !findLog_(logSheet_(), idUjian, s.nomor)) throw new Error('Akses ditolak.');
    okIds = imageIdsOf_(soalList_(idUjian));
  } else if (s.role !== 'admin') throw new Error('Akses ditolak.');
  const out = {};
  ids.forEach(function (id) {
    if (okIds && !okIds[id]) return;
    const im = imageData_(id); if (im) out[id] = 'data:' + im.mime + ';base64,' + im.data;
  });
  return { ok: true, images: out };
}

/* ===================== JAWABAN PESERTA ===================== */
function ansSheet_() { return ensureSheet_(SH.JAWAB, JAWAB_HEADERS); }
function findAns_(sh, idUjian, no) {
  const c = CacheService.getScriptCache(), key = 'arow_' + idUjian + '_' + no, w = JAWAB_HEADERS.length, hint = Number(c.get(key)) || 0;
  if (hint > 1 && hint <= sh.getLastRow()) {
    const vals = sh.getRange(hint, 1, 1, w).getValues()[0];
    if (str_(vals[J.UJIAN]) === idUjian && str_(vals[J.NO]) === no) return { row: hint, vals: vals };
  }
  const last = sh.getLastRow(); if (last < 2) return null;
  const v = sh.getRange(2, 1, last - 1, w).getValues();
  for (let i = 0; i < v.length; i++)
    if (str_(v[i][J.UJIAN]) === idUjian && str_(v[i][J.NO]) === no) { c.put(key, String(i + 2), 21600); return { row: i + 2, vals: v[i] }; }
  return null;
}
function ensureAns_(sh, idUjian, u) {                                       // panggil di dalam lock
  let a = findAns_(sh, idUjian, u.nomor);
  if (!a) {
    const line = [idUjian, u.nomor, u.nama, u.kelas, u.sekolah, '{}', '[]', '{}', stamp_()];
    sh.getRange(sh.getLastRow() + 1, 1, 1, line.length).setValues([line]);
    a = findAns_(sh, idUjian, u.nomor);
  }
  return a;
}
function simpanJawaban_(idUjian, u, jawaban, ragu, locked) {
  const soal = soalList_(idUjian), json = JSON.stringify(cleanAnswers_(soal, jawaban));
  if (json.length > 45000) throw new Error('Jawaban terlalu panjang. Persingkat jawaban essay.');
  const rg = (Array.isArray(ragu) ? ragu : []).map(String).filter(function (id) { return soal.some(function (q) { return q.id === id; }); }).slice(0, 300);
  const sh = ansSheet_();
  let a = findAns_(sh, idUjian, u.nomor);
  if (!a) a = locked ? ensureAns_(sh, idUjian, u) : withLock_(function () { return ensureAns_(sh, idUjian, u); });
  sh.getRange(a.row, J.JAWAB + 1, 1, 2).setValues([[json, JSON.stringify(rg)]]);
  sh.getRange(a.row, J.UBAH + 1).setValue(stamp_());
}
function saveAnswers_(req) {
  const u = session_(req.token, 'siswa'), idUjian = str_(req.idUjian);
  const row = examRows_().find(function (r) { return str_(r.idujian) === idUjian; });
  if (!row || !same_(row.namasekolah, u.sekolah)) throw new Error('Ujian tidak ditemukan.');
  const ex = examPublic_(row), sh = logSheet_(), lg = findLog_(sh, idUjian, u.nomor);
  if (!lg) throw new Error('Sesi ujian tidak ditemukan.');
  const st = logState_(lg.vals, ex.selesaiMs);
  if (ex.sumber !== 'internal' || st.status !== 'Sedang Mengerjakan' || st.remainingSec <= 0) return { ok: true, status: st.status, disimpan: false };
  simpanJawaban_(idUjian, u, req.jawaban, req.ragu, false);
  sh.getRange(lg.row, L.SEEN + 1).setValue(new Date());
  return { ok: true, status: st.status, disimpan: true, jam: Date.now() };
}
function ringkasHasil_(idUjian, nomor) {                                    // ringkasan untuk peserta (tanpa essay)
  const soal = soalList_(idUjian), a = findAns_(ansSheet_(), idUjian, nomor), ans = a ? parseJson_(a.vals[J.JAWAB], {}) : {};
  const ess = soal.filter(function (q) { return q.jenis === 'ESSAY'; }), h = hitung_(soal.filter(function (q) { return q.jenis !== 'ESSAY'; }), ans, {});
  return { poin: h.auto, maks: h.maks, nilai: h.nilai, adaEssay: ess.length > 0, maksEssay: ess.reduce(function (s, q) { return s + q.bobot; }, 0) };
}

/* ===================== HASIL & PENILAIAN ESSAY (guru) ===================== */
function ansRowsOf_(idUjian) {
  const sh = ansSheet_(), last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, JAWAB_HEADERS.length).getValues().filter(function (r) { return str_(r[J.UJIAN]) === idUjian; }) : [];
}
function adminGetHasil_(req) {
  const u = session_(req.token, 'admin'), idUjian = str_(req.idUjian), row = examGuard_(u, idUjian), soal = soalList_(idUjian);
  const lsh = logSheet_(), lv = lsh.getLastRow() > 1 ? lsh.getRange(2, 1, lsh.getLastRow() - 1, LOG_HEADERS.length).getValues() : [], logBy = {};
  lv.forEach(function (r) { if (str_(r[L.ID]) === idUjian) logBy[str_(r[L.NO])] = r; });
  const stat = {}; soal.forEach(function (q) { stat[q.id] = { benar: 0, sebagian: 0, salah: 0, kosong: 0, menunggu: 0, dinilai: 0, n: 0 }; });
  const rows = ansRowsOf_(idUjian).map(function (r) {
    const no = str_(r[J.NO]), h = hitung_(soal, parseJson_(r[J.JAWAB], {}), parseJson_(r[J.ESSAY], {})), lg = logBy[no];
    h.detail.forEach(function (d) { const s = stat[d.id]; if (s) { s.n++; if (s[d.status] !== undefined) s[d.status]++; } });
    return { nomor: no, nama: str_(r[J.NAMA]), kelas: str_(r[J.KELAS]), sekolah: str_(r[J.SEKOLAH]), status: lg ? str_(lg[L.STATUS]) : '-',
      pelanggaran: lg ? Number(lg[L.VIO]) || 0 : 0, auto: h.auto, essay: h.essay, belum: h.belum, total: h.total, maks: h.maks, nilai: h.nilai,
      poinSoal: h.detail.map(function (d) { return d.poin; }) };
  });
  return { ok: true, ujian: { id: idUjian, nama: str_(row.namaujian) },
    soal: soal.map(function (q, i) { return { id: q.id, no: i + 1, jenis: q.jenis, teks: q.teks.slice(0, 160), bobot: q.bobot, stat: stat[q.id] }; }), rows: rows };
}
function adminGetJawaban_(req) {
  const u = session_(req.token, 'admin'), idUjian = str_(req.idUjian), nomor = str_(req.nomor); examGuard_(u, idUjian);
  const soal = soalList_(idUjian), a = findAns_(ansSheet_(), idUjian, nomor);
  if (!a) throw new Error('Jawaban peserta tidak ditemukan.');
  const ans = parseJson_(a.vals[J.JAWAB], {}), es = parseJson_(a.vals[J.ESSAY], {}), h = hitung_(soal, ans, es);
  return { ok: true, peserta: { nomor: nomor, nama: str_(a.vals[J.NAMA]), kelas: str_(a.vals[J.KELAS]) },
    ringkas: { auto: h.auto, essay: h.essay, total: h.total, maks: h.maks, nilai: h.nilai, belum: h.belum },
    detail: soal.map(function (q, i) {
      return { no: i + 1, id: q.id, jenis: q.jenis, teks: q.teks, gambar: q.gambar, opsi: q.opsi, kunci: q.kunci, aturan: q.aturan, bobot: q.bobot,
        jawaban: ans[q.id] === undefined ? null : ans[q.id], poin: h.detail[i].poin, status: h.detail[i].status, skorEssay: es[q.id] === undefined ? null : es[q.id] };
    }) };
}
function adminNilaiEssay_(req) {
  const u = session_(req.token, 'admin'), idUjian = str_(req.idUjian), nomor = str_(req.nomor); examGuard_(u, idUjian);
  const soal = soalList_(idUjian), skor = req.skor && typeof req.skor === 'object' ? req.skor : {};
  return withLock_(function () {
    const sh = ansSheet_(), a = findAns_(sh, idUjian, nomor);
    if (!a) throw new Error('Jawaban peserta tidak ditemukan.');
    const es = parseJson_(a.vals[J.ESSAY], {});
    Object.keys(skor).forEach(function (qid) {
      const q = soal.find(function (x) { return x.id === qid; });
      if (!q || q.jenis !== 'ESSAY') throw new Error('Soal essay tidak ditemukan.');
      if (skor[qid] === '' || skor[qid] === null) { delete es[qid]; return; }
      const n = Number(skor[qid]);
      if (isNaN(n) || n < 0 || n > q.bobot) throw new Error('Skor essay harus antara 0 dan ' + q.bobot + '.');
      es[qid] = Math.round(n * 100) / 100;
    });
    sh.getRange(a.row, J.ESSAY + 1).setValue(JSON.stringify(es));
    const h = hitung_(soal, parseJson_(a.vals[J.JAWAB], {}), es);
    return { ok: true, ringkas: { auto: h.auto, essay: h.essay, total: h.total, maks: h.maks, nilai: h.nilai, belum: h.belum } };
  });
}
