const SPREADSHEET_ID = '11m-xtfldUCVuJQsB9-FUCGMK8V4d5QpekZPOndANWOQ';
const SHEET_NAME = 'Bookings';
const TZ = 'Asia/Jakarta';
const BOOKING_HOUR_START = 10;
const BOOKING_HOUR_END = 22;
const SLOT_MINUTES = 30;
const DEFAULT_BARBERS = ['Rezky', 'Iqbal'];

function doGet(e) {
  const params = (e && e.parameter) || {};
  try {
    const api = String(params.api || '').trim();
    if (api === 'getBookings') {
      return jsonOutput_({ok:true, data:getBookings()});
    }
    if (api === 'getClientBookingData') {
      return jsonOutput_({ok:true, data:getClientBookingData(String(params.date || ''), String(params.barber || ''))});
    }
    return jsonOutput_({ok:true, message:'RESO Haircut API aktif'});
  } catch (err) {
    return jsonOutput_({ok:false, error:err.message || String(err)});
  }
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const p = JSON.parse(raw);
    const action = String(p.action || '').trim();
    let result;

    switch (action) {
      case 'getBookings':
        result = getBookings();
        break;
      case 'getClientBookingData':
        result = getClientBookingData(String(p.date || ''), String(p.barber || ''));
        break;
      case 'createBooking':
        result = createBooking(p);
        break;
      case 'updateBooking':
        result = updateBooking(p);
        break;
      case 'cancelBooking':
        result = cancelBooking(p.id, p.reason);
        break;
      case 'deleteBooking':
        result = deleteBooking(p.id);
        break;
      case 'reactivateBooking':
        result = reactivateBooking(p.id);
        break;
      case 'completeBookingServer':
        result = completeBookingServer(p.id);
        break;
      default:
        throw new Error('Action tidak dikenali: ' + action);
    }

    return jsonOutput_({ok:true, data:result});
  } catch (err) {
    return jsonOutput_({ok:false, error:err.message || String(err)});
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const headers = [
    'ID','Timestamp','Name','WA','Note','Reminder','Promo',
    'Category','Service','Price','Duration','Barber','Date','Time','Status'
  ];
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,headers.length).getValues()[0];
    if (current.join('|') !== headers.join('|')) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  }
  sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange('M:M').setNumberFormat('@');
  sh.getRange('N:N').setNumberFormat('@');
  return sh;
}

function setupSheet() {
  const sh = getSheet_();
  sh.autoResizeColumns(1, 15);
  return 'Sheet siap: ' + sh.getName();
}

function createBooking(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!p || !p.name || !p.wa || !p.service || !p.barber || !p.date || !p.time) {
      throw new Error('Data booking belum lengkap.');
    }
    const wa = normalizeWa_(p.wa);
    if (!/^08\d{8,12}$/.test(wa)) throw new Error('Nomor WhatsApp tidak valid.');
    if (!isValidTime_(p.time)) throw new Error('Jam booking tidak valid.');

    const sh = getSheet_();
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const date = normalizeDate_(row[12]);
      const time = String(row[13] || '');
      const barber = String(row[11] || '');
      const status = String(row[14] || '');
      if (date === p.date && time === p.time && barber === p.barber &&
          !['Cancelled','Completed'].includes(status)) {
        throw new Error('Slot tersebut sudah dibooking. Silakan pilih jam lain.');
      }
    }

    const id = 'RS-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') +
      '-' + Math.floor(100 + Math.random() * 900);
    const price = parsePrice_(p.price);
    sh.appendRow([
      id, new Date(), String(p.name).trim(), wa, String(p.note || '').trim(),
      p.reminder ? 'Ya' : 'Tidak', String(p.promo || '').trim(),
      String(p.category || ''), String(p.service).trim(), price,
      String(p.duration || '').trim(), String(p.barber).trim(),
      p.date, p.time, 'Confirmed'
    ]);
    return {ok:true, id:id};
  } finally {
    lock.releaseLock();
  }
}

function getBookings() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2,1,last-1,15).getValues();
  return rows.filter(r => r[0]).map(rowToObject_);
}

function getClientBookingData(date, barber) {
  const booked = {};
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last >= 2) {
    const rows = sh.getRange(2,1,last-1,15).getValues();
    rows.forEach(r => {
      const d = normalizeDate_(r[12]);
      const t = String(r[13] || '');
      const b = String(r[11] || '');
      const s = String(r[14] || '');
      if (d === date && b === barber && t && !['Cancelled','Completed'].includes(s)) {
        booked[t] = true;
      }
    });
  }
  const times = [];
  for (let mins = BOOKING_HOUR_START * 60; mins < BOOKING_HOUR_END * 60; mins += SLOT_MINUTES) {
    const h = String(Math.floor(mins / 60)).padStart(2,'0');
    const m = String(mins % 60).padStart(2,'0');
    const time = h + ':' + m;
    times.push({time:time, available:!booked[time]});
  }
  return {times:times};
}

function updateBooking(p) {
  if (!p || !p.id) throw new Error('ID booking tidak ditemukan.');
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  let targetRow = -1;
  let target = null;
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]) === String(p.id)) {
      targetRow = i + 1;
      target = values[i];
      break;
    }
  }
  if (targetRow < 0) throw new Error('Booking tidak ditemukan.');
  if (p.date && p.time) {
    const conflict = findConflict_(sh, p.date, p.time, String(target[11]), String(p.id));
    if (conflict) throw new Error('Jam tersebut sudah terisi untuk capster ini.');
  }
  if (p.date) sh.getRange(targetRow,13).setValue(p.date);
  if (p.time) sh.getRange(targetRow,14).setValue(p.time);
  if (p.status) sh.getRange(targetRow,15).setValue(p.status);
  if (p.note !== undefined) sh.getRange(targetRow,5).setValue(String(p.note));
  return {ok:true};
}

function cancelBooking(id, reason) {
  // Cancel sekarang benar-benar menghapus booking dari Spreadsheet.
  // Data tidak hanya diubah statusnya menjadi Cancelled.
  return deleteBooking(id);
}

function deleteBooking(id) {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  for (let i=1; i<values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const row = i + 1;
      sh.deleteRow(row);
      return {ok:true, deleted:true, id:String(id)};
    }
  }
  throw new Error('Booking tidak ditemukan.');
}

function reactivateBooking(id) {
  setStatus_(id, 'Confirmed', 'Booking diaktifkan kembali');
  return {ok:true};
}

function completeBookingServer(id) {
  setStatus_(id, 'Completed', 'Selesai & sudah bayar');
  return {ok:true};
}

function setStatus_(id, status, note) {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]) === String(id)) {
      const row = i+1;
      sh.getRange(row,15).setValue(status);
      if (note) sh.getRange(row,5).setValue(note);
      return;
    }
  }
  throw new Error('Booking tidak ditemukan.');
}

function findConflict_(sh, date, time, barber, exceptId) {
  const last = sh.getLastRow();
  if (last < 2) return false;
  const rows = sh.getRange(2,1,last-1,15).getValues();
  return rows.some(r =>
    String(r[0]) !== String(exceptId) &&
    normalizeDate_(r[12]) === date &&
    String(r[13]) === String(time) &&
    String(r[11]) === barber &&
    !['Cancelled','Completed'].includes(String(r[14] || ''))
  );
}

function rowToObject_(r) {
  return {
    id:String(r[0] || ''),
    timestamp:r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[1] || ''),
    name:String(r[2] || ''),
    wa:String(r[3] || ''),
    note:String(r[4] || ''),
    reminder:String(r[5] || ''),
    promo:String(r[6] || ''),
    category:String(r[7] || ''),
    service:String(r[8] || ''),
    price:Number(r[9] || 0),
    duration:String(r[10] || ''),
    barber:String(r[11] || ''),
    date:normalizeDate_(r[12]),
    time:String(r[13] || ''),
    status:String(r[14] || 'Confirmed')
  };
}

function normalizeWa_(v) {
  let s = String(v || '').replace(/\D/g,'');
  if (s.indexOf('62') === 0) s = '0' + s.slice(2);
  return s;
}

function normalizeDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

function parsePrice_(v) {
  if (typeof v === 'number') return v;
  const n = String(v || '').replace(/[^\d]/g,'');
  return n ? Number(n) : 0;
}

function isValidTime_(t) {
  const m = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(String(t));
  if (!m) return false;
  const mins = Number(m[1])*60 + Number(m[2]);
  return mins >= BOOKING_HOUR_START*60 &&
         mins < BOOKING_HOUR_END*60 &&
         Number(m[2]) % SLOT_MINUTES === 0;
}
