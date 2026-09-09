const SPREADSHEET_ID = '11m-xtfldUCVuJQsB9-FUCGMK8V4d5QpekZPOndANWOQ';
const SHEET_NAME = 'Bookings';
const TZ = 'Asia/Jakarta';
const BOOKING_HOUR_START = 10;
const BOOKING_HOUR_END = 22;
const SLOT_MINUTES = 30;
const DEFAULT_BARBERS = ['Rezky', 'Iqbal'];
const CAPSTER_SHEET = 'Capsters';
const PROMO_SHEET = 'Promos';

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
    if (api === 'getSettings') {
      return jsonOutput_({ok:true, data:getSettings()});
    }
    if (api === 'getBookingMonths') {
      return jsonOutput_({ok:true, data:getBookingMonths()});
    }
    if (api === 'validatePromo') {
      return jsonOutput_({ok:true, data:validatePromo(String(params.code || ''))});
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
      case 'getSettings':
        result = getSettings();
        break;
      case 'getBookingMonths':
        result = getBookingMonths();
        break;
      case 'validatePromo':
        result = validatePromo(String(p.code || ''));
        break;
      case 'saveCapster':
        result = saveCapster(p);
        break;
      case 'toggleCapster':
        result = toggleCapster(p.name, p.active);
        break;
      case 'createPromo':
        result = createPromo(p);
        break;
      case 'togglePromo':
        result = togglePromo(p.code, p.active);
        break;
      case 'deletePromo':
        result = deletePromo(p.code);
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

function getSheet_(dateValue) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = getMonthSheetName_(dateValue);
  let sh = ss.getSheetByName(sheetName);

  // Buat tab bulan secara otomatis jika belum ada.
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    initializeBookingSheet_(sh);

    // Jika masih ada sheet legacy "Bookings", salin booking bulan ini
    // agar data lama tidak hilang ketika sistem pertama kali dipindahkan.
    migrateLegacyBookingsForMonth_(ss, sh, dateValue);
  } else {
    initializeBookingSheet_(sh);
  }
  return sh;
}

function getMonthSheetName_(dateValue) {
  let d;
  if (dateValue instanceof Date) {
    d = new Date(dateValue.getTime());
  } else {
    const s = String(dateValue || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const parts = s.split('-').map(Number);
      d = new Date(parts[0], parts[1] - 1, 1);
    } else {
      d = new Date();
    }
  }
  const months = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  return months[d.getMonth()] + ' ' + d.getFullYear();
}

function initializeBookingSheet_(sh) {
  const headers = [
    'ID','Timestamp','Name','WA','Note','Reminder','Promo',
    'Category','Service','Price','Duration','Barber','Date','Time','Status',
    'Promo Discount','Total Price'
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
}

function migrateLegacyBookingsForMonth_(ss, targetSheet, dateValue) {
  const legacy = ss.getSheetByName(SHEET_NAME);
  if (!legacy || legacy.getName() === targetSheet.getName() || legacy.getLastRow() < 2) return;

  const targetMonth = getMonthSheetName_(dateValue);
  const rows = legacy.getRange(2, 1, legacy.getLastRow() - 1, 17).getValues();
  const matching = rows.filter(r => getMonthSheetName_(r[12]) === targetMonth && r[0]);
  if (!matching.length) return;

  const existingIds = targetSheet.getLastRow() < 2 ? new Set() :
    new Set(targetSheet.getRange(2,1,targetSheet.getLastRow()-1,1).getValues().flat().map(String));
  const fresh = matching.filter(r => !existingIds.has(String(r[0])));
  if (fresh.length) targetSheet.getRange(targetSheet.getLastRow()+1,1,fresh.length,17).setValues(fresh);
}

function getAllBookingSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Pastikan tab bulan berjalan selalu ada.
  getSheet_(new Date());

  // Sertakan sheet legacy `Bookings` agar data lama yang sudah ada di
  // Spreadsheet tetap terbaca oleh website/admin.
  const sheets = ss.getSheets().filter(sh =>
    (/^.+ \d{4}$/.test(sh.getName()) || sh.getName() === SHEET_NAME) &&
    sh.getName() !== CAPSTER_SHEET &&
    sh.getName() !== PROMO_SHEET
  );
  return sheets;
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script untuk membuat trigger bulanan.
 * Setiap tanggal 1 sekitar pukul 00:00-01:00, tab bulan baru dibuat otomatis.
 * Sistem juga tetap membuat tab secara lazy saat ada booking/API, jadi aman jika trigger terlambat.
 */
function setupMonthlySheetTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'createCurrentMonthSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('createCurrentMonthSheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
  createCurrentMonthSheet();
  return 'Trigger bulanan aktif.';
}

function createCurrentMonthSheet() {
  const sh = getSheet_(new Date());
  sh.autoResizeColumns(1, 17);
  return 'Tab bulan aktif: ' + sh.getName();
}

function setupSheet() {
  const sh = getSheet_();
  sh.autoResizeColumns(1, 17);
  return 'Sheet siap: ' + sh.getName();
}

function createBooking(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!p || !p.name || !p.wa || !p.service || !p.barber || !p.date || !p.time) {
      throw new Error('Data booking belum lengkap.');
    }
    const settings = getSettings();
    const cap = settings.capsters.find(x => x.name === String(p.barber).trim() && x.active);
    if (!cap) throw new Error('Capster tersebut sedang tidak tersedia. Silakan pilih capster lain.');
    const wa = normalizeWa_(p.wa);
    if (!/^08\d{8,12}$/.test(wa)) throw new Error('Nomor WhatsApp tidak valid.');
    if (!isValidTime_(p.time)) throw new Error('Jam booking tidak valid.');

    const sh = getSheet_(p.date);
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
    const additionalFee = Number(p.additionalFee || 0);
    const baseTotal = price + additionalFee;
    const promoResult = validatePromo(String(p.promo || ''));
    const discount = promoResult.valid ? calculateDiscount_(promoResult, baseTotal) : 0;
    const totalPrice = Math.max(0, baseTotal - discount);
    if (promoResult.valid) consumePromo_(promoResult.code);
    sh.appendRow([
      id, new Date(), String(p.name).trim(), wa, String(p.note || '').trim(),
      p.reminder ? 'Ya' : 'Tidak', String(p.promo || '').trim().toUpperCase(),
      String(p.category || ''), String(p.service).trim(), price,
      String(p.duration || '').trim(), String(p.barber).trim(),
      p.date, p.time, 'Confirmed', discount, totalPrice
    ]);
    return {ok:true, id:id, promoDiscount:discount, totalPrice:totalPrice};
  } finally {
    lock.releaseLock();
  }
}

function getBookingMonths() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  getSheet_(new Date());
  const names = new Set();
  const monthSheets = ss.getSheets().filter(sh => /^.+ \d{4}$/.test(sh.getName()));
  monthSheets.forEach(sh => names.add(sh.getName()));

  // Tambahkan bulan dari data legacy `Bookings`, jika ada.
  const legacy = ss.getSheetByName(SHEET_NAME);
  if (legacy && legacy.getLastRow() >= 2) {
    const dates = legacy.getRange(2,13,legacy.getLastRow()-1,1).getValues();
    dates.forEach(r => {
      if (r[0]) names.add(getMonthSheetName_(r[0]));
    });
  }

  const idx = {'Januari':0,'Februari':1,'Maret':2,'April':3,'Mei':4,'Juni':5,'Juli':6,'Agustus':7,'September':8,'Oktober':9,'November':10,'Desember':11};
  return Array.from(names).sort((a,b) => {
    const pa=/^(.+) (\d{4})$/.exec(a), pb=/^(.+) (\d{4})$/.exec(b);
    const da=pa ? new Date(Number(pa[2]), idx[pa[1]] ?? 0, 1).getTime() : 0;
    const db=pb ? new Date(Number(pb[2]), idx[pb[1]] ?? 0, 1).getTime() : 0;
    return db-da;
  });
}

function getBookings() {
  const sheets = getAllBookingSheets_();
  const byId = new Map();
  sheets.forEach(sh => {
    const last = sh.getLastRow();
    if (last >= 2) {
      const rows = sh.getRange(2,1,last-1,17).getValues();
      rows.filter(r => r[0]).forEach(r => {
        const obj = rowToObject_(r);
        if (obj.id) byId.set(obj.id, obj);
      });
    }
  });
  return Array.from(byId.values()).sort((a,b) =>
    String(b.date+' '+b.time).localeCompare(String(a.date+' '+a.time))
  );
}

function getClientBookingData(date, barber) {
  const booked = {};
  const sh = getSheet_(date);
  const last = sh.getLastRow();
  if (last >= 2) {
    const rows = sh.getRange(2,1,last-1,17).getValues();
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

function findBookingLocation_(id) {
  const sheets = getAllBookingSheets_();
  for (const sh of sheets) {
    const last = sh.getLastRow();
    if (last < 2) continue;
    const ids = sh.getRange(2,1,last-1,1).getValues();
    for (let i=0;i<ids.length;i++) {
      if (String(ids[i][0]) === String(id)) return {sh:sh,row:i+2};
    }
  }
  return null;
}

function updateBooking(p) {
  if (!p || !p.id) throw new Error('ID booking tidak ditemukan.');
  const loc = findBookingLocation_(p.id);
  if (!loc) throw new Error('Booking tidak ditemukan.');

  const sh = loc.sh;
  const targetRow = loc.row;
  const target = sh.getRange(targetRow,1,1,17).getValues()[0];
  if (p.status !== undefined && !['Confirmed','Rescheduled','Cancelled','Completed'].includes(String(p.status))) {
    throw new Error('Status booking tidak valid.');
  }
  if (p.name !== undefined && !String(p.name).trim()) throw new Error('Nama customer wajib diisi.');
  if (p.service !== undefined && !String(p.service).trim()) throw new Error('Layanan wajib diisi.');
  if (p.duration !== undefined && !String(p.duration).trim()) throw new Error('Durasi wajib diisi.');
  const newDate = p.date ? String(p.date) : normalizeDate_(target[12]);
  const newTime = p.time ? String(p.time) : String(target[13] || '');
  const newBarber = p.barber ? String(p.barber).trim() : String(target[11] || '').trim();

  if (p.date || p.time || p.barber) {
    if (!newDate || !newTime || !newBarber) throw new Error('Tanggal, jam, dan capster wajib diisi.');
    if (!isValidTime_(newTime)) throw new Error('Jam booking tidak valid.');
    const destination = getSheet_(newDate);
    const conflict = findConflict_(destination, newDate, newTime, newBarber, String(p.id));
    if (conflict) throw new Error('Jam tersebut sudah terisi untuk capster ini.');
  }

  // Field yang boleh diedit dari admin. Field ID/timestamp tidak diubah.
  if (p.name !== undefined) sh.getRange(targetRow,3).setValue(String(p.name).trim());
  if (p.wa !== undefined) { const wa = normalizeWa_(p.wa); if (!/^08\d{8,12}$/.test(wa)) throw new Error('Nomor WhatsApp tidak valid.'); sh.getRange(targetRow,4).setValue(wa); }
  if (p.note !== undefined) sh.getRange(targetRow,5).setValue(String(p.note));
  if (p.reminder !== undefined) sh.getRange(targetRow,6).setValue(p.reminder ? 'Ya' : 'Tidak');
  if (p.category !== undefined) sh.getRange(targetRow,8).setValue(String(p.category));
  if (p.service !== undefined) sh.getRange(targetRow,9).setValue(String(p.service));
  if (p.price !== undefined) sh.getRange(targetRow,10).setValue(parsePrice_(p.price));
  if (p.duration !== undefined) sh.getRange(targetRow,11).setValue(String(p.duration));
  if (p.barber !== undefined) sh.getRange(targetRow,12).setValue(newBarber);
  if (p.date !== undefined) sh.getRange(targetRow,13).setValue(newDate);
  if (p.time !== undefined) sh.getRange(targetRow,14).setValue(newTime);
  if (p.status !== undefined) sh.getRange(targetRow,15).setValue(String(p.status));
  if (p.promo !== undefined) sh.getRange(targetRow,7).setValue(String(p.promo).trim().toUpperCase());
  if (p.promoDiscount !== undefined) sh.getRange(targetRow,16).setValue(Number(p.promoDiscount || 0));
  if (p.totalPrice !== undefined) sh.getRange(targetRow,17).setValue(Number(p.totalPrice || 0));

  // Jika tanggal pindah bulan atau sumbernya sheet legacy `Bookings`,
  // pindahkan record ke sheet bulan tujuan agar struktur tetap konsisten.
  const shouldMove = getMonthSheetName_(newDate) !== sh.getName();
  if (shouldMove) {
    const updated = sh.getRange(targetRow,1,1,17).getValues()[0];
    const destination = getSheet_(newDate);
    const conflict = findConflict_(destination, newDate, newTime, newBarber, String(p.id));
    if (conflict) throw new Error('Jam tersebut sudah terisi untuk capster ini.');
    destination.appendRow(updated);
    sh.deleteRow(targetRow);
  }

  const finalLoc = findBookingLocation_(p.id);
  return finalLoc ? rowToObject_(finalLoc.sh.getRange(finalLoc.row,1,1,17).getValues()[0]) : {id:String(p.id)};
}

function cancelBooking(id, reason) {
  // Cancel sekarang benar-benar menghapus booking dari Spreadsheet.
  // Data tidak hanya diubah statusnya menjadi Cancelled.
  return deleteBooking(id);
}

function deleteBooking(id) {
  const loc = findBookingLocation_(id);
  if (!loc) throw new Error('Booking tidak ditemukan.');
  loc.sh.deleteRow(loc.row);
  return {ok:true, deleted:true, id:String(id)};
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
  const loc = findBookingLocation_(id);
  if (!loc) throw new Error('Booking tidak ditemukan.');
  loc.sh.getRange(loc.row,15).setValue(status);
  if (note) loc.sh.getRange(loc.row,5).setValue(note);
}

function findConflict_(sh, date, time, barber, exceptId) {
  const last = sh.getLastRow();
  if (last < 2) return false;
  const rows = sh.getRange(2,1,last-1,17).getValues();
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
    status:String(r[14] || 'Confirmed'),
    promoDiscount:Number(r[15] || 0),
    totalPrice:Number(r[16] || 0)
  };
}


function getCapsterSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(CAPSTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CAPSTER_SHEET);
    sh.getRange(1,1,1,3).setValues([['Name','Active','Updated At']]);
    DEFAULT_BARBERS.forEach((name,i) => sh.getRange(i+2,1,1,3).setValues([[name,true,new Date()]]));
    sh.setFrozenRows(1);
  }
  return sh;
}

function getPromoSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(PROMO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROMO_SHEET);
    sh.getRange(1,1,1,8).setValues([['Code','Discount Type','Discount Value','Max Uses','Used','Active','Created At','Updated At']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getSettings() {
  const csh = getCapsterSheet_();
  const plast = csh.getLastRow();
  const capsters = plast < 2 ? [] : csh.getRange(2,1,plast-1,3).getValues()
    .filter(r => String(r[0] || '').trim())
    .map(r => ({name:String(r[0]).trim(), active:r[1] !== false && String(r[1]).toLowerCase() !== 'false'}));

  const psh = getPromoSheet_();
  const last = psh.getLastRow();
  const promos = last < 2 ? [] : psh.getRange(2,1,last-1,8).getValues()
    .filter(r => String(r[0] || '').trim())
    .map(r => ({
      code:String(r[0]).trim().toUpperCase(),
      discountType:String(r[1] || 'percent'),
      discountValue:Number(r[2] || 0),
      maxUses:Number(r[3] || 0),
      used:Number(r[4] || 0),
      active:r[5] !== false && String(r[5]).toLowerCase() !== 'false',
      createdAt:r[6] instanceof Date ? Utilities.formatDate(r[6], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[6] || ''),
      remaining:Math.max(0,Number(r[3] || 0)-Number(r[4] || 0))
    }));
  return {capsters:capsters, promos:promos};
}

function saveCapster(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('Nama capster wajib diisi.');
  const sh = getCapsterSheet_(), last = sh.getLastRow();
  for (let i=2;i<=last;i++) if (String(sh.getRange(i,1).getValue()).trim().toLowerCase() === name.toLowerCase()) {
    sh.getRange(i,2,1,2).setValues([[p.active !== false, new Date()]]);
    return getSettings();
  }
  sh.appendRow([name,p.active !== false,new Date()]);
  return getSettings();
}

function toggleCapster(name, active) {
  const sh=getCapsterSheet_(), last=sh.getLastRow();
  for(let i=2;i<=last;i++) if(String(sh.getRange(i,1).getValue()).trim()===String(name).trim()){
    sh.getRange(i,2,1,2).setValues([[active===true || String(active)==='true',new Date()]]);
    return getSettings();
  }
  throw new Error('Capster tidak ditemukan.');
}

function createPromo(p) {
  const code = String(p.code || '').trim().toUpperCase().replace(/[^A-Z0-9_-]/g,'');
  if (!code || code.length < 3) throw new Error('Kode promo minimal 3 karakter.');
  const type = p.discountType === 'fixed' ? 'fixed' : 'percent';
  const value = Number(p.discountValue || 0);
  const maxUses = Math.floor(Number(p.maxUses || 0));
  if (value <= 0) throw new Error('Potongan harus lebih dari 0.');
  if (type === 'percent' && value > 100) throw new Error('Potongan persen maksimal 100%.');
  if (maxUses < 1) throw new Error('Jumlah slot promo minimal 1.');
  const sh=getPromoSheet_(), last=sh.getLastRow();
  for(let i=2;i<=last;i++) if(String(sh.getRange(i,1).getValue()).trim().toUpperCase()===code) throw new Error('Kode promo sudah ada.');
  sh.appendRow([code,type,value,maxUses,0,true,new Date(),new Date()]);
  return getSettings();
}

function togglePromo(code, active) {
  const sh=getPromoSheet_(), last=sh.getLastRow();
  for(let i=2;i<=last;i++) if(String(sh.getRange(i,1).getValue()).trim().toUpperCase()===String(code).trim().toUpperCase()){
    sh.getRange(i,6,1,2).setValues([[active===true || String(active)==='true',new Date()]]);
    return getSettings();
  }
  throw new Error('Kode promo tidak ditemukan.');
}

function deletePromo(code) {
  const sh=getPromoSheet_(), last=sh.getLastRow();
  for(let i=2;i<=last;i++) if(String(sh.getRange(i,1).getValue()).trim().toUpperCase()===String(code).trim().toUpperCase()){
    sh.deleteRow(i);
    return getSettings();
  }
  throw new Error('Kode promo tidak ditemukan.');
}

function validatePromo(code) {
  const c=String(code || '').trim().toUpperCase();
  if (!c) return {valid:false, code:'', message:'Kode promo kosong.'};
  const sh=getPromoSheet_(), last=sh.getLastRow();
  for(let i=2;i<=last;i++){
    const r=sh.getRange(i,1,1,8).getValues()[0];
    const rowCode=String(r[0] || '').trim().toUpperCase();
    if(rowCode===c){
      const active=r[5] !== false && String(r[5]).toLowerCase() !== 'false';
      const maxUses=Number(r[3] || 0), used=Number(r[4] || 0);
      if(!active) return {valid:false,code:c,message:'Kode promo sedang tidak aktif.'};
      if(used>=maxUses) return {valid:false,code:c,message:'Slot kode promo sudah habis.'};
      return {valid:true,code:c,discountType:String(r[1] || 'percent'),discountValue:Number(r[2] || 0),maxUses:maxUses,used:used,remaining:maxUses-used,message:'Kode promo valid.'};
    }
  }
  return {valid:false,code:c,message:'Kode promo tidak ditemukan.'};
}

function calculateDiscount_(promo, total) {
  if (!promo || !promo.valid) return 0;
  if (promo.discountType === 'fixed') return Math.min(total, Math.max(0, Number(promo.discountValue || 0)));
  return Math.min(total, Math.round(total * Math.max(0,Math.min(100,Number(promo.discountValue || 0))) / 100));
}

function consumePromo_(code) {
    const sh=getPromoSheet_(), last=sh.getLastRow(), c=String(code||'').trim().toUpperCase();
    for(let i=2;i<=last;i++){
      const r=sh.getRange(i,1,1,8).getValues()[0];
      if(String(r[0]||'').trim().toUpperCase()===c){
        const active=r[5] !== false && String(r[5]).toLowerCase() !== 'false';
        const max=Number(r[3]||0), used=Number(r[4]||0);
        if(!active || used>=max) throw new Error('Slot kode promo sudah habis atau tidak aktif.');
        sh.getRange(i,5).setValue(used+1);
        sh.getRange(i,8).setValue(new Date());
        return;
      }
    }
    throw new Error('Kode promo tidak ditemukan.');
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
