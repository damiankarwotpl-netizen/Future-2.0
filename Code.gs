const CFG = {
  LOGIN_SEED_SHEET: 'login_seed',
  CONTACTS_SHEET: 'kontakty',
  CLOTHES_SHEET: 'ubrania_robocze',
  SUBMISSIONS_SHEET: 'form_submissions',
  ROOT_FOLDER_NAME: 'Flota_Formularze_Pracownikow',
  ADMIN_LOGIN: 'admin', // ZMIEŃ
  ADMIN_PASSWORD: 'admin123', // ZMIEŃ
  SESSION_TTL_SEC: 20 * 60,
  MAX_UPLOAD_PDF_MB: 8
};

const HEADER_SYNONYMS = {
  name: ['name','imie','imię','nombre','first_name'],
  surname: ['surname','nazwisko','apellido','last_name'],
  pesel: ['pesel'],
  plant: ['plant','zaklad','zakład','planta','workplace'],
  email: ['email','e-mail','mail'],
  phone: ['phone','telefon','telefono'],
  apartment: ['apartment','mieszkanie','apartamento'],
  hireDate: ['hiredate','hire_date','datazatrudnienia','fecha_contratacion'],
  notes: ['notes','notatki','comentarios'],
  shirt: ['shirt','koszulka','camiseta'],
  hoodie: ['hoodie','bluza','sudadera'],
  pants: ['pants','spodnie','pantalon','pantalón'],
  jacket: ['jacket','kurtka','chaqueta'],
  shoes: ['shoes','buty','zapatos'],
  workplace: ['workplace','miejscepracy']
};

function doGet(e) {
  const page = safe_(e && e.parameter && e.parameter.page).toLowerCase();
  const view = page === 'admin' ? 'Admin' : 'Index';
  const baseUrl = ScriptApp.getService().getUrl();

  const tpl = HtmlService.createTemplateFromFile(view);
  tpl.homeUrl = baseUrl;
  tpl.adminUrl = `${baseUrl}?page=admin`;

  return tpl.evaluate()
    .setTitle(page === 'admin' ? 'Panel administratora' : 'Formulario trabajador')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ======================== INIT ======================== */

function initAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const loginSeed = getOrCreateSheet_(ss, CFG.LOGIN_SEED_SHEET);
  const contacts = getOrCreateSheet_(ss, CFG.CONTACTS_SHEET);
  const clothes = getOrCreateSheet_(ss, CFG.CLOTHES_SHEET);
  const submissions = getOrCreateSheet_(ss, CFG.SUBMISSIONS_SHEET);

  ensureHeader_(loginSeed, ['name','surname','pesel','plant']);
  ensureHeader_(contacts, ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry']);
  ensureHeader_(clothes, ['name','surname','plant','shirt','hoodie','pants','jacket','shoes']);
  ensureHeader_(submissions, ['name','surname','pesel','plant','submittedAt']);

  syncSeedToCoreTables_();
}

function syncSeedToCoreTables_() {
  const seedVals = getSheet_(CFG.LOGIN_SEED_SHEET).getDataRange().getValues();
  if (seedVals.length < 2) return;
  const sh = headerMap_(seedVals[0]);

  const contacts = getSheet_(CFG.CONTACTS_SHEET);
  const clothes = getSheet_(CFG.CLOTHES_SHEET);

  for (let i = 1; i < seedVals.length; i++) {
    const name = safe_(seedVals[i][sh.name]);
    const surname = safe_(seedVals[i][sh.surname]);
    const pesel = safe_(seedVals[i][sh.pesel]);
    const plant = safe_(seedVals[i][sh.plant]);
    if (!name || !surname || !pesel || !plant) continue;

    upsertByKey_(contacts,
      ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry'],
      ['name','surname','pesel','plant'],
      { name, surname, pesel, plant, email:'', phone:'', workplace:plant, apartment:'', hireDate:'', clothesSize:'', shoesSize:'', notes:'' }
    );

    upsertByKey_(clothes,
      ['name','surname','plant','shirt','hoodie','pants','jacket','shoes'],
      ['name','surname','plant'],
      { name, surname, plant, shirt:'', hoodie:'', pants:'', jacket:'', shoes:'' }
    );
  }
}

/* ======================== LOGIN + FORM ======================== */

function loginByIdentity(identity) {
  const pesel = safe_(identity.pesel);
  if (!/^\d{11}$/.test(pesel)) throw new Error('PESEL musi mieć 11 cyfr.');

  const vals = getSheet_(CFG.CONTACTS_SHEET).getDataRange().getValues();
  if (vals.length < 2) throw new Error('Brak pracowników.');

  const h = headerMap_(vals[0]);
  const matches = [];

  for (let i = 1; i < vals.length; i++) {
    const p = safe_(vals[i][h.pesel]);
    if (p !== pesel) continue;

    const n = safe_(vals[i][h.name]);
    const s = safe_(vals[i][h.surname]);
    const pl = safe_(vals[i][h.plant] || vals[i][h.workplace]);

    matches.push({
      name:n, surname:s, pesel:p, plant:pl,
      email:safe_(vals[i][h.email]),
      phone:safe_(vals[i][h.phone]),
      apartment:safe_(vals[i][h.apartment]),
      hireDate:safe_(vals[i][h.hireDate]),
      notes:safe_(vals[i][h.notes]),
      bankAccount:safe_(vals[i][h.bankAccount]),
      birthDate:safe_(vals[i][h.birthDate]),
      passportNumber:safe_(vals[i][h.passportNumber]),
      passportExpiry:safe_(vals[i][h.passportExpiry]),
      arrivalDate:safe_(vals[i][h.arrivalDate]),
      firstWorkDate:safe_(vals[i][h.firstWorkDate]),
      intlDrivingLicense:safe_(vals[i][h.intlDrivingLicense]),
      intlDrivingLicenseExpiry:safe_(vals[i][h.intlDrivingLicenseExpiry])
    });
  }

  if (!matches.length) throw new Error('Nie znaleziono pracownika dla podanego PESEL.');
  if (matches.length > 1) throw new Error('Wykryto więcej niż jednego pracownika z tym PESEL. Skontaktuj się z administratorem.');

  const found = matches[0];

  const clothes = getClothesData_(found.name, found.surname, found.plant);
  const phone = parsePhone_(found.phone);

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    `sess:${token}`,
    JSON.stringify({ name:found.name, surname:found.surname, pesel:found.pesel, plant:found.plant }),
    CFG.SESSION_TTL_SEC
  );

  return {
    ok:true,
    token,
    employee:{ ...found, phonePrefix:phone.prefix, phoneNumber:phone.number, ...clothes }
  };
}

function saveEmployeeForm(payload) {
  if (!payload || !payload.token) throw new Error('Sesja nieważna.');
  const sessRaw = CacheService.getScriptCache().get(`sess:${payload.token}`);
  if (!sessRaw) throw new Error('Sesja wygasła.');
  const sess = JSON.parse(sessRaw);

  const name = sess.name, surname = sess.surname, pesel = sess.pesel, plant = sess.plant;
  const phonePrefix = safe_(payload.phonePrefix || '+48');
  if (!['+48', '+57'].includes(phonePrefix)) throw new Error('Kierunkowy: +48 albo +57');

  upsertByKey_(getSheet_(CFG.CONTACTS_SHEET),
    ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry'],
    ['name','surname','pesel','plant'],
    {
      name, surname, pesel, plant,
      email: safe_(payload.email),
      phone: `${phonePrefix} ${safe_(payload.phoneNumber)}`.trim(),
      workplace: plant,
      apartment: safe_(payload.apartment),
      hireDate: safe_(payload.hireDate),
      clothesSize: '',
      shoesSize: safe_(payload.shoes),
      notes: safe_(payload.notes),
      bankAccount: safe_(payload.bankAccount),
      birthDate: safe_(payload.birthDate),
      passportNumber: safe_(payload.passportNumber),
      passportExpiry: safe_(payload.passportExpiry),
      arrivalDate: safe_(payload.arrivalDate),
      firstWorkDate: safe_(payload.firstWorkDate),
      intlDrivingLicense: safe_(payload.intlDrivingLicense),
      intlDrivingLicenseExpiry: safe_(payload.intlDrivingLicenseExpiry)
    }
  );

  upsertByKey_(getSheet_(CFG.CLOTHES_SHEET),
    ['name','surname','plant','shirt','hoodie','pants','jacket','shoes'],
    ['name','surname','plant'],
    {
      name, surname, plant,
      shirt: safe_(payload.shirt),
      hoodie: safe_(payload.hoodie),
      pants: safe_(payload.pants),
      jacket: safe_(payload.jacket),
      shoes: safe_(payload.shoes)
    }
  );

  markSubmission_(name, surname, pesel, plant);

  // Drive: folder per zakład + pracownik
  const folder = getEmployeeFolder_(plant, name, surname, pesel);
  upsertTextFile_(
    folder,
    `${sanitizeFilePart_(name)}_${sanitizeFilePart_(surname)}_formularz.json`,
    JSON.stringify({
      name,surname,pesel,plant,
      email:safe_(payload.email),
      phone:`${phonePrefix} ${safe_(payload.phoneNumber)}`.trim(),
      apartment:safe_(payload.apartment),
      hireDate:safe_(payload.hireDate),
      bankAccount:safe_(payload.bankAccount),
      birthDate:safe_(payload.birthDate),
      passportNumber:safe_(payload.passportNumber),
      passportExpiry:safe_(payload.passportExpiry),
      arrivalDate:safe_(payload.arrivalDate),
      firstWorkDate:safe_(payload.firstWorkDate),
      intlDrivingLicense:safe_(payload.intlDrivingLicense),
      intlDrivingLicenseExpiry:safe_(payload.intlDrivingLicenseExpiry),
      shirt:safe_(payload.shirt), hoodie:safe_(payload.hoodie), pants:safe_(payload.pants), jacket:safe_(payload.jacket), shoes:safe_(payload.shoes),
      notes:safe_(payload.notes),
      updatedAt:new Date().toISOString()
    }, null, 2)
  );

  if (payload.attachmentBase64 && payload.attachmentName) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(payload.attachmentBase64),
      'application/pdf',
      safe_(payload.attachmentName) || 'upload.pdf'
    );
    if (blob.getBytes().length > CFG.MAX_UPLOAD_PDF_MB * 1024 * 1024) {
      throw new Error(`PDF większy niż ${CFG.MAX_UPLOAD_PDF_MB} MB`);
    }
    folder.createFile(blob);
  }

  if (payload.drivingLicensePhotoBase64 && payload.drivingLicensePhotoName) {
    const imgMime = safe_(payload.drivingLicensePhotoMimeType) || 'image/jpeg';
    const imgBlob = Utilities.newBlob(
      Utilities.base64Decode(payload.drivingLicensePhotoBase64),
      imgMime,
      safe_(payload.drivingLicensePhotoName) || 'prawo_jazdy.jpg'
    );
    folder.createFile(imgBlob);
  }

  return { ok:true, message:'Zapis przebiegł pomyślnie, dziękujemy.' };
}


function adminLogin(login, password) {
  if (safe_(login) !== CFG.ADMIN_LOGIN || safe_(password) !== CFG.ADMIN_PASSWORD) {
    throw new Error('Błędny login lub hasło administratora.');
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(`adminsess:${token}`, '1', CFG.SESSION_TTL_SEC);
  return { ok:true, token };
}

function assertAdmin_(token) {
  const t = safe_(token);
  if (!t) throw new Error('Brak tokenu administratora.');
  const ok = CacheService.getScriptCache().get(`adminsess:${t}`);
  if (!ok) throw new Error('Sesja administratora wygasła. Zaloguj się ponownie.');
}

/* ======================== ADMIN STATS ======================== */

function getAdminStats(adminToken) {
  assertAdmin_(adminToken);
  const cVals = getSheet_(CFG.CONTACTS_SHEET).getDataRange().getValues();
  const sVals = getSheet_(CFG.SUBMISSIONS_SHEET).getDataRange().getValues();
  const ch = headerMap_(cVals[0] || []);
  const sh = headerMap_(sVals[0] || []);

  const all = {}, done = {};
  const uAll = new Set(), uDone = new Set();

  for (let i = 1; i < cVals.length; i++) {
    const plant = safe_(cVals[i][ch.plant] || cVals[i][ch.workplace]);
    const key = `${safe_(cVals[i][ch.name])}|${safe_(cVals[i][ch.surname])}|${safe_(cVals[i][ch.pesel])}|${plant}`.toLowerCase();
    if (!plant || key === '|||') continue;
    if (!uAll.has(key)) { uAll.add(key); all[plant] = (all[plant] || 0) + 1; }
  }

  for (let i = 1; i < sVals.length; i++) {
    const plant = safe_(sVals[i][sh.plant]);
    const key = `${safe_(sVals[i][sh.name])}|${safe_(sVals[i][sh.surname])}|${safe_(sVals[i][sh.pesel])}|${plant}`.toLowerCase();
    if (!plant || key === '|||') continue;
    if (!uDone.has(key)) { uDone.add(key); done[plant] = (done[plant] || 0) + 1; }
  }

  return Object.keys(all).sort((a,b)=>a.localeCompare(b,'pl')).map(plant => ({
    plant, completed: done[plant] || 0, total: all[plant] || 0
  }));
}


function adminGenerateTestDatabase(adminToken) {
  assertAdmin_(adminToken);

  const loginSeed = getSheet_(CFG.LOGIN_SEED_SHEET);
  const contacts = getSheet_(CFG.CONTACTS_SHEET);
  const clothes = getSheet_(CFG.CLOTHES_SHEET);
  const submissions = getSheet_(CFG.SUBMISSIONS_SHEET);

  const loginSeedHeader = ['name','surname','pesel','plant'];
  const contactsHeader = ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry'];
  const clothesHeader = ['name','surname','plant','shirt','hoodie','pants','jacket','shoes'];
  const submissionsHeader = ['name','surname','pesel','plant','submittedAt'];

  const employees = [
    {
      name:'Jan', surname:'Kowalski', pesel:'90010112345', plant:'Krakow',
      email:'jan.kowalski@example.com', phone:'+48 600700800', workplace:'Krakow', apartment:'ul. Testowa 1/2',
      hireDate:'2024-01-15', clothesSize:'L', shoesSize:'43', notes:'Test 1', bankAccount:'11 2222 3333 4444 5555 6666 7777', birthDate:'1990-01-01', passportNumber:'PA1234567', passportExpiry:'2030-12-31', arrivalDate:'2021-04-10', firstWorkDate:'2021-04-20', intlDrivingLicense:'tak', intlDrivingLicenseExpiry:'2028-05-30',
      shirt:'L', hoodie:'L', pants:'M', jacket:'L', shoes:'43'
    },
    {
      name:'Anna', surname:'Nowak', pesel:'92020254321', plant:'Warszawa',
      email:'anna.nowak@example.com', phone:'+48 601602603', workplace:'Warszawa', apartment:'ul. Próbna 5/8',
      hireDate:'2023-11-10', clothesSize:'M', shoesSize:'39', notes:'Test 2', bankAccount:'22 3333 4444 5555 6666 7777 8888', birthDate:'1992-02-02', passportNumber:'PB7654321', passportExpiry:'2029-09-15', arrivalDate:'2022-06-01', firstWorkDate:'2022-06-15', intlDrivingLicense:'nie', intlDrivingLicenseExpiry:'',
      shirt:'M', hoodie:'M', pants:'S', jacket:'M', shoes:'39'
    },
    {
      name:'Carlos', surname:'Gomez', pesel:'85030311111', plant:'Wroclaw',
      email:'carlos.gomez@example.com', phone:'+57 3201234567', workplace:'Wroclaw', apartment:'Calle 10 #5-20',
      hireDate:'2022-09-01', clothesSize:'XL', shoesSize:'44', notes:'Test 3', bankAccount:'33 4444 5555 6666 7777 8888 9999', birthDate:'1985-03-03', passportNumber:'PC1112223', passportExpiry:'2028-08-08', arrivalDate:'2020-01-12', firstWorkDate:'2020-02-01', intlDrivingLicense:'tak', intlDrivingLicenseExpiry:'2027-01-10',
      shirt:'XL', hoodie:'XL', pants:'L', jacket:'XL', shoes:'44'
    }
  ];

  writeSheetData_(loginSeed, loginSeedHeader, employees.map(e => [e.name, e.surname, e.pesel, e.plant]));
  writeSheetData_(contacts, contactsHeader, employees.map(e => [
    e.name, e.surname, e.email, e.pesel, e.phone, e.workplace, e.apartment, e.plant, e.hireDate, e.clothesSize, e.shoesSize, e.notes,
    e.bankAccount, e.birthDate, e.passportNumber, e.passportExpiry, e.arrivalDate, e.firstWorkDate, e.intlDrivingLicense, e.intlDrivingLicenseExpiry
  ]));
  writeSheetData_(clothes, clothesHeader, employees.map(e => [e.name, e.surname, e.plant, e.shirt, e.hoodie, e.pants, e.jacket, e.shoes]));
  writeSheetData_(submissions, submissionsHeader, employees.map(e => [e.name, e.surname, e.pesel, e.plant, new Date().toISOString()]));

  return { ok:true, generated: employees.length };
}

/* ======================== ADMIN IMPORT EXCEL ======================== */

// 1) Konwersja XLSX -> Google Sheet + lista arkuszy
function adminListExcelSheets(adminToken, filePayload) {
  assertAdmin_(adminToken);
  if (!filePayload || !filePayload.fileBase64) throw new Error('Brak pliku Excel do importu.');

  const convertedSpreadsheetId = convertExcelBase64ToGoogleSheet_(filePayload);
  const ss = SpreadsheetApp.openById(convertedSpreadsheetId);
  return {
    convertedSpreadsheetId,
    sheetNames: ss.getSheets().map(s => s.getName())
  };
}

// 2) Podgląd
function adminPreviewExcel(adminToken, convertedSpreadsheetId, sheetName) {
  assertAdmin_(adminToken);
  const ss = SpreadsheetApp.openById(convertedSpreadsheetId);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Brak arkusza.');

  const maxCols = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (maxCols < 1 || lastRow < 1) return { header: [], rows: [] };

  const header = sh.getRange(1,1,1,maxCols).getValues()[0].map(v => safe_(v));
  if (lastRow < 2) return { header, rows: [] };

  const values = sh.getRange(2,1,lastRow - 1,maxCols).getValues();
  const rows = values.map((row, idx) => ({ rowNumber: idx + 2, rowValues: row }));

  return { header, rows };
}

// 3) Import z mapowaniem + wyborem wierszy
function adminImportWithFieldSelection(adminToken, params) {
  assertAdmin_(adminToken);

  const ss = SpreadsheetApp.openById(params.convertedSpreadsheetId);
  const sh = ss.getSheetByName(params.sheetName);
  if (!sh) throw new Error('Brak arkusza źródłowego.');

  const targetTable = safe_(params.targetTable);
  const mapping = params.mapping || {}; // { "NaglowekExcel": "name" lub "__SKIP__" }
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());

  const selectedRowsSet = new Set((params.selectedRows || []).map(n => Number(n)).filter(n => Number.isFinite(n) && n >= 2));
  const selectedColsSet = new Set((params.selectedColumns || []).map(safe_).filter(Boolean));

  const rowsWithIndex = selectedRowsSet.size
    ? [...selectedRowsSet].sort((a,b)=>a-b).map(r => ({ rowNumber:r, rowValues: sh.getRange(r,1,1,sh.getLastColumn()).getValues()[0] }))
    : loadRowsBySelector_(sh, params.rowSelector, params.startRow, params.endRow);

  let imported = 0;
  const target = getSheet_(targetTable);

  rowsWithIndex.forEach(({rowValues}) => {
    const rec = {};
    header.forEach((sourceHeader, i) => {
      if (selectedColsSet.size && !selectedColsSet.has(sourceHeader)) return;
      const targetField = mapping[sourceHeader];
      if (!targetField || targetField === '__SKIP__') return;
      rec[targetField] = safe_(rowValues[i]);
    });

    if (targetTable === CFG.LOGIN_SEED_SHEET) {
      if (!rec.name || !rec.surname || !rec.pesel || !rec.plant) return;
      upsertByKey_(target, ['name','surname','pesel','plant'], ['name','surname','pesel','plant'], rec);
      imported++;
      return;
    }

    if (targetTable === CFG.CONTACTS_SHEET) {
      if (!rec.name || !rec.surname || !rec.pesel || !rec.plant) return;
      upsertByKey_(target,
        ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry'],
        ['name','surname','pesel','plant'],
        {
          name:rec.name || '', surname:rec.surname || '', pesel:rec.pesel || '', plant:rec.plant || '',
          email:rec.email || '', phone:rec.phone || '', workplace:rec.workplace || rec.plant || '',
          apartment:rec.apartment || '', hireDate:rec.hireDate || '', clothesSize:'', shoesSize:rec.shoes || '', notes:rec.notes || '',
          bankAccount:rec.bankAccount || '', birthDate:rec.birthDate || '', passportNumber:rec.passportNumber || '',
          passportExpiry:rec.passportExpiry || '', arrivalDate:rec.arrivalDate || '', firstWorkDate:rec.firstWorkDate || '',
          intlDrivingLicense:rec.intlDrivingLicense || '', intlDrivingLicenseExpiry:rec.intlDrivingLicenseExpiry || ''
        }
      );
      imported++;
      return;
    }

    if (targetTable === CFG.CLOTHES_SHEET) {
      if (!rec.name || !rec.surname || !rec.plant) return;
      upsertByKey_(target,
        ['name','surname','plant','shirt','hoodie','pants','jacket','shoes'],
        ['name','surname','plant'],
        {
          name:rec.name || '', surname:rec.surname || '', plant:rec.plant || '',
          shirt:rec.shirt || '', hoodie:rec.hoodie || '', pants:rec.pants || '', jacket:rec.jacket || '', shoes:rec.shoes || ''
        }
      );
      imported++;
      return;
    }
  });

  if (targetTable === CFG.LOGIN_SEED_SHEET) {
    syncSeedToCoreTables_();
  }

  return { ok:true, imported };
}

// Auto podpowiedź mapowania dla nagłówków
function adminSuggestMapping(adminToken, convertedSpreadsheetId, sheetName, targetTable) {
  assertAdmin_(adminToken);
  const ss = SpreadsheetApp.openById(convertedSpreadsheetId);
  const sh = ss.getSheetByName(sheetName);
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());

  const sourceIdx = buildAutoMapping_(header);
  const reverse = {}; // canonical -> sourceHeader
  Object.keys(sourceIdx).forEach(c => reverse[c] = header[sourceIdx[c]]);

  const allowed = getAllowedFieldsForTarget_(targetTable);
  const out = {};
  header.forEach(h => out[h] = '__SKIP__');

  allowed.forEach(field => {
    const sourceHeader = reverse[field];
    if (sourceHeader) out[sourceHeader] = field;
  });

  return { header, suggestedMapping: out, allowedFields: [...allowed, '__SKIP__'] };
}

/* ======================== IMPORT HELPERS ======================== */

function loadRowsBySelector_(sheet, rowSelector, startRow, endRow) {
  const maxCols = sheet.getLastColumn();

  // tryb 1: selektor np. "2,5,7-10"
  const selectedRows = parseRowSelector_(rowSelector);
  if (selectedRows.length) {
    return selectedRows
      .filter(r => r >= 2 && r <= sheet.getLastRow())
      .map(r => ({ rowNumber:r, rowValues: sheet.getRange(r,1,1,maxCols).getValues()[0] }));
  }

  // tryb 2: zakres start-end
  const sr = Math.max(2, Number(startRow || 2));
  const er = Math.max(sr, Number(endRow || sr));
  const values = sheet.getRange(sr,1,er-sr+1,maxCols).getValues();
  return values.map((row, idx) => ({ rowNumber: sr + idx, rowValues: row }));
}

function parseRowSelector_(txt) {
  const raw = safe_(txt);
  if (!raw) return [];
  const parts = raw.split(',').map(s => s.trim()).filter(Boolean);

  const set = new Set();
  parts.forEach(p => {
    if (/^\d+$/.test(p)) {
      set.add(Number(p));
      return;
    }
    if (/^\d+\-\d+$/.test(p)) {
      const [a,b] = p.split('-').map(Number);
      const from = Math.min(a,b), to = Math.max(a,b);
      for (let i = from; i <= to; i++) set.add(i);
    }
  });
  return [...set].sort((a,b)=>a-b);
}

function buildAutoMapping_(headerRow) {
  const normHeader = headerRow.map(h => normalizeKey_(h));
  const idx = {};
  Object.keys(HEADER_SYNONYMS).forEach(canonical => {
    const syn = HEADER_SYNONYMS[canonical].map(normalizeKey_);
    const found = normHeader.findIndex(h => syn.includes(h));
    if (found >= 0) idx[canonical] = found;
  });
  return idx;
}

function getAllowedFieldsForTarget_(targetTable) {
  if (targetTable === CFG.LOGIN_SEED_SHEET) return ['name','surname','pesel','plant'];
  if (targetTable === CFG.CONTACTS_SHEET) return ['name','surname','pesel','plant','email','phone','workplace','apartment','hireDate','notes','shoes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry'];
  if (targetTable === CFG.CLOTHES_SHEET) return ['name','surname','plant','shirt','hoodie','pants','jacket','shoes'];
  return [];
}

function normalizeKey_(v) {
  return safe_(v).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function convertExcelBase64ToGoogleSheet_(filePayload) {
  const fileName = safe_(filePayload.fileName) || `import_${Date.now()}.xlsx`;
  const mimeType = safe_(filePayload.mimeType) || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const blob = Utilities.newBlob(Utilities.base64Decode(filePayload.fileBase64), mimeType, fileName);

  const created = Drive.Files.create(
    { name:`IMPORT_${Date.now()}_${fileName}`, mimeType: MimeType.GOOGLE_SHEETS },
    blob
  );
  return created.id;
}

/* ======================== DRIVE + COMMON ======================== */

function sanitizeFilePart_(v){
  return safe_(v).replace(/[^\w\-\sąćęłńóśźżĄĆĘŁŃÓŚŹŻ.]/g,'_').trim().replace(/\s+/g,'_');
}
function getRootFolder_(){
  const it = DriveApp.getFoldersByName(CFG.ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CFG.ROOT_FOLDER_NAME);
}
function getOrCreateSubFolder_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function getEmployeeFolder_(plant, name, surname, pesel){
  const root = getRootFolder_();
  const plantFolder = getOrCreateSubFolder_(root, sanitizeFilePart_(plant));
  return getOrCreateSubFolder_(plantFolder, `${sanitizeFilePart_(name)}_${sanitizeFilePart_(surname)}_${sanitizeFilePart_(pesel)}`);
}
function upsertTextFile_(folder, fileName, content){
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) files.next().setContent(content);
  else folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
}

function getOrCreateSheet_(ss, name){ return ss.getSheetByName(name) || ss.insertSheet(name); }
function writeSheetData_(sheet, header, rows){
  sheet.clearContents();
  sheet.getRange(1,1,1,header.length).setValues([header]);
  if (rows && rows.length) {
    sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  }
}
function ensureHeader_(sheet, header){
  if (sheet.getLastRow() === 0) { sheet.appendRow(header); return; }
  const cur = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(), header.length)).getValues()[0];
  const same = header.every((h,i)=>String(cur[i]||'').trim()===h);
  if (!same) { sheet.clear(); sheet.appendRow(header); }
}
function getSheet_(name){
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error(`Brak zakładki: ${name}`);
  return sh;
}
function safe_(v){ return v == null ? '' : String(v).trim(); }
function headerMap_(headers){ const m={}; headers.forEach((h,i)=>m[String(h).trim()]=i); return m; }
function parsePhone_(raw){
  const t = safe_(raw);
  if (t.startsWith('+57')) return {prefix:'+57', number:t.replace('+57','').trim()};
  if (t.startsWith('+48')) return {prefix:'+48', number:t.replace('+48','').trim()};
  return {prefix:'+48', number:t};
}
function getClothesData_(name, surname, plant){
  const vals = getSheet_(CFG.CLOTHES_SHEET).getDataRange().getValues();
  if (vals.length < 2) return {shirt:'',hoodie:'',pants:'',jacket:'',shoes:''};
  const h = headerMap_(vals[0]);
  for (let i=1;i<vals.length;i++){
    if (safe_(vals[i][h.name])===name && safe_(vals[i][h.surname])===surname && safe_(vals[i][h.plant])===plant){
      return {
        shirt:safe_(vals[i][h.shirt]), hoodie:safe_(vals[i][h.hoodie]),
        pants:safe_(vals[i][h.pants]), jacket:safe_(vals[i][h.jacket]), shoes:safe_(vals[i][h.shoes])
      };
    }
  }
  return {shirt:'',hoodie:'',pants:'',jacket:'',shoes:''};
}
function upsertByKey_(sheet, headers, keyFields, obj){
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  const vals = sheet.getDataRange().getValues();
  const h = headerMap_(vals[0]);
  const target = keyFields.map(k=>safe_(obj[k]).toLowerCase()).join('|');
  let row=-1;
  for (let i=1;i<vals.length;i++){
    const key = keyFields.map(k=>safe_(vals[i][h[k]]).toLowerCase()).join('|');
    if (key===target){ row=i+1; break; }
  }
  const out = headers.map(c=>obj[c] ?? '');
  if (row===-1) sheet.appendRow(out);
  else sheet.getRange(row,1,1,out.length).setValues([out]);
}
function markSubmission_(name,surname,pesel,plant){
  const sh = getSheet_(CFG.SUBMISSIONS_SHEET);
  upsertByKey_(sh, ['name','surname','pesel','plant','submittedAt'], ['name','surname','pesel','plant'],
    {name,surname,pesel,plant,submittedAt:new Date().toISOString()});
}
