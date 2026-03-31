const CFG = {
  LOGIN_SEED_SHEET: 'login_seed',
  CONTACTS_SHEET: 'kontakty',
  CLOTHES_SHEET: 'ubrania_robocze',
  SUBMISSIONS_SHEET: 'form_submissions',
  REGISTRY_SHEET: 'employee_registry',
  PESEL_LIST_SHEET: 'pesel_list',
  PLANT_LIST_SHEET: 'plant_list',
  APARTMENT_LIST_SHEET: 'apartment_list',
  ROOT_FOLDER_NAME: 'Flota_Formularze_Pracownikow',
  ADMIN_LOGIN: 'admin', // ZMIEŃ
  ADMIN_PASSWORD: 'admin123', // ZMIEŃ
  SESSION_TTL_SEC: 20 * 60,
  MAX_UPLOAD_PDF_MB: 8,
  TEST_LOGIN_BYPASS_PESELS: ['99999999999'] // PESEL testowy - brak blokady po zapisie
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
  tpl.canonicalUrl = page === 'admin' ? `${baseUrl}?page=admin` : baseUrl;

  return tpl.evaluate()
    .setTitle(page === 'admin' ? 'Panel administratora' : 'Formulario trabajador')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
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
  const registry = getOrCreateSheet_(ss, CFG.REGISTRY_SHEET);
  const peselList = getOrCreateSheet_(ss, CFG.PESEL_LIST_SHEET);
  const plantList = getOrCreateSheet_(ss, CFG.PLANT_LIST_SHEET);
  const apartmentList = getOrCreateSheet_(ss, CFG.APARTMENT_LIST_SHEET);

  ensureHeader_(loginSeed, ['name','surname','pesel','plant']);
  ensureHeader_(contacts, ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry']);
  ensureHeader_(clothes, ['name','surname','plant','shirt','hoodie','pants','jacket','shoes']);
  ensureHeader_(submissions, ['name','surname','pesel','plant','submittedAt']);
  ensureHeader_(registry, ['pesel','name','surname','plant','apartment']);
  ensureHeader_(peselList, ['pesel','plant']);
  if (peselList.getMaxRows() > 0) peselList.getRange(1,1,peselList.getMaxRows(),1).setNumberFormat('@');
  ensureHeader_(plantList, ['plant']);
  ensureHeader_(apartmentList, ['apartment']);

  syncRegistryToLoginSeed_();
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


function syncRegistryToLoginSeed_() {
  const regVals = getSheet_(CFG.REGISTRY_SHEET).getDataRange().getValues();
  if (regVals.length < 2) return;
  const rh = headerMap_(regVals[0]);

  const seed = getSheet_(CFG.LOGIN_SEED_SHEET);
  writeSheetData_(seed, ['name','surname','pesel','plant'], []);

  const out = [];
  const uniq = new Set();
  for (let i = 1; i < regVals.length; i++) {
    const pesel = safe_(regVals[i][rh.pesel]);
    const name = safe_(regVals[i][rh.name]);
    const surname = safe_(regVals[i][rh.surname]);
    const plant = safe_(regVals[i][rh.plant]);
    const key = `${pesel}|${plant}`.toLowerCase();
    if (!pesel || !name || !surname || !plant || uniq.has(key)) continue;
    uniq.add(key);
    out.push([name, surname, pesel, plant]);
  }

  if (out.length) {
    seed.getRange(2,1,out.length,4).setValues(out);
  }
}

/* ======================== LOGIN + FORM ======================== */

function loginByIdentity(identity) {
  const pesel = normalizePesel_(identity.pesel);
  if (!/^\d{11}$/.test(pesel)) throw new Error('PESEL musi mieć 11 cyfr.');

  // Blokada ponownego logowania po wcześniejszym zapisie formularza
  const subVals = getSheet_(CFG.SUBMISSIONS_SHEET).getDataRange().getValues();
  const subHeader = headerMap_(subVals[0] || []);
  const bypassLoginBlock = (CFG.TEST_LOGIN_BYPASS_PESELS || []).includes(pesel);
  const alreadySubmitted = subVals.slice(1).some(r => normalizePesel_(r[subHeader.pesel]) === pesel);
  if (alreadySubmitted && !bypassLoginBlock) {
    throw new Error('Formularz dla tego PESEL został już zapisany. Ponowne logowanie jest zablokowane.');
  }

  const peselVals = getSheet_(CFG.PESEL_LIST_SHEET).getDataRange().getValues();
  const ph = headerMap_(peselVals[0] || []);
  const peselRows = peselVals.slice(1).filter(r => safe_(r[ph.pesel]) === pesel);
  if (!peselRows.length && !bypassLoginBlock) throw new Error('PESEL nie jest na liście logowania.');

  let plantOptions = [...new Set(peselRows.map(r => safe_(r[ph.plant])).filter(Boolean))];
  if (!plantOptions.length && bypassLoginBlock) {
    const plantVals = getSheet_(CFG.PLANT_LIST_SHEET).getDataRange().getValues();
    const plh = headerMap_(plantVals[0] || []);
    plantOptions = [...new Set(plantVals.slice(1).map(r => safe_(r[plh.plant])).filter(Boolean))];
  }
  if (!plantOptions.length) plantOptions = ['Krakow'];

  const apartmentVals = getSheet_(CFG.APARTMENT_LIST_SHEET).getDataRange().getValues();
  const ah = headerMap_(apartmentVals[0] || []);
  const apartmentOptions = [...new Set(apartmentVals.slice(1).map(r => safe_(r[ah.apartment])).filter(Boolean))];

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    `sess:${token}`,
    JSON.stringify({ name:'', surname:'', pesel, plant: plantOptions[0] || '' }),
    CFG.SESSION_TTL_SEC
  );

  return {
    ok:true,
    token,
    employee:{
      name:'', surname:'', pesel,
      email:'', phone:'', apartment:'', hireDate:'', notes:'',
      bankAccount:'', birthDate:'', passportNumber:'', passportExpiry:'',
      arrivalDate:'', firstWorkDate:'', intlDrivingLicense:'nie', intlDrivingLicenseExpiry:'',
      shirt:'m', hoodie:'m', pants:'50', jacket:'m', shoes:'40',
      phonePrefix:'+48', phoneNumber:'',
      plantOptions, apartmentOptions
    }
  };
}
function saveEmployeeForm(payload) {
  if (!payload || !payload.token) throw new Error('Sesja nieważna.');
  const sessRaw = CacheService.getScriptCache().get(`sess:${payload.token}`);
  if (!sessRaw) throw new Error('Sesja wygasła.');
  const sess = JSON.parse(sessRaw);

  const name = safe_(payload.name) || sess.name;
  const surname = safe_(payload.surname) || sess.surname;
  const pesel = sess.pesel;
  const plant = safe_(payload.plant) || sess.plant;
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


function adminSaveCoreLists(adminToken, payload) {
  assertAdmin_(adminToken);

  const plants = parseMultilineList_(payload && payload.plants);
  const apartments = parseMultilineList_(payload && payload.apartments);
  const selectedPlant = safe_(payload && payload.selectedPlant);
  const pesels = parseMultilineList_(payload && payload.pesels)
    .map(normalizePesel_)
    .filter(v => /^\d{11}$/.test(v));

  // Listy zakładów i mieszkań działają jako słowniki (nadpisanie aktualnym zbiorem wejściowym)
  if (plants.length) writeUniqueColumnSheet_(getSheet_(CFG.PLANT_LIST_SHEET), 'plant', plants);
  if (apartments.length) writeUniqueColumnSheet_(getSheet_(CFG.APARTMENT_LIST_SHEET), 'apartment', apartments);

  // PESEL dopisywany do wybranego zakładu (wielokrotnego użytku, bez duplikatów pary pesel+plant)
  let savedPesels = 0;
  if (pesels.length) {
    if (!selectedPlant) throw new Error('Najpierw wybierz zakład dla listy PESEL.');
    savedPesels = upsertPeselsForPlant_(pesels, selectedPlant);
  }

  return { ok:true, pesels:savedPesels, plants:plants.length, apartments:apartments.length };
}

function adminRemoveFromLists(adminToken, payload) {
  assertAdmin_(adminToken);

  const peselKey = safe_(payload && payload.peselKey);
  const plant = safe_(payload && payload.plant);
  const apartment = safe_(payload && payload.apartment);

  if (peselKey) {
    const [pesel, p] = peselKey.split('|');
    removePeselAssignment_(getSheet_(CFG.PESEL_LIST_SHEET), safe_(pesel), safe_(p));
  }
  if (plant) removeFromSingleColumnSheet_(getSheet_(CFG.PLANT_LIST_SHEET), 'plant', new Set([plant]));
  if (apartment) removeFromSingleColumnSheet_(getSheet_(CFG.APARTMENT_LIST_SHEET), 'apartment', new Set([apartment]));

  return { ok:true };
}



function adminGetCoreLists(adminToken) {
  assertAdmin_(adminToken);

  const plantVals = getSheet_(CFG.PLANT_LIST_SHEET).getDataRange().getValues();
  const plh = headerMap_(plantVals[0] || []);
  const plants = [...new Set(plantVals.slice(1).map(r => safe_(r[plh.plant])).filter(Boolean))];

  const apartmentVals = getSheet_(CFG.APARTMENT_LIST_SHEET).getDataRange().getValues();
  const ah = headerMap_(apartmentVals[0] || []);
  const apartments = [...new Set(apartmentVals.slice(1).map(r => safe_(r[ah.apartment])).filter(Boolean))];

  const cVals = getSheet_(CFG.CONTACTS_SHEET).getDataRange().getValues();
  const ch = headerMap_(cVals[0] || []);
  const nameByKey = {};
  for (let i = 1; i < cVals.length; i++) {
    const pesel = safe_(cVals[i][ch.pesel]);
    const plant = safe_(cVals[i][ch.plant] || cVals[i][ch.workplace]);
    if (!pesel || !plant) continue;
    nameByKey[`${pesel}|${plant}`.toLowerCase()] = `${safe_(cVals[i][ch.surname])} ${safe_(cVals[i][ch.name])}`.trim();
  }

  const pVals = getSheet_(CFG.PESEL_LIST_SHEET).getDataRange().getValues();
  const ph = headerMap_(pVals[0] || []);
  const peselOptions = pVals.slice(1)
    .map(r => {
      const pesel = safe_(r[ph.pesel]);
      const plant = safe_(r[ph.plant]);
      if (!pesel || !plant) return null;
      const key = `${pesel}|${plant}`;
      const fullName = nameByKey[key.toLowerCase()] || 'Brak danych';
      return { value:key, label:`${fullName} | ${pesel} | ${plant}` };
    })
    .filter(Boolean);

  return { plants, apartments, peselOptions };
}


function adminListCompletedEmployees(adminToken) {
  assertAdmin_(adminToken);

  const vals = getSheet_(CFG.SUBMISSIONS_SHEET).getDataRange().getValues();
  const h = headerMap_(vals[0] || []);
  const uniq = new Map();

  for (let i = 1; i < vals.length; i++) {
    const pesel = safe_(vals[i][h.pesel]);
    const plant = safe_(vals[i][h.plant]);
    const name = safe_(vals[i][h.name]);
    const surname = safe_(vals[i][h.surname]);
    if (!pesel || !plant) continue;
    uniq.set(`${pesel}|${plant}`.toLowerCase(), { pesel, plant, name, surname });
  }

  return [...uniq.values()].sort((a,b)=>`${a.surname} ${a.name}`.localeCompare(`${b.surname} ${b.name}`, 'pl'));
}

function adminGetEmployeeForEdit(adminToken, pesel, plant) {
  assertAdmin_(adminToken);

  const p = safe_(pesel), pl = safe_(plant);
  const cVals = getSheet_(CFG.CONTACTS_SHEET).getDataRange().getValues();
  const ch = headerMap_(cVals[0] || []);

  let found = null;
  for (let i = 1; i < cVals.length; i++) {
    const cp = safe_(cVals[i][ch.pesel]);
    const cpl = safe_(cVals[i][ch.plant] || cVals[i][ch.workplace]);
    if (cp === p && cpl.toLowerCase() === pl.toLowerCase()) {
      found = {
        name: safe_(cVals[i][ch.name]), surname: safe_(cVals[i][ch.surname]), pesel: cp, plant: cpl,
        email: safe_(cVals[i][ch.email]), phone: safe_(cVals[i][ch.phone]), apartment: safe_(cVals[i][ch.apartment]),
        hireDate: safe_(cVals[i][ch.hireDate]), notes: safe_(cVals[i][ch.notes]),
        bankAccount: safe_(cVals[i][ch.bankAccount]), birthDate: safe_(cVals[i][ch.birthDate]),
        passportNumber: safe_(cVals[i][ch.passportNumber]), passportExpiry: safe_(cVals[i][ch.passportExpiry]),
        arrivalDate: safe_(cVals[i][ch.arrivalDate]), firstWorkDate: safe_(cVals[i][ch.firstWorkDate]),
        intlDrivingLicense: safe_(cVals[i][ch.intlDrivingLicense]), intlDrivingLicenseExpiry: safe_(cVals[i][ch.intlDrivingLicenseExpiry])
      };
      break;
    }
  }

  if (!found) throw new Error('Nie znaleziono pracownika do edycji.');
  return { ...found, ...getClothesData_(found.name, found.surname, found.plant) };
}

function adminSaveEmployeeByAdmin(adminToken, payload) {
  assertAdmin_(adminToken);

  const name = safe_(payload.name), surname = safe_(payload.surname), pesel = safe_(payload.pesel), plant = safe_(payload.plant);
  if (!name || !surname || !pesel || !plant) throw new Error('Wymagane: imię, nazwisko, PESEL, zakład.');

  const oldPesel = safe_(payload.oldPesel) || pesel;
  const oldPlant = safe_(payload.oldPlant) || plant;
  const oldName = safe_(payload.oldName) || name;
  const oldSurname = safe_(payload.oldSurname) || surname;

  const keyChanged = oldPesel !== pesel || oldPlant.toLowerCase() !== plant.toLowerCase() || oldName !== name || oldSurname !== surname;
  if (keyChanged) {
    deleteByKey_(getSheet_(CFG.CONTACTS_SHEET), ['name','surname','pesel','plant'], { name:oldName, surname:oldSurname, pesel:oldPesel, plant:oldPlant });
    deleteByKey_(getSheet_(CFG.CLOTHES_SHEET), ['name','surname','plant'], { name:oldName, surname:oldSurname, plant:oldPlant });
    deleteByKey_(getSheet_(CFG.SUBMISSIONS_SHEET), ['name','surname','pesel','plant'], { name:oldName, surname:oldSurname, pesel:oldPesel, plant:oldPlant });
    moveEmployeeFolder_(oldPlant, oldName, oldSurname, oldPesel, plant, name, surname, pesel);
  }

  upsertByKey_(getSheet_(CFG.CONTACTS_SHEET),
    ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry'],
    ['name','surname','pesel','plant'],
    {
      name, surname, pesel, plant,
      email:safe_(payload.email), phone:safe_(payload.phone), workplace:plant, apartment:safe_(payload.apartment),
      hireDate:safe_(payload.hireDate), clothesSize:'', shoesSize:safe_(payload.shoes), notes:safe_(payload.notes),
      bankAccount:safe_(payload.bankAccount), birthDate:safe_(payload.birthDate), passportNumber:safe_(payload.passportNumber),
      passportExpiry:safe_(payload.passportExpiry), arrivalDate:safe_(payload.arrivalDate), firstWorkDate:safe_(payload.firstWorkDate),
      intlDrivingLicense:safe_(payload.intlDrivingLicense), intlDrivingLicenseExpiry:safe_(payload.intlDrivingLicenseExpiry)
    }
  );

  upsertByKey_(getSheet_(CFG.CLOTHES_SHEET),
    ['name','surname','plant','shirt','hoodie','pants','jacket','shoes'],
    ['name','surname','plant'],
    {
      name, surname, plant,
      shirt:safe_(payload.shirt), hoodie:safe_(payload.hoodie), pants:safe_(payload.pants), jacket:safe_(payload.jacket), shoes:safe_(payload.shoes)
    }
  );

  markSubmission_(name, surname, pesel, plant);
  return { ok:true };
}

/* ======================== ADMIN STATS ======================== */

function getAdminStats(adminToken) {
  assertAdmin_(adminToken);
  const pVals = getSheet_(CFG.PESEL_LIST_SHEET).getDataRange().getValues();
  const sVals = getSheet_(CFG.SUBMISSIONS_SHEET).getDataRange().getValues();
  const ph = headerMap_(pVals[0] || []);
  const sh = headerMap_(sVals[0] || []);

  const allByPlant = {};
  const doneByPlant = {};

  for (let i = 1; i < pVals.length; i++) {
    const plant = safe_(pVals[i][ph.plant]);
    const pesel = normalizePesel_(pVals[i][ph.pesel]);
    if (!plant || !/^\d{11}$/.test(pesel)) continue;
    if (!allByPlant[plant]) allByPlant[plant] = new Set();
    allByPlant[plant].add(pesel);
  }

  for (let i = 1; i < sVals.length; i++) {
    const plant = safe_(sVals[i][sh.plant]);
    const pesel = normalizePesel_(sVals[i][sh.pesel]);
    if (!plant || !/^\d{11}$/.test(pesel)) continue;
    if (!doneByPlant[plant]) doneByPlant[plant] = new Set();
    doneByPlant[plant].add(pesel);
  }

  const plants = [...new Set(Object.keys(allByPlant).concat(Object.keys(doneByPlant)))].sort((a,b)=>a.localeCompare(b,'pl'));
  return plants.map(plant => {
    const total = allByPlant[plant] ? allByPlant[plant].size : 0;
    const completed = doneByPlant[plant]
      ? [...doneByPlant[plant]].filter(p => !allByPlant[plant] || allByPlant[plant].has(p)).length
      : 0;
    const remaining = Math.max(0, total - completed);
    const percent = total ? Math.round((completed / total) * 100) : 0;
    return { plant, completed, total, remaining, percent };
  });
}

function getAdminMissingByPlant(adminToken, plantName) {
  assertAdmin_(adminToken);
  const plant = safe_(plantName);
  if (!plant) return [];

  const pVals = getSheet_(CFG.PESEL_LIST_SHEET).getDataRange().getValues();
  const sVals = getSheet_(CFG.SUBMISSIONS_SHEET).getDataRange().getValues();
  const ph = headerMap_(pVals[0] || []);
  const sh = headerMap_(sVals[0] || []);

  const all = new Set();
  const done = new Set();

  for (let i = 1; i < pVals.length; i++) {
    const p = safe_(pVals[i][ph.plant]);
    if (p !== plant) continue;
    const pesel = normalizePesel_(pVals[i][ph.pesel]);
    if (/^\d{11}$/.test(pesel)) all.add(pesel);
  }

  for (let i = 1; i < sVals.length; i++) {
    const p = safe_(sVals[i][sh.plant]);
    if (p !== plant) continue;
    const pesel = normalizePesel_(sVals[i][sh.pesel]);
    if (/^\d{11}$/.test(pesel)) done.add(pesel);
  }

  return [...all].filter(p => !done.has(p)).sort((a,b)=>a.localeCompare(b,'pl'));
}


function adminGenerateTestDatabase(adminToken) {
  assertAdmin_(adminToken);

  const loginSeed = getSheet_(CFG.LOGIN_SEED_SHEET);
  const contacts = getSheet_(CFG.CONTACTS_SHEET);
  const clothes = getSheet_(CFG.CLOTHES_SHEET);
  const submissions = getSheet_(CFG.SUBMISSIONS_SHEET);
  const registry = getSheet_(CFG.REGISTRY_SHEET);

  const loginSeedHeader = ['name','surname','pesel','plant'];
  const contactsHeader = ['name','surname','email','pesel','phone','workplace','apartment','plant','hireDate','clothesSize','shoesSize','notes','bankAccount','birthDate','passportNumber','passportExpiry','arrivalDate','firstWorkDate','intlDrivingLicense','intlDrivingLicenseExpiry'];
  const clothesHeader = ['name','surname','plant','shirt','hoodie','pants','jacket','shoes'];
  const submissionsHeader = ['name','surname','pesel','plant','submittedAt'];
  const registryHeader = ['pesel','name','surname','plant','apartment'];

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
  writeSheetData_(registry, registryHeader, employees.map(e => [e.pesel, e.name, e.surname, e.plant, e.apartment]));

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

    if (targetTable === CFG.REGISTRY_SHEET) {
      if (!rec.pesel || !rec.name || !rec.surname || !rec.plant) return;
      upsertByKey_(target,
        ['pesel','name','surname','plant','apartment'],
        ['pesel','plant'],
        { pesel:rec.pesel || '', name:rec.name || '', surname:rec.surname || '', plant:rec.plant || '', apartment:rec.apartment || '' }
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

  if (targetTable === CFG.REGISTRY_SHEET) {
    syncRegistryToLoginSeed_();
    syncSeedToCoreTables_();
  }

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
  if (targetTable === CFG.REGISTRY_SHEET) return ['pesel','name','surname','plant','apartment'];
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
function normalizePesel_(v){
  const digits = safe_(v).replace(/\D/g, '');
  if (!digits) return '';
  return digits.slice(-11).padStart(11, '0');
}
function parseMultilineList_(txt){
  return [...new Set(String(txt || '')
    .split(/\r?\n/)
    .map(s => safe_(s))
    .filter(Boolean))];
}
function writeUniqueColumnSheet_(sheet, headerName, values){
  writeSheetData_(sheet, [headerName], values.map(v => [v]));
}
function removeFromSingleColumnSheet_(sheet, headerName, removeSet){
  const vals = sheet.getDataRange().getValues();
  const h = headerMap_(vals[0] || []);
  const idx = h[headerName];
  if (idx == null) return;
  const kept = vals.slice(1).map(r => safe_(r[idx])).filter(v => v && !removeSet.has(v));
  writeUniqueColumnSheet_(sheet, headerName, kept);
}
function upsertPeselsForPlant_(pesels, plant){
  const sheet = getSheet_(CFG.PESEL_LIST_SHEET);
  const vals = sheet.getDataRange().getValues();
  const h = headerMap_(vals[0] || []);

  const existing = new Set(vals.slice(1).map(r => `${safe_(r[h.pesel])}|${safe_(r[h.plant])}`.toLowerCase()));
  const out = vals.slice(1).map(r => [safe_(r[h.pesel]), safe_(r[h.plant])]);

  let added = 0;
  pesels.forEach(p => {
    const np = normalizePesel_(p);
    if (!/^\d{11}$/.test(np)) return;
    const key = `${np}|${safe_(plant)}`.toLowerCase();
    if (!existing.has(key)) {
      existing.add(key);
      out.push([np, safe_(plant)]);
      added++;
    }
  });

  writeSheetData_(sheet, ['pesel','plant'], out);
  if (sheet.getMaxRows() > 0) sheet.getRange(1,1,sheet.getMaxRows(),1).setNumberFormat('@');
  return added;
}
function removePeselAssignment_(sheet, pesel, plant){
  const vals = sheet.getDataRange().getValues();
  const h = headerMap_(vals[0] || []);
  const kept = vals.slice(1)
    .filter(r => !(safe_(r[h.pesel]) === pesel && safe_(r[h.plant]).toLowerCase() === plant.toLowerCase()))
    .map(r => [safe_(r[h.pesel]), safe_(r[h.plant])]);
  writeSheetData_(sheet, ['pesel','plant'], kept);
  if (sheet.getMaxRows() > 0) sheet.getRange(1,1,sheet.getMaxRows(),1).setNumberFormat('@');
}
function removePeselsFromSheet_(sheet, removeSet){
  const vals = sheet.getDataRange().getValues();
  const h = headerMap_(vals[0] || []);
  const kept = vals.slice(1)
    .filter(r => !removeSet.has(safe_(r[h.pesel])))
    .map(r => [safe_(r[h.pesel]), safe_(r[h.plant])]);
  writeSheetData_(sheet, ['pesel','plant'], kept);
  if (sheet.getMaxRows() > 0) sheet.getRange(1,1,sheet.getMaxRows(),1).setNumberFormat('@');
}
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
function deleteByKey_(sheet, keyFields, obj){
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return;
  const h = headerMap_(vals[0]);
  const target = keyFields.map(k=>safe_(obj[k]).toLowerCase()).join('|');
  for (let i = vals.length - 1; i >= 1; i--) {
    const key = keyFields.map(k=>safe_(vals[i][h[k]]).toLowerCase()).join('|');
    if (key === target) sheet.deleteRow(i + 1);
  }
}
function findEmployeeFolder_(plant, name, surname, pesel){
  const root = getRootFolder_();
  const pit = root.getFoldersByName(sanitizeFilePart_(plant));
  if (!pit.hasNext()) return null;
  const plantFolder = pit.next();
  const folderName = `${sanitizeFilePart_(name)}_${sanitizeFilePart_(surname)}_${sanitizeFilePart_(pesel)}`;
  const fit = plantFolder.getFoldersByName(folderName);
  return fit.hasNext() ? fit.next() : null;
}
function moveEmployeeFolder_(oldPlant, oldName, oldSurname, oldPesel, newPlant, newName, newSurname, newPesel){
  const oldFolder = findEmployeeFolder_(oldPlant, oldName, oldSurname, oldPesel);
  if (!oldFolder) return;
  const newFolder = getEmployeeFolder_(newPlant, newName, newSurname, newPesel);

  const files = oldFolder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    newFolder.addFile(f);
    oldFolder.removeFile(f);
  }

  oldFolder.setTrashed(true);
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
