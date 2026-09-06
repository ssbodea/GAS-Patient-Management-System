const CONFIG = {
  SHEET_NAME: "Form Responses 1",

  TIMESTAMP: 0,
  NAME: 1,
  AGE: 2,
  GENDER: 3,
  CONTRACEPTION: 4,
  LMP: 5,
  ADDRESS: 6,
  OCCUPATION: 7,
  FACULTY: 8,
  YEAR: 9,
  LANGUAGE: 10,
  REASON: 11,
  SYMPTOMS: 12,
  TREATMENT: 13,
  DISEASES: 14,
  ALLERGIES: 15,

  EC_TEMPERATURE: 16,
  EC_BLOOD_PRESSURE: 17,
  EC_PULSE: 18,
  EC_SATURATION: 19,
  EC_DIAGNOSIS: 20,
  EC_DISEASE_CODES: 21,
  EC_RECOMMENDATIONS: 22,
  EC_VACCINATIONS: 23,

  RP_FULL: 24,
  RP_FREE: 25,
  BT_CAS_1_SERIAL: 26,
  BT_CAS_1_SPECIALTY: 27,
  BT_CAS_1_TYPE: 28,
  BT_CAS_2_SERIAL: 29,
  BT_CAS_2_SPECIALTY: 30,
  BT_CAS_2_TYPE: 31,
  BT_CAS_3_SERIAL: 32,
  BT_CAS_3_SPECIALTY: 33,
  BT_CAS_3_TYPE: 34,
  BT_SIMPLE: 35,
  BT_CHRONIC_SPECIALIST: 36,

  AE_ABSENCE_EXCUSE_START: 37,
  AE_ABSENCE_EXCUSE_END: 38,
  AE_SPORT_EXCUSE_START: 39,
  AE_SPORT_EXCUSE_END: 40,
  AE_SPORT_ENDORSE_START: 41,
  AE_SPORT_ENDORSE_END: 42,
  AE_OTHER_PURPOSE: 43,
  AE_MEDICAL_SCHOLARSHIP: 44,
  AE_EPIDEMIOLOGIC_CLEARANCE: 45,
  AE_SPORT_COMPETITION: 46,
  AE_ABSENCE_ENDORSE_1_START: 47,
  AE_ABSENCE_ENDORSE_1_END: 48,
  AE_ABSENCE_ENDORSE_2_START: 49,
  AE_ABSENCE_ENDORSE_2_END: 50,
  AE_ABSENCE_ENDORSE_3_START: 51,
  AE_ABSENCE_ENDORSE_3_END: 52,
  AE_ABSENCE_ENDORSE_4_START: 53,
  AE_ABSENCE_ENDORSE_4_END: 54,
  AE_ABSENCE_ENDORSE_5_START: 55,
  AE_ABSENCE_ENDORSE_5_END: 56,

  EB_HEIGHT: 57,
  EB_HEIGHT_INDEX: 58,
  EB_WEIGHT: 59,
  EB_WEIGHT_INDEX: 60,
  EB_BMI: 61,
  EB_PHYSICAL_DEVELOPMENT: 62,
  EB_DISEASE_CODES_OLD: 63,
  EB_DISEASE_CODES_NEW: 64,

  COLUMN_COUNT: 65,
  LOCK_WAIT: 5000,
  TARGET_DATA_ROWS: 3000,
  TIMEZONE: "Europe/Bucharest",
  LOCALE: "ro_RO",
  FACULTIES: ["MG", "MM", "MD", "PHARMA", "AMG", "BFK", "RI", "ND", "TD", "CM", "MASTER"],
  BODY_INDEX_LEVELS: ["M-3", "M-2", "M±1", "M+2", "M+3"],
  BODY_INDEX_LABELS: {
    "M-3": "F.MICĂ",
    "M-2": "MICĂ",
    "M±1": "MIJLOCIE",
    "M+2": "MARE",
    "M+3": "F.MARE",
  },
  PHYSICAL_DEV_LEVELS: ["Armonică", "Dizarmonică +G", "Dizarmonică -G"],
  CLINIC_CODE_INTERVALS: [
    [1, 79, "Boli infecțioase și parazitare, din care:"],
    [54, 56, "- hepatita virală"],
    [36, 36, "- infecție gonococică"],
    [33, 33, "- sifilis recent"],
    [65, 67, "- dermatomicoze și alte micoze"],
    [13, 13, "- boli diareice"],
    [80, 202, "Tumori"],
    [234, 298, "Boli endocrine, metabolism, nutriție"],
    [203, 233, "Boli ale sângelui și organelor hematopoietice"],
    [299, 355, "Tulburări mintale"],
    [356, 397, "Boli ale sistemului nervos"],
    [398, 427, "Boli ale ochiului"],
    [428, 444, "Boli ale urechii și apofizei mastoide"],
    [445, 497, "Boli ale aparatului circulator"],
    [498, 542, "Boli ale aparatului respirator, din care:"],
    [498, 503, "- boli ale căilor respiratorii superioare"],
    [506, 511, "- pneumonie"],
    [504, 505, "- gripă"],
    [543, 591, "Boli ale aparatului digestiv, din care:"],
    [543, 552, "- boli ale cavității bucale, din care:"],
    [544, 544, "- caria dentară"],
    [592, 625, "Boli ale pielii și țesutului subcutanat"],
    [670, 732, "Boli ale aparatului genito-urinar"],
    [879, 975, "Accidente, traumatisme, otrăviri"],
  ],
  PREVALENCE_CODE_INTERVALS: [
    [2, "Tuberculoză (indiferent de localizare)", [[14, 17]]],
    [3, "Hepatită virală (în ultimele 12 luni)", [[54, 56]]],
    [4, "Tumori maligne", [[80, 160], [166, 176]]],
    [5, "Leucemii", [[161, 165]]],
    [6, "Tumori benigne", [[177, 202]]],
    [7, "Anemii prin carență de fier", [[203, 203]]],
    [8, "Alte anemii cronice", [[204, 207], [209, 211]]],
    [9, "Talasemie", [[208, 208]]],
    [10, "Alte boli ale sângelui și org. hematopoietice", [[212, 233]]],
    [11, "Gușă simplă și alte boli ale tiroidei", [[234, 240]]],
    [12, "Diabet zaharat", [[241, 245]]],
    [13, "Alte boli endocrine și de metabolism", [[246, 261]]],
    [14, "Hipotrofie ponderală", [[264, 264]]],
    [15, "Hipotrofie staturală", [[264, 264]]],
    [16, "Sechele de rahitism", [[272, 272], [274, 274]]],
    [17, "Obezitatea de origine neendocrină", [[279, 279]]],
    [18, "Alte tulburări mintale", [[338, 341], [343, 345], [299, 324]]],
    [19, "Tulburări nevrotice", [[325, 330]]],
    [20, "Instabilitate psihomotorie", [[331, 337], [348, 349]]],
    [21, "Întârziere mintală ușoară", [[342, 342]]],
    [22, "Întârziere mintală de nivel neprecizat", [[346, 346]]],
    [23, "Tulburări de vorbire", [[347, 347]]],
    [24, "Tulburări de comportament și de adaptare școlară", [[350, 354]]],
    [25, "Alte boli ale sistemului nervos", [[356, 372], [375, 397]]],
    [26, "Epilepsie", [[373, 374]]],
    [27, "Tulburări de vedere, altele decât prin vicii de refracție", [[423, 424], [415, 415]]],
    [28, "Vicii de refracție", [[420, 422]]],
    [29, "Alte boli cronice ale ochiului și anexelor sale", [[425, 427]]],
    [30, "Otita medie cronică", [[430, 431]]],
    [31, "Tulburări de auz (hipoacuzia, surditatea)", [[440, 441]]],
    [32, "Alte boli cronice ale urechii și apofizei mastoide", [[442, 444], [438, 439], [432, 436]]],
    [33, "Reumatismul articular acut (în ultimii 5 ani)", [[445, 447]]],
    [34, "Cardiopatii reumatismale cronice", [[448, 452]]],
    [35, "B. hipertensive (incl. oscilațiile tensionale pubertare sau post pubert.)", [[453, 457]]],
    [36, "Alte forme de cardiopatii", [[467, 476]]],
    [37, "Alte boli vasculare periferice", [[485, 485]]],
    [38, "Bolile arterelor și arteriolelor", [[487, 487]]],
    [39, "Bolile capilarelor", [[488, 488]]],
    [40, "Alte boli cronice ale ap. respirator", [[522, 526], [528, 528], [514, 515], [518, 521]]],
    [41, "Sinuzita cronică", [[516, 516]]],
    [42, "Afecțiuni cronice ale amigdalelor și vegetațiilor adenoide", [[517, 517]]],
    [43, "Astmul (bronșic și bronșita asmatiformă)", [[527, 527]]],
    [44, "Ulcerul gastric și duodenal", [[555, 558]]],
    [45, "Boli cronice hepatice (hepatice, ciroze)", [[578, 583]]],
    [46, "Afecțiuni cronice biliare (litiazice și nelitiazice)", [[584, 588]]],
    [47, "Alte boli cronice ale ap. digestiv", [[589, 591]]],
    [48, "Boli ale pielii și țesutului celular subcutanat", [[592, 625]]],
    [49, "Afecțiuni reumatice cronice (fără status post RAA)", [[632, 635], [646, 653]]],
    [50, "Deformații câștigate ale membrelor", [[636, 639]]],
    [51, "Deformații câștigate ale coloanei vertebrale", [[643, 645]]],
    [52, "Alte boli cronice ale sist. osteoart., mușchilor și ale țesut. conjunctiv", [[640, 642]]],
    [53, "Glomerulonefrita (în ultimele 12 luni)", [[678, 683]]],
    [54, "Sindromul nefrotic și nefrozele", [[670, 677]]],
    [55, "Alte boli cronice ale ap. urinar", [[690, 693], [685, 686]]],
    [56, "Calculoza căilor urinare", [[687, 689]]],
    [57, "Afecțiuni ale org. genitale feminine", [[709, 713], [725, 728]]],
    [58, "Anomalii congenitale ale inimii și ale ap. circulator", [[829, 833]]],
    [59, "Anomalii congenitale ale sistemului osteoarticular", [[851, 857]]],
  ],
};

const ANAMNESIS_KEYS = ["REASON", "SYMPTOMS", "TREATMENT", "DISEASES", "ALLERGIES"];
const IDENTITY_KEYS = Object.fromEntries(
  [
    "TIMESTAMP", "NAME", "AGE", "GENDER", "CONTRACEPTION", "LMP", "ADDRESS",
    "OCCUPATION", "FACULTY", "YEAR", "LANGUAGE",
  ]
    .concat(ANAMNESIS_KEYS)
    .map((k) => [k, true]),
);

const CHECKBOX_KEYS = {
  AE_MEDICAL_SCHOLARSHIP: true,
  AE_EPIDEMIOLOGIC_CLEARANCE: true,
  AE_SPORT_COMPETITION: true,
  BT_CHRONIC_SPECIALIST: true,
};

const MSG_NO_PATIENTS = "Nu există pacienți în baza de date";
const SHIFT_END_EMAIL = "SHIFT_END_EMAIL";
const EXPORT_SHEET_PREFIX = "Export ";
const REPORT_CODE_HEADERS = ["Nr. crt.", "Cod", "Specificare", "Total nr. cazuri"];

const EXPORT_AE_PRIMARY = [
  ["Scutire Absență", "ae_absence_excuse_start", "ae_absence_excuse_end"],
  ["Scutire Sport", "ae_sport_excuse_start", "ae_sport_excuse_end"],
  ["Vizare Sport", "ae_sport_endorse_start", "ae_sport_endorse_end"],
];

const EXPORT_AE_ABSENCE_ENDORSE = [1, 2, 3, 4, 5].map((n) => [
  "Vizare Absență" + n,
  "ae_absence_endorse_" + n + "_start",
  "ae_absence_endorse_" + n + "_end",
]);

const EXPORT_HEADERS = [
  "Nr. crt.",
  "Ziua",
  "Numele și prenumele",
  "Vârsta",
  "Sexul",
  "Domiciliul, județ, localitate, str., nr.",
  "Ocupație",
  "Simptome",
  "Diagnostic",
  "Cod",
  "Prescripții medicamente, analize, concediu medical, tratament, etc.",
];

const EXPORT_COLUMN_WIDTHS = [41, 81, 133, 25, 20, 133, 71, 133, 132, 36, 270];

const REPORT_CELL =
  'style="padding:8px; border:1px solid #ddd; text-align:center; vertical-align:middle; height:100%;"';
const REPORT_CELL_GROUP =
  'style="padding:8px; border:1px solid #ddd; text-align:center; font-weight:bold; background:#f9f9f9; vertical-align:middle; height:100%;"';
const REPORT_TH =
  'style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center; vertical-align:middle;"';
const REPORT_TABLE_CSS =
  "border-collapse:collapse; border-spacing:0; margin:0; mso-table-lspace:0pt; mso-table-rspace:0pt; border:1px solid #ddd; table-layout:fixed; width:100%;";

const REPORT_ACT_FLAGS = [
  ["vaccinations", "ec_vaccinations"],
  ["rp_full", "rp_full"],
  ["rp_free", "rp_free"],
  ["bt_simple", "bt_simple"],
  ["bt_chronic_specialist", "bt_chronic_specialist"],
  ["ae_absence_excuse", "ae_absence_excuse_start"],
  ["ae_sport_excuse", "ae_sport_excuse_start"],
  ["ae_sport_endorse", "ae_sport_endorse_start"],
  ["ae_other_purpose", "ae_other_purpose"],
  ["ae_medical_scholarship", "ae_medical_scholarship"],
  ["ae_epidemiologic_clearance", "ae_epidemiologic_clearance"],
  ["ae_sport_competition", "ae_sport_competition"],
];

const COLUMN_KEYS = Object.keys(CONFIG).filter((key) => {
  const value = CONFIG[key];
  return (
    typeof value === "number" &&
    value >= 0 &&
    value < CONFIG.COLUMN_COUNT &&
    Math.floor(value) === value
  );
});

function isDateKey(key) {
  return (
    key === "TIMESTAMP" ||
    key === "LMP" ||
    key.endsWith("_START") ||
    key.endsWith("_END")
  );
}

function columnNumberFormat(key) {
  if (key === "TIMESTAMP") return "dd.MM.yyyy HH:mm:ss";
  if (isDateKey(key)) return "dd.MM.yyyy";
  return "@";
}

function forEachColumn(fn) {
  COLUMN_KEYS.forEach((key) => fn(key, CONFIG[key]));
}

function applyRoboto(range, weight) {
  return range.setFontFamily("Roboto").setFontSize(10).setFontWeight(weight);
}

function reply(type, message, data) {
  const result = {
    success: type !== "error",
    type,
    message: message || (type === "error" ? "Operație eșuată" : ""),
  };
  if (data !== undefined) result.data = data;
  return result;
}

function ok(message, data) {
  return reply("success", message, data);
}

function info(message, data) {
  return reply("info", message, data);
}

function fail(message, data) {
  return reply("error", message, data);
}

function errorMessage(error) {
  if (error == null) return "Eroare necunoscută";
  if (typeof error === "string") return error;
  return String(error.message || error);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    lock.waitLock(CONFIG.LOCK_WAIT);
    acquired = true;
    const result = fn();
    if (result == null || typeof result !== "object" || result.success == null) {
      return fail("Răspuns invalid de la server");
    }
    return result;
  } catch (error) {
    return fail(errorMessage(error));
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function upper(value) {
  return String(value || "").toUpperCase();
}

function normName(value) {
  return String(value == null ? "" : value).trim().toLowerCase();
}

function normAge(value) {
  if (value === "" || value == null) return "";
  const n = Number(value);
  return isFinite(n) ? String(n) : String(value).trim();
}

function genderCode(value) {
  const g = upper(value);
  return g === "M" || g === "F" ? g : "";
}

function pacientWord(n) {
  return n === 1 ? "pacient" : "pacienți";
}

function pacientOk(n, stem, suffix) {
  return n + " " + pacientWord(n) + " " + stem + (n === 1 ? "t" : "ți") + (suffix || "");
}

function sessionEmail() {
  return Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || "";
}

function emptyLangs() {
  return { n: 0, RO: 0, EN: 0, FR: 0 };
}

function addLang(langs, value) {
  langs.n++;
  const lang = upper(value);
  if (lang === "RO" || lang === "EN" || lang === "FR") langs[lang]++;
}

function umfSubject(kind, when, n, langs, startTime) {
  const sec = ((Date.now() - startTime) / 1000).toFixed(2);
  return `UMF · ${kind} · ${when} · ${n} ${pacientWord(n)} · RO×${langs.RO} EN×${langs.EN} FR×${langs.FR} · generat în ${sec}s`;
}

function isDateObject(value) {
  return value && typeof value.getTime === "function" && !isNaN(value.getTime());
}

function asDate(value) {
  if (isDateObject(value)) {
    return value instanceof Date ? value : new Date(value.getTime());
  }
  if (value == null || value === "") return null;
  const text = String(value).trim();
  const ro = text.match(
    /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/,
  );
  if (ro) {
    const d = new Date(
      Number(ro[3]),
      Number(ro[2]) - 1,
      Number(ro[1]),
      Number(ro[4] || 0),
      Number(ro[5] || 0),
      Number(ro[6] || 0),
    );
    return isNaN(d.getTime()) ? null : d;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    try {
      return Utilities.parseDate(text, CONFIG.TIMEZONE, "yyyy-MM-dd");
    } catch (error) {
      return null;
    }
  }
  const d = new Date(text);
  return isNaN(d.getTime()) ? null : d;
}

function toSheetDate(value) {
  return value ? asDate(value) || "" : "";
}

function toTimestampMs(value) {
  const d = asDate(value);
  return d ? d.getTime() : null;
}

function formatSheetDate(date, pattern) {
  return Utilities.formatDate(date, CONFIG.TIMEZONE, pattern);
}

function formatWith(value, pattern) {
  const d = asDate(value);
  return d ? formatSheetDate(d, pattern) : "";
}

function formatDay(value) {
  return formatWith(value, "dd.MM.yyyy");
}

function dayBound(value, endOfDay) {
  return Utilities.parseDate(
    String(value).slice(0, 10) + (endOfDay ? " 23:59:59" : " 00:00:00"),
    CONFIG.TIMEZONE,
    "yyyy-MM-dd HH:mm:ss",
  );
}

function timestampInRange(start, end, values) {
  const ts = asDate(values[CONFIG.TIMESTAMP]);
  return !!(ts && ts >= start && ts <= end);
}

function getSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("Foaia " + CONFIG.SHEET_NAME + " nu a fost găsită.");
  return sheet;
}

function formatDataRows(sheet, fromRow, toRow) {
  if (toRow < fromRow) return;
  const n = toRow - fromRow + 1;
  const rowFmt = new Array(CONFIG.COLUMN_COUNT);
  forEachColumn((key, index) => {
    rowFmt[index] = columnNumberFormat(key);
  });
  const grid = new Array(n);
  for (let i = 0; i < n; i++) grid[i] = rowFmt;
  applyRoboto(
    sheet.getRange(fromRow, 1, n, CONFIG.COLUMN_COUNT).setNumberFormats(grid),
    "normal",
  );
}

function getDataRows() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet
    .getRange(2, 1, lastRow - 1, CONFIG.COLUMN_COUNT)
    .getValues()
    .map((values, i) => ({ rowNum: i + 2, values }));
}

function writeRow(sheet, rowNum, row) {
  sheet.getRange(rowNum, 1, 1, CONFIG.COLUMN_COUNT).setValues([row]);
}

function rowToPatient(values, rowNum) {
  const patient = { id: rowNum - 1 };
  forEachColumn((key, index) => {
    let value = index < values.length ? values[index] : "";
    if (value == null) value = "";
    if (isDateKey(key)) {
      const d = asDate(value);
      if (!d) value = "";
      else if (key === "TIMESTAMP") value = d.toISOString();
      else value = formatSheetDate(d, "yyyy-MM-dd");
    } else if (typeof value === "number" && !isFinite(value)) {
      value = "";
    } else if (typeof value === "object") {
      value = "";
    }
    patient[key.toLowerCase()] = value;
  });
  return patient;
}

function getPatientRow(patientData) {
  const key = patientData && patientData.lookup;
  const wantMs = key && toTimestampMs(key.timestamp);
  if (!key || wantMs == null) {
    return { error: key ? "Timestamp invalid" : "Pacientul nu a fost găsit" };
  }
  const wantName = normName(key.name);
  const wantGender = upper(key.gender);
  const wantAge = normAge(key.age);
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { error: MSG_NO_PATIENTS };
  const n = lastRow - 1;
  const lookup = sheet.getRange(2, 1, n, CONFIG.GENDER + 1).getValues();
  for (let i = n - 1; i >= 0; i--) {
    const values = lookup[i];
    if (
      toTimestampMs(values[CONFIG.TIMESTAMP]) === wantMs &&
      normName(values[CONFIG.NAME]) === wantName &&
      upper(values[CONFIG.GENDER]) === wantGender &&
      normAge(values[CONFIG.AGE]) === wantAge
    ) {
      const rowNum = i + 2;
      return {
        sheet,
        rowNum,
        row: sheet.getRange(rowNum, 1, 1, CONFIG.COLUMN_COUNT).getValues()[0],
      };
    }
  }
  return { error: "Pacientul nu a fost găsit" };
}

function applyIdentityFields(row, patientData) {
  const facultyRaw = String(patientData.faculty || "").trim();
  const facultyKey = upper(facultyRaw);
  row[CONFIG.NAME] = patientData.name || "";
  row[CONFIG.AGE] =
    patientData.age === "" || patientData.age == null ? "" : patientData.age;
  row[CONFIG.GENDER] = upper(patientData.gender);
  row[CONFIG.CONTRACEPTION] = upper(patientData.contraception);
  row[CONFIG.LMP] = toSheetDate(patientData.lmp);
  row[CONFIG.ADDRESS] = patientData.address || "";
  row[CONFIG.OCCUPATION] = upper(patientData.occupation);
  row[CONFIG.FACULTY] = CONFIG.FACULTIES.includes(facultyKey)
    ? facultyKey
    : facultyRaw;
  row[CONFIG.YEAR] = patientData.year || "";
  row[CONFIG.LANGUAGE] = upper(patientData.language);
  ANAMNESIS_KEYS.forEach((k) => {
    row[CONFIG[k]] = patientData[k.toLowerCase()] || "";
  });
}

function applyClinicalFields(row, patientData) {
  forEachColumn((key, index) => {
    if (IDENTITY_KEYS[key]) return;
    const prop = key.toLowerCase();
    if (isDateKey(key)) row[index] = toSheetDate(patientData[prop]);
    else if (CHECKBOX_KEYS[key]) row[index] = patientData[prop] ? "TRUE" : "";
    else row[index] = patientData[prop] || "";
  });
}

function updatePatientRow(patientData, applyFn, message, includePatient) {
  return withLock(() => {
    const found = getPatientRow(patientData);
    if (found.error) return fail(found.error);
    applyFn(found.row, patientData);
    writeRow(found.sheet, found.rowNum, found.row);
    return includePatient
      ? ok(message, rowToPatient(found.row, found.rowNum))
      : ok(message);
  });
}

function formatExportMetric(value) {
  if (value == null || value === "") return "";
  const n = typeof value === "number" ? value : parseFloat(String(value).trim());
  if (!isFinite(n) || Math.abs(n) > 500) return "";
  return String(Math.round(n * 10) / 10);
}

function pushLabeled(lines, label, value) {
  if (value) lines.push(label + " " + value);
}

function pushExportDateLines(lines, patient, pairs) {
  pairs.forEach(([label, startKey, endKey]) => {
    const text = [formatDay(patient[startKey]), formatDay(patient[endKey])]
      .filter(Boolean)
      .join("-");
    if (text) lines.push(label + " " + text);
  });
}

function countIndexed(patient, prefix, suffix, from, to) {
  let n = 0;
  for (let i = from; i <= to; i++) {
    if (patient[prefix + i + suffix]) n++;
  }
  return n;
}

function occupationExport(patient) {
  const occ = upper(patient.occupation);
  if (occ === "STUDENT") {
    return ["STUDENT", patient.faculty, patient.year, patient.language]
      .filter(Boolean)
      .join(" ");
  }
  return occ;
}

function exportDiagnostic(patient) {
  const lines = [];
  const temp = formatExportMetric(patient.ec_temperature);
  if (temp) lines.push("T " + temp + " °C");
  if (patient.ec_blood_pressure) {
    lines.push("TA " + patient.ec_blood_pressure + " mmHg");
  }
  if (patient.ec_pulse) lines.push("P " + patient.ec_pulse + " bătăi/min");
  if (patient.ec_saturation) lines.push("S " + patient.ec_saturation + " %");
  if (patient.ec_diagnosis) lines.push(String(patient.ec_diagnosis));
  return lines.join("\n");
}

function exportPrescriptions(patient) {
  const lines = [];
  pushLabeled(lines, "Recomandări", patient.ec_recommendations);
  pushLabeled(lines, "Vaccinări", patient.ec_vaccinations);
  pushLabeled(lines, "Rețetă integrală", patient.rp_full);
  pushLabeled(lines, "Rețetă gratuită", patient.rp_free);
  for (let n = 1; n <= 3; n++) {
    const cas = [
      patient["bt_cas_" + n + "_serial"],
      patient["bt_cas_" + n + "_specialty"],
      patient["bt_cas_" + n + "_type"],
    ]
      .filter(Boolean)
      .join(" ");
    if (cas) lines.push("BC" + n + " " + cas);
  }
  pushLabeled(lines, "BS", patient.bt_simple);
  if (patient.bt_chronic_specialist) lines.push("Cronici Specialist");
  pushExportDateLines(lines, patient, EXPORT_AE_PRIMARY);
  pushLabeled(lines, "Alt Scop", patient.ae_other_purpose);
  if (patient.ae_medical_scholarship) lines.push("Bursă Medicală");
  if (patient.ae_epidemiologic_clearance) lines.push("Aviz Epidemiologic");
  if (patient.ae_sport_competition) lines.push("Competiții Sportive");
  pushExportDateLines(lines, patient, EXPORT_AE_ABSENCE_ENDORSE);
  const height = formatExportMetric(patient.eb_height);
  const weight = formatExportMetric(patient.eb_weight);
  const bmi = formatExportMetric(patient.eb_bmi);
  const ebMetrics = [
    height && height + "cm",
    patient.eb_height_index && "(" + patient.eb_height_index + ")",
    weight && weight + "kg",
    patient.eb_weight_index && "(" + patient.eb_weight_index + ")",
    bmi && bmi + "imc",
  ].filter(Boolean);
  if (ebMetrics.length) lines.push("EB " + ebMetrics.join(" "));
  if (patient.eb_physical_development) {
    lines.push("DF " + patient.eb_physical_development);
  }
  if (patient.eb_disease_codes_old) lines.push("CBV " + patient.eb_disease_codes_old);
  if (patient.eb_disease_codes_new) lines.push("CBN " + patient.eb_disease_codes_new);
  return lines.join("\n");
}

function rowsForExportOrReport(startDate, endDate) {
  const start = dayBound(startDate, false);
  const end = dayBound(endDate, true);
  const rows = getDataRows().filter((entry) =>
    timestampInRange(start, end, entry.values),
  );
  if (rows.length === 0) {
    return fail("Nu s-au găsit pacienți în intervalul de date specificat");
  }
  return { success: true, rows, start, end };
}

function countDiseaseCodes(value, counts) {
  if (!value) return 0;
  let total = 0;
  for (const part of String(value).split(/\s+/)) {
    const cod = parseInt(part, 10);
    if (cod >= 0 && cod < 1000) {
      counts[cod]++;
      total++;
    }
  }
  return total;
}

function codesList(counts) {
  const list = [];
  for (let cod = 0; cod < 1000; cod++) {
    if (counts[cod] > 0) list.push({ cod, n: counts[cod] });
  }
  return list;
}

function formatCodeInterval(lo, hi) {
  if (lo === hi) return String(lo);
  const pad = (n) => ("000" + n).slice(-3);
  return pad(lo) + "-" + pad(hi);
}

function countCodesInInterval(counts, lo, hi) {
  let n = 0;
  const from = Math.max(0, lo);
  const to = Math.min(999, hi);
  for (let c = from; c <= to; c++) n += counts[c];
  return n;
}

function bumpKeyed(map, allowed, key) {
  if (allowed.includes(key)) map[key]++;
}

function emptyEbGender() {
  const levels = (arr) => Object.fromEntries(arr.map((level) => [level, 0]));
  return {
    total: 0,
    height: levels(CONFIG.BODY_INDEX_LEVELS),
    weight: levels(CONFIG.BODY_INDEX_LEVELS),
    development: levels(CONFIG.PHYSICAL_DEV_LEVELS),
  };
}

function tallyPatients(rows) {
  const langs = emptyLangs();
  const ecCounts = new Uint16Array(1000);
  const ebCounts = new Uint16Array(1000);
  const acts = {
    vaccinations: 0, rp_full: 0, rp_free: 0, bt_cas: 0, bt_simple: 0,
    ae_absence_excuse: 0, ae_absence_endorse: 0, ae_sport_excuse: 0,
    ae_sport_endorse: 0, ae_other_purpose: 0, ae_medical_scholarship: 0,
    ae_epidemiologic_clearance: 0, ae_sport_competition: 0,
    bt_chronic_specialist: 0, ec_codes: 0, eb_codes: 0, eb_codes_old: 0, eb_codes_new: 0,
  };
  const ebByGender = { F: emptyEbGender(), M: emptyEbGender() };

  for (const entry of rows) {
    const p = rowToPatient(entry.values, entry.rowNum);
    addLang(langs, p.language);
    acts.ec_codes += countDiseaseCodes(p.ec_disease_codes, ecCounts);
    acts.eb_codes_old += countDiseaseCodes(p.eb_disease_codes_old, ebCounts);
    acts.eb_codes_new += countDiseaseCodes(p.eb_disease_codes_new, ebCounts);
    REPORT_ACT_FLAGS.forEach(([act, prop]) => {
      if (p[prop]) acts[act]++;
    });
    acts.bt_cas += countIndexed(p, "bt_cas_", "_serial", 1, 3);
    acts.ae_absence_endorse += countIndexed(p, "ae_absence_endorse_", "_start", 1, 5);
    if (p.eb_height) {
      const g = genderCode(p.gender);
      if (g) {
        const s = ebByGender[g];
        s.total++;
        bumpKeyed(s.height, CONFIG.BODY_INDEX_LEVELS, p.eb_height_index);
        bumpKeyed(s.weight, CONFIG.BODY_INDEX_LEVELS, p.eb_weight_index);
        bumpKeyed(s.development, CONFIG.PHYSICAL_DEV_LEVELS, p.eb_physical_development);
      }
    }
  }

  acts.eb_codes = acts.eb_codes_old + acts.eb_codes_new;
  return { langs, ecCounts, ebCounts, acts, ebByGender };
}

function registerTotals() {
  const langs = emptyLangs();
  for (const entry of getDataRows()) {
    if (toTimestampMs(entry.values[CONFIG.TIMESTAMP]) == null) continue;
    addLang(langs, entry.values[CONFIG.LANGUAGE]);
  }
  return langs;
}

function td(style, value) {
  return `<td ${style}>${value == null ? "" : value}</td>`;
}

function tr(inner) {
  return "<tr>" + inner + "</tr>";
}

function twoCells(a, b) {
  return td(REPORT_CELL, a) + td(REPORT_CELL, b);
}

function th(label, span) {
  return `<th${span ? ` colspan="${span}"` : ""} ${REPORT_TH}>${label}</th>`;
}

function htmlTable(cols, inner) {
  const w = (100 / cols).toFixed(4) + "%";
  const group =
    "<colgroup>" + Array(cols).fill('<col style="width:' + w + ';">').join("") + "</colgroup>";
  return `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="${REPORT_TABLE_CSS}">${group}${inner}</table>`;
}

function thRow(labels) {
  return tr(labels.map((label) => th(label)).join(""));
}

function codePair(item) {
  return twoCells(item ? item.cod : "", item ? item.n : "");
}

function ebNums(nF, nM) {
  return td(REPORT_CELL, nF + nM) + td(REPORT_CELL, nF) + td(REPORT_CELL, nM);
}

function codeStatsTable(title, body) {
  return htmlTable(4, tr(th(title, 4)) + thRow(REPORT_CODE_HEADERS) + body);
}

function rowspanBlock(title, levels, labelFn, valueFn) {
  const n = levels.length;
  return levels
    .map((level, i) => {
      const group = i === 0 ? `<td rowspan="${n}" ${REPORT_CELL_GROUP}>${title}</td>` : "";
      return tr(group + td(REPORT_CELL, labelFn(level)) + ebNums(valueFn("F", level), valueFn("M", level)));
    })
    .join("");
}

function developmentBlock(ebByGender) {
  const n = (gender, level) => ebByGender[gender].development[level] || 0;
  const plusF = n("F", "Dizarmonică +G");
  const plusM = n("M", "Dizarmonică +G");
  const minusF = n("F", "Dizarmonică -G");
  const minusM = n("M", "Dizarmonică -G");
  const group = `<td rowspan="4" ${REPORT_CELL_GROUP}>Dezvoltare fizică</td>`;
  return (
    tr(group + td(REPORT_CELL, "Armonică") + ebNums(n("F", "Armonică"), n("M", "Armonică"))) +
    tr(td(REPORT_CELL, "Dizarmonică") + ebNums(plusF + minusF, plusM + minusM)) +
    tr(td(REPORT_CELL, "Dizarmonică +G") + ebNums(plusF, plusM)) +
    tr(td(REPORT_CELL, "Dizarmonică -G") + ebNums(minusF, minusM))
  );
}

function buildReportHtml(tally) {
  const { ecCounts, ebCounts, acts, ebByGender } = tally;
  const ecCodes = codesList(ecCounts);
  const ebCodes = codesList(ebCounts);
  const tipPairs = [
    ["Vaccinări", acts.vaccinations],
    ["Rețete integrale", acts.rp_full],
    ["Rețete gratuite", acts.rp_free],
    ["Bilete CAS", acts.bt_cas],
    ["Bilete simple", acts.bt_simple],
    ["Cronici specialist", acts.bt_chronic_specialist],
    ["Scutiri absențe", acts.ae_absence_excuse],
    ["Vizări scutiri absențe", acts.ae_absence_endorse],
    ["Scutiri sport", acts.ae_sport_excuse],
    ["Vizări scutiri sport", acts.ae_sport_endorse],
    ["Alte scopuri", acts.ae_other_purpose],
    ["Burse medicale", acts.ae_medical_scholarship],
    ["Avize epidemiologice", acts.ae_epidemiologic_clearance],
    ["Competiții sportive", acts.ae_sport_competition],
  ];
  const maxRows = Math.max(tipPairs.length, ecCodes.length, ebCodes.length);
  let bandBody = "";
  for (let i = 0; i < maxRows; i++) {
    const pair = tipPairs[i];
    bandBody += tr(
      twoCells(pair ? pair[0] : "", pair ? pair[1] : "") +
      codePair(ecCodes[i]) +
      codePair(ebCodes[i]),
    );
  }

  const codesHead = (title, sub) =>
    th(
      `${title}<br><span style="font-weight:normal; color:#555;">${sub}</span>`,
      2,
    );

  const band = htmlTable(
    6,
    tr(
      th("Acte eliberate", 2) +
      codesHead("Coduri boală (clinic)", ecCodes.length + " unice · " + acts.ec_codes + " apariții") +
      codesHead(
        "Coduri boală (bilanț)",
        ebCodes.length +
        " unice · " +
        acts.eb_codes +
        " apariții · " +
        acts.eb_codes_old +
        " vechi · " +
        acts.eb_codes_new +
        " noi",
      ),
    ) +
    thRow(["Tip", "Nr.", "Cod", "Nr.", "Cod", "Nr."]) +
    bandBody,
  );

  let incidence = "";
  CONFIG.CLINIC_CODE_INTERVALS.forEach((row, i) => {
    incidence += tr(
      td(REPORT_CELL, i + 1) +
      td(REPORT_CELL, formatCodeInterval(row[0], row[1])) +
      td(REPORT_CELL, row[2]) +
      td(REPORT_CELL, countCodesInInterval(ecCounts, row[0], row[1])),
    );
  });

  let prevalence = "";
  let prevSum = 0;
  CONFIG.PREVALENCE_CODE_INTERVALS.forEach((row) => {
    const n = row[2].reduce(
      (sum, [lo, hi]) => sum + countCodesInInterval(ecCounts, lo, hi),
      0,
    );
    prevSum += n;
    prevalence += tr(
      td(REPORT_CELL, row[0]) +
      td(REPORT_CELL, row[2].map(([lo, hi]) => formatCodeInterval(lo, hi)).join("; ")) +
      td(REPORT_CELL, row[1]) +
      td(REPORT_CELL, n),
    );
  });
  prevalence =
    tr(
      td(REPORT_CELL, 1) +
      td(REPORT_CELL, "Toate") +
      td(REPORT_CELL, "Total") +
      td(REPORT_CELL, prevSum),
    ) +
    prevalence;

  const metricFn = (metric) => (gender, level) => ebByGender[gender][metric][level] || 0;
  const indexLabel = (level) =>
    level + " (" + (CONFIG.BODY_INDEX_LABELS[level] || level) + ")";
  const ebBody =
    tr(th("Examene", 2) + ebNums(ebByGender.F.total, ebByGender.M.total)) +
    developmentBlock(ebByGender) +
    rowspanBlock("Indicatori greutate", CONFIG.BODY_INDEX_LEVELS, indexLabel, metricFn("weight")) +
    rowspanBlock("Indicatori înălțime", CONFIG.BODY_INDEX_LEVELS, indexLabel, metricFn("height"));

  return (
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0; padding:5mm; font-family:Roboto, Arial, Helvetica, sans-serif; font-size:16px; color:#333; width:100%;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:1050px; margin:0 auto; width:100%;"><tr><td style="padding:0;">' +
    band +
    codeStatsTable("Incidența (cazuri noi de îmbolnăvire depistate)", incidence) +
    htmlTable(5, tr(th("Examene bilanț", 5)) + thRow(["Categorie", "Tip", "Total", "Fete", "Băieți"]) + ebBody) +
    codeStatsTable("Morbiditate (prevalența de moment)", prevalence) +
    "</td></tr></table></body></html>"
  );
}

function deleteTriggersNamed(name) {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === name) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function ensureFormSubmitTrigger() {
  deleteTriggersNamed("onFormSubmit");
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
}

function ensureShiftEndTrigger() {
  const email = sessionEmail();
  if (email) {
    PropertiesService.getScriptProperties().setProperty(SHIFT_END_EMAIL, email);
  }
  deleteTriggersNamed("onShiftEnd");
  ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"].forEach((day) => {
    ScriptApp.newTrigger("onShiftEnd")
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay[day])
      .atHour(18)
      .inTimezone(CONFIG.TIMEZONE)
      .create();
  });
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  DriveApp.getFileById(ss.getId());
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  ss.setSpreadsheetLocale(CONFIG.LOCALE);
  const sheet = getSheet();
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < CONFIG.COLUMN_COUNT) {
    sheet.insertColumnsAfter(maxColumns, CONFIG.COLUMN_COUNT - maxColumns);
  } else if (maxColumns > CONFIG.COLUMN_COUNT) {
    sheet.deleteColumns(CONFIG.COLUMN_COUNT + 1, maxColumns - CONFIG.COLUMN_COUNT);
  }

  const keepRows = Math.max(1 + CONFIG.TARGET_DATA_ROWS, sheet.getLastRow());
  const maxRows = sheet.getMaxRows();
  if (maxRows < keepRows) {
    sheet.insertRowsAfter(maxRows, keepRows - maxRows);
  } else if (maxRows > keepRows) {
    sheet.deleteRows(keepRows + 1, maxRows - keepRows);
  }

  const headers = new Array(CONFIG.COLUMN_COUNT);
  forEachColumn((key, index) => {
    headers[index] = key;
  });
  formatDataRows(sheet, 2, keepRows);
  applyRoboto(
    sheet.getRange(1, 1, 1, CONFIG.COLUMN_COUNT).setValues([headers]),
    "bold",
  );
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(CONFIG.GENDER + 1);
  ensureFormSubmitTrigger();
  ensureShiftEndTrigger();
}

function addPatientData(patientData) {
  return withLock(() => {
    const sheet = getSheet();
    const nextRow = Math.max(sheet.getLastRow(), 1) + 1;
    const row = new Array(CONFIG.COLUMN_COUNT).fill("");
    row[CONFIG.TIMESTAMP] = new Date();
    applyIdentityFields(row, patientData);
    writeRow(sheet, nextRow, row);
    return ok("Pacient adăugat cu succes!", rowToPatient(row, nextRow));
  });
}

function editPatientData(patientData) {
  return updatePatientRow(
    patientData,
    applyIdentityFields,
    "Pacient actualizat cu succes!",
    true,
  );
}

function savePatientData(patientData) {
  return updatePatientRow(
    patientData,
    applyClinicalFields,
    "Date salvate cu succes!",
    false,
  );
}

function loadFiltered(include, emptyMsg, stem, suffix) {
  try {
    const allRows = getDataRows();
    if (allRows.length === 0) return info(MSG_NO_PATIENTS, []);
    const patients = [];
    for (let i = allRows.length - 1; i >= 0; i--) {
      const { rowNum, values } = allRows[i];
      if (include(values)) patients.push(rowToPatient(values, rowNum));
    }
    if (patients.length === 0) return info(emptyMsg, patients);
    return ok(pacientOk(patients.length, stem, suffix), patients);
  } catch (error) {
    return fail(errorMessage(error), []);
  }
}

function loadPatientData() {
  const today = formatWith(new Date(), "yyyy-MM-dd");
  const start = dayBound(today, false);
  const end = dayBound(today, true);
  return loadFiltered(
    (values) => timestampInRange(start, end, values),
    "Nu există pacienți înregistrați astăzi",
    "încărca",
    " cu succes",
  );
}

function searchPatientData(searchTerm) {
  const parts = normName(searchTerm).split(/\s+/).filter(Boolean);
  if (parts.length === 0) return info("Introduceți nume pacient", []);
  const q = ' pentru "' + searchTerm + '"';
  return loadFiltered(
    (values) => parts.every((part) => normName(values[CONFIG.NAME]).includes(part)),
    "Nu s-au găsit pacienți" + q,
    "găsi",
    q,
  );
}

function exportPatientData(startDate, endDate) {
  try {
    const found = rowsForExportOrReport(startDate, endDate);
    if (!found.success) return found;

    const exportName =
      EXPORT_SHEET_PREFIX + formatDay(found.start) + " - " + formatDay(found.end);
    const ss = SpreadsheetApp.getActive();
    const existingSheet = ss.getSheetByName(exportName);
    if (existingSheet) ss.deleteSheet(existingSheet);
    const sheet = ss.insertSheet(exportName);
    const rows = found.rows.map((entry) => {
      const patient = rowToPatient(entry.values, entry.rowNum);
      return [
        patient.id,
        asDate(patient.timestamp) || "",
        patient.name,
        patient.age,
        genderCode(patient.gender),
        patient.address,
        occupationExport(patient),
        patient.symptoms,
        exportDiagnostic(patient),
        patient.ec_disease_codes,
        exportPrescriptions(patient),
      ];
    });

    const output = [EXPORT_HEADERS, ...rows];
    const dataRange = sheet.getRange(1, 1, output.length, EXPORT_HEADERS.length);
    applyRoboto(dataRange.setValues(output).setNumberFormat("@"), "normal")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
    if (output.length > 1) {
      sheet.getRange(2, 2, output.length - 1, 1).setNumberFormat("dd.MM.yyyy");
    }
    sheet
      .getRange(1, 1, 1, EXPORT_HEADERS.length)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    EXPORT_COLUMN_WIDTHS.forEach((width, i) => {
      sheet.setColumnWidth(i + 1, width);
    });
    sheet.setFrozenRows(1);
    ss.setActiveSheet(sheet);
    SpreadsheetApp.flush();
    return ok(exportName + " a fost creat cu succes!");
  } catch (error) {
    return fail(errorMessage(error));
  }
}

function reportPatientData(startDate, endDate) {
  try {
    const startTime = Date.now();
    const userEmail = sessionEmail();
    if (!userEmail) return fail("Nu s-a putut determina adresa de email");
    const found = rowsForExportOrReport(startDate, endDate);
    if (!found.success) return found;
    const tally = tallyPatients(found.rows);
    MailApp.sendEmail({
      to: userEmail,
      subject: umfSubject(
        "raport registru medical",
        formatDay(found.start) + " – " + formatDay(found.end),
        found.rows.length,
        tally.langs,
        startTime,
      ),
      htmlBody: buildReportHtml(tally),
      noReply: true,
    });
    return ok("Raportul a fost trimis cu succes la " + userEmail);
  } catch (error) {
    return fail(errorMessage(error));
  }
}

function onFormSubmit(e) {
  try {
    withLock(() => {
      if (!e || !e.range) return ok();
      const sheet = e.range.getSheet();
      if (sheet.getName() !== CONFIG.SHEET_NAME) return ok();
      const row = e.range.getRow();
      const lastRow = sheet.getLastRow();
      const atBottom = row >= lastRow;
      try {
        const values = sheet.getRange(row, 1, 1, CONFIG.COLUMN_COUNT).getValues()[0];
        forEachColumn((key, index) => {
          if (isDateKey(key)) return;
          const value = values[index];
          if (typeof value === "number" && isFinite(value)) values[index] = String(value);
        });
        const dest = atBottom ? row : lastRow + 1;
        writeRow(sheet, dest, values);
        formatDataRows(sheet, dest, dest);
      } catch (error) {
        Logger.log("onFormSubmit: " + errorMessage(error));
      }
      if (!atBottom) sheet.deleteRow(row);
      return ok();
    });
  } catch (error) {
    Logger.log("onFormSubmit: " + errorMessage(error));
  }
}

function removeExportSheets() {
  const ss = SpreadsheetApp.getActive();
  const keep = CONFIG.SHEET_NAME;
  ss.getSheets()
    .filter((s) => {
      const name = s.getName();
      if (name === keep || name.indexOf(EXPORT_SHEET_PREFIX) !== 0) return false;
      try {
        return !s.getFormUrl();
      } catch (error) {
        return false;
      }
    })
    .forEach((s) => ss.deleteSheet(s));
  SpreadsheetApp.flush();
}

function formResponsesXlsx(sheet) {
  const fileName = CONFIG.SHEET_NAME + " " + formatDay(new Date());
  const response = UrlFetchApp.fetch(
    "https://docs.google.com/spreadsheets/d/" +
    sheet.getParent().getId() +
    "/export?format=xlsx&gid=" +
    sheet.getSheetId(),
    {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    },
  );
  const blob = response.getBlob();
  const type = String(blob.getContentType() || "");
  if (response.getResponseCode() !== 200 || type.indexOf("html") !== -1) {
    throw new Error(
      "Export xlsx eșuat (" +
      response.getResponseCode() +
      "): " +
      response.getContentText().slice(0, 300),
    );
  }
  return blob.setName(fileName + ".xlsx").setContentType(MimeType.MICROSOFT_EXCEL);
}

function onShiftEnd() {
  const startTime = Date.now();
  const recipient =
    PropertiesService.getScriptProperties().getProperty(SHIFT_END_EMAIL) ||
    sessionEmail();
  if (!recipient) return;
  try {
    const langs = registerTotals();
    const sheet = getSheet();
    const xlsx = formResponsesXlsx(sheet);
    MailApp.sendEmail({
      to: recipient,
      subject: umfSubject(
        "copie registru medical",
        formatDay(new Date()),
        langs.n,
        langs,
        startTime,
      ),
      body: "",
      attachments: [xlsx],
      noReply: true,
    });
    removeExportSheets();
  } catch (error) {
    Logger.log("onShiftEnd: " + errorMessage(error));
  }
}

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile("Index").setTitle(
      "UMF Registru Medical",
    );
  } catch (error) {
    return ContentService.createTextOutput(errorMessage(error)).setMimeType(
      ContentService.MimeType.TEXT,
    );
  }
}
