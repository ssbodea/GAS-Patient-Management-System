const CONFIG = {
    SHEET_NAME: "Form Responses 1",
    ID: 0, TIMESTAMP: 1, NAME: 2, AGE: 3, GENDER: 4, LMP: 5, ADDRESS: 6, FACULTY: 7, YEAR: 8, LANGUAGE: 9, SYMPTOMS: 10, TREATMENT: 11, DISEASES: 12, ALLERGIES: 13,
    VACCIN: 14, DIAGNOSTIC: 15, CODURI_BOALA: 16, RP_INTEGRALA: 17, RP_GRATUITA: 18, BT_CAS_1: 19, BT_CAS_2: 20, BT_CAS_3: 21, BT_SIMPLU: 22,
    AM_S_ABSENTA: 23, AM_V_ABSENTA_1: 24, AM_V_ABSENTA_2: 25, AM_S_SPORT: 26, AM_V_SPORT: 27, AM_ALT_SCOP: 28, AM_BURSA_MEDICALA: 29, AE_AVIZ_EPIDEMIOLOGIC: 30, EB_INALTIME: 31, EB_GREUTATE: 32, EB_IMC: 33, EB_CODURI_BOALA: 34,
    COLUMN_COUNT: 35,
    ALLOWED_EMAILS: ["ss.bodea@gmail.com", "agent07.marius@gmail.com", "qmaedica@gmail.com"],
};

// ==================== UTILITY FUNCTIONS ====================

function checkUserAuthorization() {
    const userEmail = Session.getActiveUser().getEmail();
    if (!CONFIG.ALLOWED_EMAILS.includes(userEmail)) {
        throw new Error("Acces neautorizat. Email: " + userEmail);
    }
}

function getSheet() {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Foaia ${CONFIG.SHEET_NAME} nu a fost găsită.`);
    return sheet;
}

function formatDateForDisplay(date) {
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const y = date.getFullYear();
    return `${d}/${m}/${y}`;
}

function convertIsoToRo(isoDate) {
    if (!isoDate) return '';
    const [y, m, d] = isoDate.split('-');
    return `${d}/${m}/${y}`;
}

function convertRoToIso(roDate) {
    if (!roDate) return '';
    const [d, m, y] = roDate.split('/');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
}

// ==================== BT CAS FUNCTIONS ====================

function processBtCasForSave(Cod, Specialitate, Tip) {
    const parts = [Cod, Specialitate, Tip].filter(p => p && p.trim());
    return parts.join(' | ');
}

function parseBtCasForDisplay(fieldValue) {
    if (!fieldValue || typeof fieldValue !== 'string') return ['', '', ''];
    const trimmed = fieldValue.trim();
    return trimmed.includes('|')
        ? trimmed.split('|').map(p => p.trim()).slice(0, 3)
        : [trimmed, '', ''];
}

// ==================== DATE RANGE FUNCTIONS ====================

function processDateRangeForSave(start, end) {
    if (!start && !end) return '';
    if (!start || !end) return start || end || '';

    const startRO = convertIsoToRo(start);
    const endRO = convertIsoToRo(end);

    return `${startRO}-${endRO}`;
}

function parseDateRangeForDisplay(rangeStr) {
    const defaultResult = { start: '', end: '' };
    if (!rangeStr) return defaultResult;
    if (!rangeStr.includes('-')) return defaultResult;

    const [startRO, endRO] = rangeStr.split('-').map(s => s.trim());

    return {
        start: convertRoToIso(startRO),
        end: convertRoToIso(endRO)
    };
}

// ==================== BINARY SEARCH FUNCTIONS ====================

function findFirstRowByDateTime(targetDateTime) {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return -1;

    const timestamps = sheet.getRange(2, CONFIG.TIMESTAMP + 1, lastRow - 1, 1).getValues();
    const targetTime = new Date(targetDateTime).getTime();

    let left = 0;
    let right = timestamps.length - 1;
    let result = -1;

    while (left <= right) {
        const mid = Math.floor((left + right) / 2);
        const midDate = timestamps[mid][0];

        if (!(midDate instanceof Date)) {
            left = mid + 1;
            continue;
        }

        const midTime = midDate.getTime();

        if (midTime >= targetTime) {
            result = mid;
            right = mid - 1;
        } else {
            left = mid + 1;
        }
    }

    return result !== -1 ? result + 2 : -1;
}

function findLastRowByDateTime(targetDateTime) {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return -1;

    const timestamps = sheet.getRange(2, CONFIG.TIMESTAMP + 1, lastRow - 1, 1).getValues();
    const targetTime = new Date(targetDateTime).getTime();

    let left = 0;
    let right = timestamps.length - 1;
    let result = -1;

    while (left <= right) {
        const mid = Math.floor((left + right) / 2);
        const midDate = timestamps[mid][0];

        if (!(midDate instanceof Date)) {
            left = mid + 1;
            continue;
        }

        const midTime = midDate.getTime();

        if (midTime <= targetTime) {
            result = mid;
            left = mid + 1;
        } else {
            right = mid - 1;
        }
    }

    return result !== -1 ? result + 2 : -1;
}

// ==================== PATIENT DATA FUNCTIONS ====================

function createPatientFromRow(row) {
    const [btCas1Cod, btCas1Specialitate, btCas1Tip] = parseBtCasForDisplay(row[CONFIG.BT_CAS_1]);
    const [btCas2Cod, btCas2Specialitate, btCas2Tip] = parseBtCasForDisplay(row[CONFIG.BT_CAS_2]);
    const [btCas3Cod, btCas3Specialitate, btCas3Tip] = parseBtCasForDisplay(row[CONFIG.BT_CAS_3]);

    const amSAbsenta = parseDateRangeForDisplay(row[CONFIG.AM_S_ABSENTA]);
    const amVAbsenta1 = parseDateRangeForDisplay(row[CONFIG.AM_V_ABSENTA_1]);
    const amVAbsenta2 = parseDateRangeForDisplay(row[CONFIG.AM_V_ABSENTA_2]);
    const amSSport = parseDateRangeForDisplay(row[CONFIG.AM_S_SPORT]);
    const amVSport = parseDateRangeForDisplay(row[CONFIG.AM_V_SPORT]);

    return {
        id: row[CONFIG.ID] || '',
        timestamp: row[CONFIG.TIMESTAMP] instanceof Date ? formatDateForDisplay(row[CONFIG.TIMESTAMP]) : '',
        name: row[CONFIG.NAME] || '',
        age: row[CONFIG.AGE] || '',
        gender: row[CONFIG.GENDER] || '',
        lmp: row[CONFIG.LMP] instanceof Date ? formatDateForDisplay(row[CONFIG.LMP]) : '',
        address: row[CONFIG.ADDRESS] || '',
        faculty: row[CONFIG.FACULTY] || '',
        year: row[CONFIG.YEAR] || '',
        language: row[CONFIG.LANGUAGE] || '',
        symptoms: row[CONFIG.SYMPTOMS] || '',
        treatment: row[CONFIG.TREATMENT] || '',
        diseases: row[CONFIG.DISEASES] || '',
        allergies: row[CONFIG.ALLERGIES] || '',
        vaccin: row[CONFIG.VACCIN] || '',
        diagnostic: row[CONFIG.DIAGNOSTIC] || '',
        coduriBoala: row[CONFIG.CODURI_BOALA] || '',
        rpIntegrala: row[CONFIG.RP_INTEGRALA] || '',
        rpGratuita: row[CONFIG.RP_GRATUITA] || '',
        btSimplu: row[CONFIG.BT_SIMPLU] || '',
        btCas1Cod, btCas1Specialitate, btCas1Tip,
        btCas2Cod, btCas2Specialitate, btCas2Tip,
        btCas3Cod, btCas3Specialitate, btCas3Tip,
        amSAbsentaStart: amSAbsenta.start || '',
        amSAbsentaEnd: amSAbsenta.end || '',
        amVAbsenta1Start: amVAbsenta1.start || '',
        amVAbsenta1End: amVAbsenta1.end || '',
        amVAbsenta2Start: amVAbsenta2.start || '',
        amVAbsenta2End: amVAbsenta2.end || '',
        amSSportStart: amSSport.start || '',
        amSSportEnd: amSSport.end || '',
        amVSportStart: amVSport.start || '',
        amVSportEnd: amVSport.end || '',
        amAltScop: row[CONFIG.AM_ALT_SCOP] || '',
        amBursaMedicala: row[CONFIG.AM_BURSA_MEDICALA] || '',
        aeAvizEpidemiologic: row[CONFIG.AE_AVIZ_EPIDEMIOLOGIC] || '',
        ebInaltime: row[CONFIG.EB_INALTIME] || '',
        ebGreutate: row[CONFIG.EB_GREUTATE] || '',
        ebIMC: row[CONFIG.EB_IMC] || '',
        ebCoduriBoala: row[CONFIG.EB_CODURI_BOALA] || ''
    };
}

// ==================== TRIGGER FUNCTIONS ====================

function onFormSubmit(e) {
    const sheet = getSheet();
    const row = e.range.getRow();

    if (row < 2) {
        sheet.deleteRow(row);
        return;
    }

    try {
        sheet.getRange(row, CONFIG.ID + 1).setValue(row - 1);
        const timestampCell = sheet.getRange(row, CONFIG.TIMESTAMP + 1);
        timestampCell.setValue(new Date());
        timestampCell.setNumberFormat('dd/mm/yyyy hh:mm:ss');

        const lmpValue = e.namedValues['LMP']?.[0];
        if (lmpValue) {
            const lmpDate = new Date(lmpValue);
            if (!isNaN(lmpDate.getTime())) {
                const lmpCell = sheet.getRange(row, CONFIG.LMP + 1);
                lmpCell.setValue(lmpDate);
                lmpCell.setNumberFormat('dd/mm/yyyy');
            }
        }

    } catch (error) {
        sheet.deleteRow(row);
    }
}

function fixAllIds() {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) return;

    const ids = [];
    for (let i = 1; i <= lastRow - 1; i++) {
        ids.push([i]);
    }

    sheet.getRange(2, CONFIG.ID + 1, lastRow - 1, 1).setValues(ids);
}

// ==================== API FUNCTIONS ====================

function doGet() {
    try {
        return HtmlService.createHtmlOutputFromFile('Index').setTitle('UMF Registru Medical');
    } catch (e) {
        return ContentService.createTextOutput(e.message).setMimeType(ContentService.MimeType.TEXT);
    }
}

function loadPatientData() {
    try {
        checkUserAuthorization();
        const sheet = getSheet();
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
            return {
                success: true,
                data: [],
                message: "Nu există pacienți în baza de date"
            };
        }

        const allRows = sheet.getRange(2, 1, lastRow - 1, CONFIG.COLUMN_COUNT).getValues();
        const patients = [];

        const today = new Date();
        const todayDay = today.getDate();
        const todayMonth = today.getMonth();
        const todayYear = today.getFullYear();

        for (let i = allRows.length - 1; i >= 0; i--) {
            const row = allRows[i];
            const timestamp = row[CONFIG.TIMESTAMP];

            if (!(timestamp instanceof Date)) continue;

            if (timestamp.getDate() === todayDay &&
                timestamp.getMonth() === todayMonth &&
                timestamp.getFullYear() === todayYear) {
                patients.push(createPatientFromRow(row));
            } else {
                return {
                    success: true,
                    data: patients,
                    message: patients.length === 0
                        ? "Nu există pacienți înregistrați astăzi"
                        : `${patients.length} pacient${patients.length === 1 ? '' : 'i'} încărcați cu succes`
                };
            }
        }

        return {
            success: true,
            data: patients,
            message: patients.length === 0
                ? "Nu există pacienți înregistrați astăzi"
                : `${patients.length} pacient${patients.length === 1 ? '' : 'i'} încărcați cu succes`
        };

    } catch (error) {
        return {
            success: false,
            data: [],
            message: error.message
        };
    }
}

function savePatientData(patientData) {
    try {
        checkUserAuthorization();
        const sheet = getSheet();
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) {
            return {
                success: false,
                message: "Nu există pacienți în baza de date"
            };
        }

        const patientId = parseInt(patientData.id, 10);
        if (isNaN(patientId) || patientId < 1) {
            return {
                success: false,
                message: `ID invalid: ${patientData.id}`
            };
        }

        const rowNum = patientId + 1;
        if (rowNum > lastRow) {
            return {
                success: false,
                message: `Pacientul cu ID-ul ${patientData.id} nu a fost găsit`
            };
        }

        const updateData = [
            patientData.vaccin || '',
            patientData.diagnostic || '',
            patientData.coduriBoala || '',
            patientData.rpIntegrala || '',
            patientData.rpGratuita || '',
            processBtCasForSave(patientData.btCas1Cod, patientData.btCas1Specialitate, patientData.btCas1Tip),
            processBtCasForSave(patientData.btCas2Cod, patientData.btCas2Specialitate, patientData.btCas2Tip),
            processBtCasForSave(patientData.btCas3Cod, patientData.btCas3Specialitate, patientData.btCas3Tip),
            patientData.btSimplu || '',
            processDateRangeForSave(patientData.amSAbsentaStart, patientData.amSAbsentaEnd),
            processDateRangeForSave(patientData.amVAbsenta1Start, patientData.amVAbsenta1End),
            processDateRangeForSave(patientData.amVAbsenta2Start, patientData.amVAbsenta2End),
            processDateRangeForSave(patientData.amSSportStart, patientData.amSSportEnd),
            processDateRangeForSave(patientData.amVSportStart, patientData.amVSportEnd),
            patientData.amAltScop || '',
            patientData.amBursaMedicala ? true : '',
            patientData.aeAvizEpidemiologic ? true : '',
            patientData.ebInaltime || '',
            patientData.ebGreutate || '',
            patientData.ebIMC || '',
            patientData.ebCoduriBoala || ''
        ];

        const startCol = CONFIG.VACCIN + 1;
        sheet.getRange(rowNum, startCol, 1, updateData.length).setValues([updateData]);
        SpreadsheetApp.flush();

        return {
            success: true,
            message: "Date salvate cu succes!"
        };

    } catch (error) {
        return {
            success: false,
            message: error.message
        };
    }
}

function searchPatientData(searchTerm) {
    try {
        checkUserAuthorization();

        const sheet = getSheet();
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) {
            return {
                success: true,
                data: [],
                message: "Nu există pacienți în baza de date"
            };
        }

        const searchLower = searchTerm.toString().trim().toLowerCase();
        const searchParts = searchLower.split(/\s+/).filter(part => part.length > 0);

        const nameData = sheet.getRange(2, CONFIG.NAME + 1, lastRow - 1, 1).getValues();
        const matchingRowNumbers = [];

        for (let i = nameData.length - 1; i >= 0; i--) {
            const name = (nameData[i][0] || '').toString().toLowerCase();

            if (searchParts.every(part => name.includes(part))) {
                matchingRowNumbers.push(i + 2);
            }
        }

        if (matchingRowNumbers.length === 0) {
            return {
                success: true,
                data: [],
                message: `Nu s-au găsit pacienți pentru "${searchTerm}"`
            };
        }

        const minRow = Math.min(...matchingRowNumbers);
        const maxRow = Math.max(...matchingRowNumbers);
        const rowCount = maxRow - minRow + 1;

        const allRows = sheet.getRange(minRow, 1, rowCount, CONFIG.COLUMN_COUNT).getValues();

        const rowMap = new Map();
        for (let i = 0; i < allRows.length; i++) {
            rowMap.set(minRow + i, allRows[i]);
        }

        const patients = [];
        for (let i = matchingRowNumbers.length - 1; i >= 0; i--) {
            const rowData = rowMap.get(matchingRowNumbers[i]);
            if (rowData) {
                patients.push(createPatientFromRow(rowData));
            }
        }

        return {
            success: true,
            data: patients,
            message: patients.length === 1
                ? `1 pacient găsit pentru "${searchTerm}"`
                : `${patients.length} pacienți găsiți pentru "${searchTerm}"`
        };

    } catch (error) {
        return {
            success: false,
            data: [],
            message: error.message
        };
    }
}

function exportPatientData(startDate, endDate) {
    try {
        checkUserAuthorization();

        const firstRow = findFirstRowByDateTime(startDate);
        const lastRow = findLastRowByDateTime(endDate);

        if (firstRow === -1 || lastRow === -1) {
            return {
                success: false,
                message: "Nu s-au găsit pacienți în intervalul de date specificat"
            };
        }

        const sourceSheet = getSheet();
        const rowCount = lastRow - firstRow + 1;
        const filteredData = sourceSheet.getRange(firstRow, 1, rowCount, CONFIG.COLUMN_COUNT).getValues();

        const start = new Date(startDate);
        const end = new Date(endDate);
        const exportName = `Export_${formatDateForDisplay(start)}_${formatDateForDisplay(end)}`;

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const existingSheet = ss.getSheetByName(exportName);
        existingSheet && ss.deleteSheet(existingSheet);

        const sheet = ss.insertSheet(exportName);
        const rows = [];

        for (let i = 0; i < filteredData.length; i++) {
            const row = filteredData[i];

            const timestamp = row[CONFIG.TIMESTAMP];
            const dateOnly = timestamp instanceof Date ? formatDateForDisplay(timestamp) : '';

            let prescriptionsText = '';

            row[CONFIG.VACCIN] && (prescriptionsText += `Vaccin ${row[CONFIG.VACCIN]}\n`);
            row[CONFIG.RP_INTEGRALA] && (prescriptionsText += `RP Integrală ${row[CONFIG.RP_INTEGRALA]}\n`);
            row[CONFIG.RP_GRATUITA] && (prescriptionsText += `RP Gratuită ${row[CONFIG.RP_GRATUITA]}\n`);
            row[CONFIG.BT_CAS_1] && (prescriptionsText += `CAS1 ${row[CONFIG.BT_CAS_1]}\n`);
            row[CONFIG.BT_CAS_2] && (prescriptionsText += `CAS2 ${row[CONFIG.BT_CAS_2]}\n`);
            row[CONFIG.BT_CAS_3] && (prescriptionsText += `CAS3 ${row[CONFIG.BT_CAS_3]}\n`);
            row[CONFIG.BT_SIMPLU] && (prescriptionsText += `BTS ${row[CONFIG.BT_SIMPLU]}\n`);
            row[CONFIG.AM_S_ABSENTA] && (prescriptionsText += `Scutire Absență ${row[CONFIG.AM_S_ABSENTA]}\n`);
            row[CONFIG.AM_V_ABSENTA_1] && (prescriptionsText += `Vizare Absență1 ${row[CONFIG.AM_V_ABSENTA_1]}\n`);
            row[CONFIG.AM_V_ABSENTA_2] && (prescriptionsText += `Vizare Absență2 ${row[CONFIG.AM_V_ABSENTA_2]}\n`);
            row[CONFIG.AM_S_SPORT] && (prescriptionsText += `Scutire Sport ${row[CONFIG.AM_S_SPORT]}\n`);
            row[CONFIG.AM_V_SPORT] && (prescriptionsText += `Vizare Sport ${row[CONFIG.AM_V_SPORT]}\n`);
            row[CONFIG.AM_ALT_SCOP] && (prescriptionsText += `Alt Scop ${row[CONFIG.AM_ALT_SCOP]}\n`);
            row[CONFIG.AM_BURSA_MEDICALA] && (prescriptionsText += 'Bursă Medicală\n');
            row[CONFIG.AE_AVIZ_EPIDEMIOLOGIC] && (prescriptionsText += 'Aviz Epidemiologic\n');
            row[CONFIG.EB_INALTIME] && (prescriptionsText += `EB ${row[CONFIG.EB_INALTIME]}cm ${row[CONFIG.EB_GREUTATE]}kg ${row[CONFIG.EB_IMC]}imc CB ${row[CONFIG.EB_CODURI_BOALA]}\n`);

            prescriptionsText.endsWith('\n') && (prescriptionsText = prescriptionsText.slice(0, -1));

            rows.push([
                row[CONFIG.ID],
                dateOnly,
                row[CONFIG.NAME] || '',
                row[CONFIG.AGE] || '',
                row[CONFIG.GENDER] || '',
                row[CONFIG.ADDRESS] || '',
                `${row[CONFIG.FACULTY] || ''} ${row[CONFIG.YEAR] || ''} ${row[CONFIG.LANGUAGE] || ''}`.trim(),
                row[CONFIG.SYMPTOMS] || '',
                row[CONFIG.DIAGNOSTIC] || '',
                row[CONFIG.CODURI_BOALA] || '',
                prescriptionsText
            ]);
        }

        const headers = [
            'Nr. crt.', 'Ziua', 'Numele și prenumele', 'Vârsta', 'Sexul',
            'Domiciliul', 'Ocupație', 'Simptome', 'Diagnostic', 'Cod',
            'Prescripții med., analize, adev. med., trat. etc.'
        ];

        const output = [headers, ...rows];
        const dataRange = sheet.getRange(1, 1, output.length, headers.length);
        dataRange.setValues(output);

        dataRange.setFontSize(11);
        dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        dataRange.setHorizontalAlignment("left");
        dataRange.setVerticalAlignment("middle");

        sheet.getRange(1, 1, 1, headers.length)
            .setFontWeight("bold")
            .setHorizontalAlignment("center");

        //const columnWidths = [55, 113, 151, 57, 57, 302, 113, 208, 454, 151, 454];
        const columnWidths = [37, 86, 127, 20, 18, 127, 70, 127, 127, 36, 294];
        for (let col = 0; col < headers.length; col++) {
            sheet.setColumnWidth(col + 1, columnWidths[col]);
        }

        sheet.setFrozenRows(1);

        return {
            success: true,
            message: `${exportName} a fost creat cu succes!`
        };

    } catch (error) {
        return {
            success: false,
            message: error.message
        };
    }
}

function reportPatientData(startDate, endDate) {
    try {
        checkUserAuthorization();

        const startTime = Date.now();
        const userEmail = Session.getActiveUser().getEmail();

        const firstRow = findFirstRowByDateTime(startDate);
        const lastRow = findLastRowByDateTime(endDate);

        if (firstRow === -1 || lastRow === -1) {
            return {
                success: false,
                message: "Nu s-au găsit pacienți în intervalul de date specificat"
            };
        }

        const sourceSheet = getSheet();
        const rowCount = lastRow - firstRow + 1;
        const filteredData = sourceSheet.getRange(firstRow, 1, rowCount, CONFIG.COLUMN_COUNT).getValues();

        const CodCounts = new Int32Array(1000);
        const ebCodCounts = new Int32Array(1000);

        let vaccin = 0, rpIntegrala = 0, rpGratuita = 0, btCas = 0, btSimplu = 0;
        let amScutireAbsenta = 0, amVizareAbsenta = 0, amScutireSport = 0, amVizareSport = 0;
        let amAltScop = 0, amBursaMedicala = 0, aeAvizEpidemiologic = 0, ebTotal = 0;

        let coduriBoalaTotal = 0, ebCoduriBoalaTotal = 0;

        for (let i = 0; i < filteredData.length; i++) {
            const row = filteredData[i];

            const coduriBoala = row[CONFIG.CODURI_BOALA];
            if (coduriBoala) {
                const Cods = String(coduriBoala).split(/\s+/);
                for (let j = 0; j < Cods.length; j++) {
                    const Cod = parseInt(Cods[j], 10);
                    if (Cod >= 0 && Cod < 1000) {
                        CodCounts[Cod]++;
                        coduriBoalaTotal++;
                    }
                }
            }

            const ebCoduriBoala = row[CONFIG.EB_CODURI_BOALA];
            if (ebCoduriBoala) {
                const Cods = String(ebCoduriBoala).split(/\s+/);
                for (let j = 0; j < Cods.length; j++) {
                    const Cod = parseInt(Cods[j], 10);
                    if (Cod >= 0 && Cod < 1000) {
                        ebCodCounts[Cod]++;
                        ebCoduriBoalaTotal++;
                    }
                }
            }

            row[CONFIG.VACCIN] && vaccin++;
            row[CONFIG.RP_INTEGRALA] && rpIntegrala++;
            row[CONFIG.RP_GRATUITA] && rpGratuita++;
            row[CONFIG.BT_CAS_1] && btCas++;
            row[CONFIG.BT_CAS_2] && btCas++;
            row[CONFIG.BT_CAS_3] && btCas++;
            row[CONFIG.BT_SIMPLU] && btSimplu++;
            row[CONFIG.AM_S_ABSENTA] && amScutireAbsenta++;
            row[CONFIG.AM_V_ABSENTA_1] && amVizareAbsenta++;
            row[CONFIG.AM_V_ABSENTA_2] && amVizareAbsenta++;
            row[CONFIG.AM_S_SPORT] && amScutireSport++;
            row[CONFIG.AM_V_SPORT] && amVizareSport++;
            row[CONFIG.AM_ALT_SCOP] && amAltScop++;
            row[CONFIG.AM_BURSA_MEDICALA] && amBursaMedicala++;
            row[CONFIG.AE_AVIZ_EPIDEMIOLOGIC] && aeAvizEpidemiologic++;
            row[CONFIG.EB_INALTIME] && ebTotal++;
        }

        let uniqueCodes = 0;
        let ebUniqueCodes = 0;
        let CodesHtml = '';
        let ebCodesHtml = '';

        for (let Cod = 0; Cod < 1000; Cod++) {
            const count = CodCounts[Cod];
            if (count > 0) {
                uniqueCodes++;
                CodesHtml += `<tr><td>${Cod}</td><td>${count}</td></tr>`;
            }
            const ebCount = ebCodCounts[Cod];
            if (ebCount > 0) {
                ebUniqueCodes++;
                ebCodesHtml += `<tr><td>${Cod}</td><td>${ebCount}</td></tr>`;
            }
        }

        const totalAbsente = amScutireAbsenta + amVizareAbsenta;
        const totalSport = amScutireSport + amVizareSport;
        const duration = ((Date.now() - startTime) / 1000).toFixed(2);

        const start = new Date(startDate);
        const end = new Date(endDate);
        const startFormatted = formatDateForDisplay(start);
        const endFormatted = formatDateForDisplay(end);

        const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin:0; padding:5mm; font-family:Arial, Helvetica, sans-serif; font-size:18px; color:#333; width:100%;">

  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:1050px; margin:0 auto; width:100%;">
    <tr>
      <td style="padding:0;">
        
        <p style="margin:0 0 30px 0; font-size:18px; font-weight:bold;">
          📊 Perioadă raportări ${startFormatted} -> ${endFormatted} | Total pacienți ${filteredData.length}
        </p>

        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td width="350" style="width:350px; vertical-align:top; padding-right:30px;">
              <p style="margin:0 0 10px 0; font-size:18px; font-weight:bold;">📝 Prescripții & Totaluri</p>
              <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; border:1px solid #ddd; table-layout:fixed;">
                <tr>
                  <th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center; width:70%;">Tipuri de prescripții</th>
                  <th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center; width:30%;">Nr. apariții</th>
                </tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">Vaccinări</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${vaccin}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">RP Integrale</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${rpIntegrala}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">RP Gratuite</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${rpGratuita}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">BT CAS (1-3)</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${btCas}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">BT Simple</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${btSimplu}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">AM Scutiri Absențe</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${amScutireAbsenta}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">AM Vizări Absențe (1-2)</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${amVizareAbsenta}</td></tr>
                <tr style="background:#f9f9f9; font-weight:bold;">
                  <td style="padding:8px; border:1px solid #ddd; background:#f9f9f9; text-align:center;">AM Total Absențe</td>
                  <td style="padding:8px; border:1px solid #ddd; background:#f9f9f9; text-align:center;">${totalAbsente}</td>
                </tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">AM Scutiri Sport</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${amScutireSport}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">AM Vizări Sport</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${amVizareSport}</td></tr>
                <tr style="background:#f9f9f9; font-weight:bold;">
                  <td style="padding:8px; border:1px solid #ddd; background:#f9f9f9; text-align:center;">AM Total Sport</td>
                  <td style="padding:8px; border:1px solid #ddd; background:#f9f9f9; text-align:center;">${totalSport}</td>
                </tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">AM Alte Scopuri</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${amAltScop}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">AM Burse Medicale</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${amBursaMedicala}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">AE Avize Epidemiologice</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${aeAvizEpidemiologic}</td></tr>
                <tr><td style="padding:8px; border:1px solid #ddd; text-align:center;">EB Examene Bilanț</td><td style="padding:8px; border:1px solid #ddd; text-align:center;">${ebTotal}</td></tr>
              </table>
            </td>
            
            <td width="350" style="width:350px; vertical-align:top; padding-right:30px;">
              <p style="margin:0 0 10px 0; font-size:18px; font-weight:bold;">
                📋 Coduri boală
                <span style="font-size:18px; color:#555; font-weight:normal; margin-left:8px;">
                  (${uniqueCodes} unice, ${coduriBoalaTotal} apariții)
                </span>
              </p>
              ${CodesHtml ?
                `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; border:1px solid #ddd; table-layout:fixed;">
                <tr>
                  <th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center; width:50%;">Cod</th>
                  <th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center; width:50%;">Nr. apariții</th>
                </tr>
                ${CodesHtml.replace(/<td/g, '<td style="padding:8px; border:1px solid #ddd; text-align:center;"').replace(/<th/g, '<th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center;"')}
              </table>` :
                '<p style="color:#666; font-style:italic;">Niciun cod</p>'}
            </td>
            
            <td width="350" style="width:350px; vertical-align:top;">
              <p style="margin:0 0 10px 0; font-size:18px; font-weight:bold;">
                📋 EB Coduri boală
                <span style="font-size:18px; color:#555; font-weight:normal; margin-left:8px;">
                  (${ebUniqueCodes} unice, ${ebCoduriBoalaTotal} apariții)
                </span>
              </p>
              ${ebCodesHtml ?
                `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; border:1px solid #ddd; table-layout:fixed;">
                <tr>
                  <th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center; width:50%;">Cod</th>
                  <th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center; width:50%;">Nr. apariții</th>
                </tr>
                ${ebCodesHtml.replace(/<td/g, '<td style="padding:8px; border:1px solid #ddd; text-align:center;"').replace(/<th/g, '<th style="background:#f2f2f2; padding:8px; border:1px solid #ddd; text-align:center;"')}
              </table>` :
                '<p style="color:#666; font-style:italic;">Niciun cod</p>'}
            </td>
          </tr>
        </table>

        <p style="margin-top:30px; color:#666; font-style:italic; border-top:1px solid #ddd; padding-top:15px; font-size:18px;">
          ⏱️ Generat în ${duration} secunde
        </p>

      </td>
    </tr>
  </table>

</body>
</html>`

        MailApp.sendEmail({
            to: userEmail,
            subject: `UMF Raport Registru Medical (${startFormatted} - ${endFormatted})`,
            htmlBody: htmlBody,
            noReply: true
        });

        return {
            success: true,
            message: `Raportul a fost trimis cu succes la ${userEmail}`
        };

    } catch (error) {
        return {
            success: false,
            message: error.message
        };
    }
}

function deletePatientData(password) {
    try {
        checkUserAuthorization();

        const correctPassword = PropertiesService.getScriptProperties().getProperty('DELETE_PASSWORD');
        if (!correctPassword) {
            return {
                success: false,
                message: "Parola pentru ștergere nu este configurată"
            };
        }

        if (password !== correctPassword) {
            return {
                success: false,
                message: "Parolă incorectă"
            };
        }

        const sheet = getSheet();
        const lastRow = sheet.getLastRow();

        if (lastRow <= 1) {
            return {
                success: false,
                message: "Nu există date de șters"
            };
        }

        sheet.deleteRows(2, lastRow - 1);

        return {
            success: true,
            message: "Toate datele au fost șterse cu succes!"
        };

    } catch (error) {
        return {
            success: false,
            message: error.message
        };
    }
}