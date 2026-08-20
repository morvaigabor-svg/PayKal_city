/**
 * ============================================================================
 * PAYKAL MASTER SHEET - CSOPORT ÉS VÁROSI KIMUTATÁS GENERÁLÓ SCRIPT
 * ============================================================================
 */

const MASTER_SHEET_ID = "1vjjkOriMEJVKBPzY_k3NNJItmLQySWZTQCCZHiovDYE";
const MASTER_SHEET_URL = "https://docs.google.com/spreadsheets/d/" + MASTER_SHEET_ID + "/edit";

// ⚠️ IDE MÁSOLD BE A SABLON TÁBLÁZATOD ID-JÁT:
const TEMPLATE_SHEET_ID = "1k4McbTpZZctMQ4IghiuDYNtLxt1HXRK595RZ3cpMWmg"; 

/**
 * HTML Párbeszédablak megnyitása
 */
function openReportGeneratorDialog() {
  const html = HtmlService.createTemplateFromFile('ReportDialog')
    .evaluate()
    .setWidth(1000)
    .setHeight(680)
    .setTitle('Új Kimutatás Létrehozása (Csoport / Város)');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Új Kimutatás Létrehozása');
}

/**
 * OAuth Token biztosítása a Google Picker API-nak
 */
function getOAuthToken() {
  DriveApp.getRootFolder(); // Kikényszeríti a Drive jogosultságot
  return ScriptApp.getOAuthToken();
}

/**
 * Aktív csoportok és városok kigyűjtése a "Csoportok" munkalapról
 * A oszlop: Csoport ID | B oszlop: Város | E oszlop: Státusz (Aktív)
 */
function getReportSelectionData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Csoportok');
  
  const result = {
    groups: [],
    cities: [],
    cityToGroups: {}
  };

  if (!sheet) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  // A oszlop (1): Csoport ID | B oszlop (2): Város | E oszlop (5): Státusz
  const range = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const citySet = new Set();

  range.forEach(row => {
    const groupId = String(row[0] || "").trim();
    const varos = String(row[1] || "").trim();
    const statusz = String(row[4] || "").trim();

    if (statusz.toLowerCase() === 'aktív') {
      if (groupId) {
        result.groups.push(groupId);
      }
      if (varos) {
        citySet.add(varos);
        if (!result.cityToGroups[varos]) {
          result.cityToGroups[varos] = [];
        }
        if (groupId) {
          result.cityToGroups[varos].push(groupId);
        }
      }
    }
  });

  result.cities = Array.from(citySet).sort();
  result.groups.sort();

  return result;
}

/**
 * A HTML Modal-ból meghívható generáló függvény (Csoport és Város kimutatáshoz egyaránt)
 */
function createGroupReportSheet(payload) {
  try {
    const reportType = payload.reportType || 'group'; // 'group' vagy 'city'
    const selectedValue = payload.selectedValue || payload.groupId;
    const fileName = payload.fileName;
    const targetFolderId = payload.targetFolderId;

    if (!selectedValue || !fileName) {
      throw new Error("A Csoport / Város kiválasztása és a Fájlnév megadása kötelező!");
    }

    if (!TEMPLATE_SHEET_ID || TEMPLATE_SHEET_ID.includes("IDE_JÖJJÖN")) {
      throw new Error("A TEMPLATE_SHEET_ID nincs beállítva a generáló scriptben!");
    }

    // QUERY feltétel összeállítása (Where clause)
    let whereClause = "";

    if (reportType === 'group') {
      whereClause = `Col1 = '${selectedValue}'`;
    } else if (reportType === 'city') {
      // Városhoz tartozó aktív Csoport ID-k lekérése
      const selectionData = getReportSelectionData();
      const groupIds = selectionData.cityToGroups[selectedValue] || [];

      if (groupIds.length === 0) {
        throw new Error(`A(z) '${selectedValue}' városhoz nem található egyetlen aktív csoport sem!`);
      }

      // RegEx minta 'matches' opcióhoz: CS01|CS02|CS03
      const pattern = groupIds.join('|');
      whereClause = `Col1 matches '${pattern}'`;
    }

    // 1. Sablon fájl kérése és a Célmappa kijelölése
    const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
    const targetFolder = (targetFolderId && targetFolderId !== 'root') 
      ? DriveApp.getFolderById(targetFolderId) 
      : DriveApp.getRootFolder();

    // 2. Sablon másolása a kívánt névvel és mappába
    const newFile = templateFile.makeCopy(fileName, targetFolder);
    const newSS = SpreadsheetApp.openById(newFile.getId());

    // 3. QUERY képletek beállítása
    const sheetSzamolotabla = newSS.getSheetByName("Számolótábla");
    if (sheetSzamolotabla) {
      sheetSzamolotabla.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Számolótábla!A2:P"); "SELECT * WHERE ${whereClause}"; 0)`
      );
    }

    const sheetKoltsegek = newSS.getSheetByName("Költségek");
    if (sheetKoltsegek) {
      sheetKoltsegek.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Költségek!A2:O"); "SELECT Col1, Col2, Col3, Col4, Col5, Col7, Col8, Col9, Col10, Col11, Col12, Col13, Col14, Col15 WHERE ${whereClause}"; 0)`
      );
    }

    const sheetBevetelek = newSS.getSheetByName("Bevételek");
    if (sheetBevetelek) {
      sheetBevetelek.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Bevételek!A2:L"); "SELECT * WHERE ${whereClause}"; 0)`
      );
    }

    const sheetPenzmozgasok = newSS.getSheetByName("Pénzmozgások");
    if (sheetPenzmozgasok) {
      sheetPenzmozgasok.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Pénzmozgások!A2:L"); "SELECT * WHERE ${whereClause}"; 0)`
      );
    }

    const sheetProjektek = newSS.getSheetByName("Projektek");
    if (sheetProjektek) {
      sheetProjektek.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Projektek!A2:F"); "SELECT * WHERE ${whereClause}"; 0)`
      );
    }

    return {
      success: true,
      fileUrl: newSS.getUrl(),
      fileName: fileName
    };

  } catch (err) {
    Logger.log("Hiba a táblázat generálása során: " + err.toString());
    return { 
      success: false, 
      message: err.toString() 
    };
  }
}