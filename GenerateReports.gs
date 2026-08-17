/**
 * ============================================================================
 * PAYKAL MASTER SHEET - CSOPORT GENERÁLÓ SCRIPT (generate.report.gs)
 * ============================================================================
 */

const MASTER_SHEET_ID = "1vjjkOriMEJVKBPzY_k3NNJItmLQySWZTQCCZHiovDYE";
const MASTER_SHEET_URL = "https://docs.google.com/spreadsheets/d/" + MASTER_SHEET_ID + "/edit";

// ⚠️ IDE MÁSOLD BE A SABLON TÁBLÁZATOD ID-JÁT:
const TEMPLATE_SHEET_ID = "1k4McbTpZZctMQ4IghiuDYNtLxt1HXRK595RZ3cpMWmg"; 


 /* HTML Párbeszédablak megnyitása
 */
function openReportGeneratorDialog() {
  const html = HtmlService.createTemplateFromFile('ReportDialog')
    .evaluate()
    .setWidth(1000)
    .setHeight(650)
    .setTitle('Új Csoport Kimutatás Létrehozása');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Új Csoport Kimutatás Létrehozása');
}

/**
 * OAuth Token biztosítása a Google Picker API-nak
 */
function getOAuthToken() {
  DriveApp.getRootFolder(); // Kikényszeríti a Drive jogosultságot
  return ScriptApp.getOAuthToken();
}

/**
 * Aktív csoportok beolvasása a "Csoportok" munkalapról (A2:A, ha E oszlop == "Aktív")
 */
function getActiveGroupIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Csoportok');
  const aktivCsoportok = [];

  if (!sheet) return aktivCsoportok;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return aktivCsoportok;

  // A oszlop: Csoport ID (index 0), E oszlop: Státusz (index 4)
  const range = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  range.forEach(row => {
    const groupId = String(row[0] || "").trim();
    const statusz = String(row[4] || "").trim();

    if (groupId && statusz.toLowerCase() === 'aktív') {
      aktivCsoportok.push(groupId);
    }
  });

  return aktivCsoportok;
}

/**
 * A HTML Modal-ból meghívható generáló függvény (Pontosvesszős elválasztással magyar Sheet-hez).
 */
function createGroupReportSheet(payload) {
  try {
    const groupId = payload.groupId;
    const fileName = payload.fileName;
    const targetFolderId = payload.targetFolderId;

    if (!groupId || !fileName) {
      throw new Error("A Csoport ID és a Fájlnév megadása kötelező!");
    }

    if (!TEMPLATE_SHEET_ID || TEMPLATE_SHEET_ID.includes("IDE_JÖJJÖN")) {
      throw new Error("A TEMPLATE_SHEET_ID nincs beállítva a generáló scriptben!");
    }

    // 1. Sablon fájl kérése és a Célmappa kijelölése
    const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
    const targetFolder = (targetFolderId && targetFolderId !== 'root') 
      ? DriveApp.getFolderById(targetFolderId) 
      : DriveApp.getRootFolder();

    // 2. Sablon másolása a kívánt névvel és mappába
    const newFile = templateFile.makeCopy(fileName, targetFolder);
    const newSS = SpreadsheetApp.openById(newFile.getId());

    // 3. QUERY képletek beállítása pontosvesszővel (;)
    const sheetSzamolotabla = newSS.getSheetByName("Számolótábla");
    if (sheetSzamolotabla) {
      sheetSzamolotabla.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Számolótábla!A2:P"); "SELECT * WHERE Col1 = '${groupId}'"; 0)`
      );
    }

    const sheetKoltsegek = newSS.getSheetByName("Költségek");
    if (sheetKoltsegek) {
      sheetKoltsegek.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Költségek!A2:O"); "SELECT Col1, Col2, Col3, Col4, Col5, Col7, Col8, Col9, Col10, Col11, Col12, Col13, Col14, Col15 WHERE Col1 = '${groupId}'"; 0)`
      );
    }

    const sheetBevetelek = newSS.getSheetByName("Bevételek");
    if (sheetBevetelek) {
      sheetBevetelek.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Bevételek!A2:L"); "SELECT * WHERE Col1 = '${groupId}'"; 0)`
      );
    }

    const sheetPenzmozgasok = newSS.getSheetByName("Pénzmozgások");
    if (sheetPenzmozgasok) {
      sheetPenzmozgasok.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Pénzmozgások!A2:L"); "SELECT * WHERE Col1 = '${groupId}'"; 0)`
      );
    }

    const sheetProjektek = newSS.getSheetByName("Projektek");
    if (sheetProjektek) {
      sheetProjektek.getRange("A2").setFormula(
        `=QUERY(IMPORTRANGE("${MASTER_SHEET_URL}"; "Projektek!A2:F"); "SELECT * WHERE Col1 = '${groupId}'"; 0)`
      );
    }

    return {
      success: true,
      fileUrl: newSS.getUrl(),
      fileName: fileName
    };

  } catch (err) {
    Logger.log("Hiba a csoport táblázat generálása során: " + err.toString());
    return { 
      success: false, 
      message: err.toString() 
    };
  }
}