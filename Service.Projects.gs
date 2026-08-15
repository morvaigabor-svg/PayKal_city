/**
 * Service.Projects.gs
 * Dinamikus projektek és beállítások kezelése
 */

function getSettingsData() {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const csoportId = String(auth.csoportId || "").trim();
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  let costCentersBase = [];
  let paymentMethods = [];
  let incomePurposesBase = [];
  let incomePaymentMethods = [];
  let settingF = [];

  // 1. Beállítások munkalap beolvasása (A-F oszlopok)
  const settingsSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.SETTINGS) ? APP.SHEETS.SETTINGS : "Beállítások";
  const settingsSheet = ss.getSheetByName(settingsSheetName);

  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    const lastRow = settingsSheet.getLastRow();
    const data = settingsSheet.getRange(2, 1, lastRow - 1, 6).getValues();

    data.forEach(row => {
      if (row[0]) costCentersBase.push(String(row[0]).trim());       // A2:A
      if (row[1]) paymentMethods.push(String(row[1]).trim());        // B2:B
      if (row[2]) incomePurposesBase.push(String(row[2]).trim());    // C2:C
      if (row[3]) incomePaymentMethods.push(String(row[3]).trim());  // D2:D
      if (row[5]) settingF.push(String(row[5]).trim());              // F2:F
    });
  }

  // 2. KM Tagok beolvasása (A oszlop: Név | B oszlop: Csoport ID)
  let payers = [];
  const membersSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.MEMBERS) ? APP.SHEETS.MEMBERS : "Tagok";
  const membersSheet = ss.getSheetByName(membersSheetName);

  if (membersSheet && membersSheet.getLastRow() >= 2) {
    const mData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();

    mData.forEach(row => {
      const memberName = row[0] ? String(row[0]).trim() : "";
      const groupVal = row[1] ? String(row[1]).trim() : "";

      if (memberName !== "" && groupVal === csoportId) {
        payers.push(memberName);
      }
    });
  }
  
// Koordinátorok esetén a "KM Egyesület" automatikus hozzáadása a lista elejére
  if (auth.isCoordinator) {
    payers.unshift("KM Egyesület");
  }
  
  // 3. Aktív projektek kinyerése
  const activeProjects = getActiveProjectsData(csoportId, ss);
  const activeProjectNames = activeProjects.map(p => p.name);

  // Kiadások összekészítése
  const costCenters = [
    ...costCentersBase.filter(item => item.toLowerCase() !== "egyéb" && item.toLowerCase() !== "egyéb kiadások"),
    ...activeProjectNames,
    "Egyéb kiadások"
  ];

  // Bevételek összekészítése
  const incomePurposes = [
    ...incomePurposesBase.filter(item => item.toLowerCase() !== "egyéb" && item.toLowerCase() !== "egyéb bevételek"),
    ...activeProjectNames,
    "Egyéb bevételek"
  ];

  return {
    costCenters: [...new Set(costCenters)],
    paymentMethods: [...new Set(paymentMethods)],
    incomePurposes: [...new Set(incomePurposes)],
    incomePaymentMethods: [...new Set(incomePaymentMethods)],
    payers: [...new Set(payers)],
    settingF: [...new Set(settingF)],
    projectTypes: getProjectTypes(ss),
    activeProjects: activeProjects
  };
}

function getActiveProjectsData(csoportId, ssInstance) {
  const ss = ssInstance || SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName("Projektek");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const projects = [];

  for (let i = 0; i < data.length; i++) {
    const rowCsoport = String(data[i][0]).trim();
    const projektId = String(data[i][1]).trim();
    const projektNeve = String(data[i][2]).trim();
    const tipus = String(data[i][3]).trim();
    const aktiv = data[i][4];

    if (rowCsoport === String(csoportId).trim() && (aktiv === true || String(aktiv).toUpperCase() === 'TRUE')) {
      if (projektNeve) {
        projects.push({
          id: projektId,
          name: projektNeve,
          type: tipus
        });
      }
    }
  }
  return projects;
}

function getProjectTypes(ssInstance) {
  const auth = getUserAuth();
  if (!auth.authorized) {
    throw new Error("Jogosulatlan hozzáférés!");
  }

  const ss = ssInstance || SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const settingsSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.SETTINGS) ? APP.SHEETS.SETTINGS : "Beállítások";
  const settingsSheet = ss.getSheetByName(settingsSheetName);

  if (!settingsSheet || settingsSheet.getLastRow() < 2) {
    return [];
  }

  // Oszlop meghatározása szerepkör alapján:
  // Koordinátor -> G oszlop (7. oszlop)
  // Csoportvezető -> C oszlop (3. oszlop)
  const targetColumn = auth.isCoordinator ? 7 : 3;

  const values = settingsSheet.getRange(2, targetColumn, settingsSheet.getLastRow() - 1, 1).getValues();

  return values
    .map(row => String(row[0] || "").trim())
    .filter(val => val !== "");
}

function getActiveProjects() {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");
  return getActiveProjectsData(auth.csoportId);
}

function addProject(name, type, targetCsoportId) {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const projName = String(name || "").trim();
  const projType = String(type || "").trim();

  if (!projName || !projType) {
    throw new Error("A projekt neve és típusa is kötelező!");
  }

  // Csoport azonosító feloldása (ha a koordinátor átadja, vagy fallback a saját csoportjára)
  const csoportId = String(targetCsoportId || auth.csoportId || "").trim();
  if (!csoportId) {
    throw new Error("Nem azonosítható a célcsoport azonosítója!");
  }

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName("Projektek");
  if (!sheet) {
    sheet = ss.insertSheet("Projektek");
    sheet.appendRow(["Csoport_ID", "Projekt_ID", "Projekt_Neve", "Projekt_Típusa", "Aktív"]);
  }

  const timeZone = (typeof APP !== 'undefined' && APP.TIMEZONE) ? APP.TIMEZONE : Session.getScriptTimeZone();
  const timeStamp = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");
  const projectId = "PRJ-" + timeStamp + "-" + Math.floor(1000 + Math.random() * 9000);

  // Új projekt hozzáfűzése a Projektek munkalaphoz (A-E oszlopok)
  sheet.appendRow([
    csoportId,
    projectId,
    projName,
    projType,
    true
  ]);

  return { 
    success: true, 
    projectId: projectId, 
    csoportId: csoportId 
  };
}

function archiveProject(projectId) {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName("Projektek");
  if (!sheet || sheet.getLastRow() < 2) return { success: false };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(auth.csoportId).trim() && String(data[i][1]).trim() === String(projectId).trim()) {
      sheet.getRange(i + 2, 5).setValue(false);
      break;
    }
  }
  return { success: true };
}