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

  // Beállítások munkalap
  const settingsSheet = ss.getSheetByName(APP.SHEETS.SETTINGS);
  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    const lastRow = settingsSheet.getLastRow();
    const data = settingsSheet.getRange(2, 1, lastRow - 1, 6).getValues();

    data.forEach(row => {
      if (row[0]) costCentersBase.push(String(row[0]).trim()); // A2:A
      if (row[1]) paymentMethods.push(String(row[1]).trim());  // B2:B
      if (row[2]) incomePurposesBase.push(String(row[2]).trim());// C2:C
      if (row[3]) incomePaymentMethods.push(String(row[3]).trim());// D2:D
      if (row[5]) settingF.push(String(row[5]).trim());        // F2:F
    });
  }

  // KM Tagok
  let payers = [];
  const membersSheet = ss.getSheetByName(APP.SHEETS.MEMBERS);
  if (membersSheet && membersSheet.getLastRow() >= 2) {
    const mData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
    mData.forEach(row => {
      if (String(row[0]).trim() === csoportId && row[1]) {
        payers.push(String(row[1]).trim());
      }
    });
  }

  // Aktív projektek kinyerése
  const activeProjects = getActiveProjectsData(csoportId, ss);
  const activeProjectNames = activeProjects.map(p => p.name);

  // 1. Kiadások: Beállítások A2:A + Aktív projektek + "Egyéb kiadások"
  const costCenters = [
    ...costCentersBase.filter(item => item.toLowerCase() !== "egyéb" && item.toLowerCase() !== "egyéb kiadások"),
    ...activeProjectNames,
    "Egyéb kiadások"
  ];

  // 2. Bevételek: Beállítások C2:C + Aktív projektek + "Egyéb bevételek"
  const incomePurposes = [
    ...incomePurposesBase.filter(item => item.toLowerCase() !== "egyéb" && item.toLowerCase() !== "egyéb bevételek"),
    ...activeProjectNames,
    "Egyéb bevételek"
  ];

  return {
    costCenters: costCenters,
    paymentMethods: paymentMethods,
    incomePurposes: incomePurposes,
    incomePaymentMethods: incomePaymentMethods,
    payers: payers,
    settingF: settingF,
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

function getProjectTypes() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const settingsSheet = ss.getSheetByName(APP.SHEETS.SETTINGS);
  if (!settingsSheet || settingsSheet.getLastRow() < 2) return [];
  
  const values = settingsSheet.getRange(2, 5, settingsSheet.getLastRow() - 1, 1).getValues(); // E oszlop
  return values.map(row => String(row[0]).trim()).filter(Boolean);
}

function getActiveProjects() {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");
  return getActiveProjectsData(auth.csoportId);
}

function addProject(name, type) {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName("Projektek");
  if (!sheet) {
    sheet = ss.insertSheet("Projektek");
    sheet.appendRow(["Csoport_ID", "Projekt_ID", "Projekt_Neve", "Projekt_Típusa", "Aktív"]);
  }

  const now = new Date();
  const timeStamp = Utilities.formatDate(now, APP.TIMEZONE, "yyyyMMdd");
  const projectId = "PRJ-" + timeStamp + "-" + Math.floor(Math.random() * 1000);

  sheet.appendRow([
    auth.csoportId,
    projectId,
    name,
    type,
    true
  ]);

  return { success: true };
}

function archiveProject(projectId) {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName("Projektek");
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(auth.csoportId).trim() && String(data[i][1]).trim() === String(projectId).trim()) {
      sheet.getRange(i + 2, 5).setValue(false);
      break;
    }
  }
  return { success: true };
}