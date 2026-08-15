/**
 * Service.Projects.gs
 * Dinamikus projektek és fix kategóriák kezelése
 */

/**
 * Legördülő menü elemeinek lekérése az űrlapokhoz (Bevétel / Kiadás)
 * @param {string} formType - 'income' vagy 'expense'
 */
function getCategoriesForForm(formType) {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const csoportId = String(auth.csoportId || "").trim();
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // 1. FIX ELEMEK BEOLVASÁSA A 'Beállítások' LAPRÓL
  const settingsSheet = ss.getSheetByName("Beállítások");
  let fixedCategories = [];

  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    // Kiadás: A oszlop (1), Bevétel: C oszlop (3)
    const colIndex = (formType === 'income') ? 3 : 1; 
    const rawData = settingsSheet.getRange(2, colIndex, settingsSheet.getLastRow() - 1, 1).getValues();

    fixedCategories = rawData
      .map(row => String(row[0]).trim())
      .filter(val => val !== "" && val.toLowerCase() !== "egyéb");
  }

  // 2. DINAMIKUS PROJEKTEK BEOLVASÁSA A 'Projektek' LAPRÓL
  const projectSheet = ss.getSheetByName("Projektek");
  let dynamicProjects = [];

  if (projectSheet && projectSheet.getLastRow() >= 2) {
    const data = projectSheet.getRange(2, 1, projectSheet.getLastRow() - 1, 5).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowCsoport = String(data[i][0]).trim();
      const projektNeve = String(data[i][2]).trim();
      const aktiv = data[i][4];

      if (rowCsoport === csoportId && (aktiv === true || String(aktiv).toUpperCase() === 'TRUE')) {
        if (projektNeve && projektNeve.toLowerCase() !== "egyéb") {
          dynamicProjects.push(projektNeve);
        }
      }
    }
  }

  // 3. SZIGORÚ SORREND: Fix elemek -> Dinamikus projektek -> "Egyéb" (mint utolsó)
  return [...fixedCategories, ...dynamicProjects, "Egyéb"];
}

/**
 * A 'Projektek kezelése' felugró ablakhoz tartozó aktív projektek listája
 */
function getActiveProjects() {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const csoportId = String(auth.csoportId || "").trim();
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
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

    if (rowCsoport === csoportId && (aktiv === true || String(aktiv).toUpperCase() === 'TRUE')) {
      projects.push({ id: projektId, nev: projektNeve, name: projektNeve, tipus: tipus });
    }
  }

  return projects;
}

/**
 * Új projekt rögzítése
 */
function addProject(projektNeve, tipus) {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const csoportId = String(auth.csoportId || "").trim();
  if (!projektNeve || !tipus) throw new Error("Minden mező kitöltése kötelező!");

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName("Projektek");

  if (!sheet) {
    sheet = ss.insertSheet("Projektek");
    sheet.appendRow(["Csoport_ID", "Projekt_ID", "Projekt_Neve", "Tipus", "Aktiv"]);
  }

  const projektId = "PRJ-" + Utilities.getUuid().substring(0, 8);
  sheet.appendRow([csoportId, projektId, projektNeve.trim(), tipus.trim(), true]);

  return { success: true };
}

/**
 * Projekt archiválása
 */
function archiveProject(projektId) {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const csoportId = String(auth.csoportId || "").trim();
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName("Projektek");

  if (!sheet || sheet.getLastRow() < 2) throw new Error("Nincs megnevezett projekt!");

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === csoportId && String(data[i][1]).trim() === projektId) {
      sheet.getRange(i + 2, 5).setValue(false);
      return { success: true };
    }
  }

  throw new Error("A projekt nem található!");
}

/**
 * Projekt típusok lekérése a Beállítások munkalap E2:E oszlopából
 */
function getProjectTypes() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName("Beállítások");
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const values = sheet.getRange("E2:E" + lastRow).getValues();
  // Üres és duplikált elemek kiszűrése
  return values
    .map(row => row[0])
    .filter(val => val && val.toString().trim() !== "");
}