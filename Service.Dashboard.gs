/**
 * Service.Dashboard.gs
 * PayKal Dashboard Adatkezelő - Számolótábla alapú hibatűrő feldolgozás
 */

function getPayKalDashboardDataImpl(timeFilter, selectedProject) {
  const auth = getUserAuth();
  if (!auth.authorized) {
    throw new Error("Jogosulatlan hozzáférés!");
  }
  const csoportId = String(auth.csoportId || "").trim();
  const currentProject = selectedProject ? String(selectedProject).trim() : "ALL";

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  
  let totalBalance = 0;
  let cashBalance = 0;
  let bankBalance = 0;

  // 1. ÖSSZVAGYON LEKÉRÉSE (ha "ALL" van kijelölve)
  if (currentProject === "ALL") {
    const balanceSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.BALANCE) ? APP.SHEETS.BALANCE : "Egyenleg";
    const balanceSheet = ss.getSheetByName(balanceSheetName);
    
    if (balanceSheet && balanceSheet.getLastRow() >= 2) {
      const balanceData = balanceSheet.getRange(2, 1, balanceSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < balanceData.length; i++) {
        if (String(balanceData[i][0]).trim() === csoportId) {
          cashBalance = parseAmount(balanceData[i][1]);
          bankBalance = parseAmount(balanceData[i][2]);
          totalBalance = parseAmount(balanceData[i][3]) || (cashBalance + bankBalance);
          break;
        }
      }
    }
  }

  // 2. SZÁMOLÓTÁBLA FELDOLGOZÁSA
  const calcSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.CALC) ? APP.SHEETS.CALC : "Számolótábla";
  const calcSheet = ss.getSheetByName(calcSheetName);

  const dailyDeltas = {};
  let totalNetDelta = 0;

  // Projekt szűrőhöz tartozó változók
  let projCashDelta = 0;
  let projBankDelta = 0;
  let projTotalDelta = 0;

  // Csoport szűrőhöz tartozó globális változók (tartalék főegyenleghez)
  let allGroupCashDelta = 0;
  let allGroupBankDelta = 0;
  let allGroupTotalDelta = 0;

  if (calcSheet && calcSheet.getLastRow() >= 2) {
    // A-K oszlopok beolvasása (11 oszlop)
    const calcData = calcSheet.getRange(2, 1, calcSheet.getLastRow() - 1, 11).getValues(); 
    
    for (let i = 0; i < calcData.length; i++) {
      const rowCsoport = String(calcData[i][0] || "").trim();   // A: Csoport_ID
      const rawDate = calcData[i][2];                            // C: Dátum
      const rowProjId = String(calcData[i][3] || "").trim();     // D: Projekt ID
      const rowProjName = String(calcData[i][4] || "").trim();   // E: Tranzakció célja / Projekt neve
      
      const netDelta = parseAmount(calcData[i][8]);              // I: Egyenleg változás (Ft)
      const cashDelta = parseAmount(calcData[i][9]);             // J: Készpénz változás
      const bankDelta = parseAmount(calcData[i][10]);            // K: Számla változás

      if (rowCsoport === csoportId) {
        // Globális csoportösszegek gyűjtése
        allGroupCashDelta += cashDelta;
        allGroupBankDelta += bankDelta;
        allGroupTotalDelta += netDelta;

        // Szűrés ellenőrzése: Egyezik ID-ra, Névre vagy "ALL"-ra
        const matchesProject = (currentProject === "ALL") || 
                               (rowProjId.toLowerCase() === currentProject.toLowerCase()) || 
                               (rowProjName.toLowerCase() === currentProject.toLowerCase());

        if (matchesProject) {
          projCashDelta += cashDelta;
          projBankDelta += bankDelta;
          projTotalDelta += netDelta;

          if (rawDate) {
            const dateKey = formatDateKey(rawDate);
            if (dateKey) {
              dailyDeltas[dateKey] = (dailyDeltas[dateKey] || 0) + netDelta;
              totalNetDelta += netDelta;
            }
          }
        }
      }
    }
  }

  // Egyenlegek beállítása
  if (currentProject !== "ALL") {
    cashBalance = projCashDelta;
    bankBalance = projBankDelta;
    totalBalance = projTotalDelta;
  } else if (totalBalance === 0 && (allGroupTotalDelta !== 0 || allGroupCashDelta !== 0)) {
    // Ha az Egyenleg munkalapról nem jött adat, a Számolótábla összesítését használja főegyenlegként
    cashBalance = allGroupCashDelta;
    bankBalance = allGroupBankDelta;
    totalBalance = allGroupTotalDelta;
  }

  // 3. IDŐSOROS DIAGRAM ADATOK ELŐÁLLÍTÁSA
  const sortedDates = Object.keys(dailyDeltas).sort();
  const initialBalance = totalBalance - totalNetDelta;

  let runningBalance = initialBalance;
  const timeSeries = sortedDates.map(dateStr => {
    runningBalance += dailyDeltas[dateStr];
    
    const parts = dateStr.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    const dObj = new Date(year, month, day);

    return {
      dateStr: dateStr,
      dateObj: dObj,
      balance: runningBalance
    };
  });

  const now = new Date();
  let startDate = new Date();

  switch (timeFilter) {
    case '1H': startDate.setMonth(now.getMonth() - 1); break;
    case '3H': startDate.setMonth(now.getMonth() - 3); break;
    case '6H': startDate.setMonth(now.getMonth() - 6); break;
    case '1E': startDate.setFullYear(now.getFullYear() - 1); break;
    case '2E': startDate.setFullYear(now.getFullYear() - 2); break;
    case 'O':  startDate = new Date(2000, 0, 1); break;
    default:   startDate.setMonth(now.getMonth() - 1); break;
  }

  const filteredPoints = timeSeries.filter(pt => pt.dateObj >= startDate && pt.dateObj <= now);

  const labels = [];
  const values = [];

  if (filteredPoints.length > 0) {
    filteredPoints.forEach(pt => {
      const month = String(pt.dateObj.getMonth() + 1).padStart(2, '0');
      const day = String(pt.dateObj.getDate()).padStart(2, '0');
      labels.push(`${month}.${day}.`);
      values.push(pt.balance);
    });
  } else {
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    labels.push(`${month}.${day}.`);
    values.push(totalBalance);
  }

  return {
    csoportId: csoportId,
    selectedProject: currentProject,
    totalBalance: totalBalance,
    cashBalance: cashBalance,
    bankBalance: bankBalance,
    labels: labels,
    values: values
  };
}

function formatDateKey(rawDate) {
  if (!rawDate) return null;
  const tz = (typeof APP !== 'undefined' && APP.TIMEZONE) ? APP.TIMEZONE : Session.getScriptTimeZone();

  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    return Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
  }
  
  if (typeof rawDate === 'string' && rawDate.trim() !== '') {
    let clean = rawDate.trim().replace(/\./g, '-').replace(/\s+/g, '');
    if (clean.endsWith('-')) clean = clean.slice(0, -1);
    const d = new Date(clean);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, "yyyy-MM-dd");
    }
  }
  return null;
}

function parseAmount(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  
  let str = String(val).trim().replace(/[\s\u200B-\u200D\uFEFF]/g, '');
  str = str.replace(',', '.');
  str = str.replace(/[−–]/g, '-');
  str = str.replace(/[^0-9.-]/g, '');
  
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}