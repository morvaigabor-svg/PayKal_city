/**
 * Service.Dashboard.gs
 * PayKal Dashboard Adatkezelő
 * Csoportra szabott aktuális egyenlegek és idősoros összegegyenleg-alakulás lekérése.
 */

function getPayKalDashboardDataImpl(timeFilter) {
  const auth = getUserAuth();
  if (!auth.authorized) {
    throw new Error("Jogosulatlan hozzáférés!");
  }
  const csoportId = String(auth.csoportId || "").trim();

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  
  // 1. AKTUÁLIS EGYENLEGEK LEKÉRÉSE AZ 'Egyenleg' LAPRÓL CSOPORT ALAPJÁN
  let totalBalance = 0;
  let cashBalance = 0;
  let bankBalance = 0;

  const balanceSheet = ss.getSheetByName(APP.SHEETS.BALANCE);
  if (balanceSheet && balanceSheet.getLastRow() >= 2) {
    const balanceData = balanceSheet.getRange(2, 1, balanceSheet.getLastRow() - 1, 4).getValues();
    // A: Csoport_ID, B: Készpénz, C: Számla, D: Összegyenleg
    for (let i = 0; i < balanceData.length; i++) {
      if (String(balanceData[i][0]).trim() === csoportId) {
        cashBalance = parseAmount(balanceData[i][1]);
        bankBalance = parseAmount(balanceData[i][2]);
        totalBalance = parseAmount(balanceData[i][3]) || (cashBalance + bankBalance);
        break;
      }
    }
  }

  // 2. IDŐSOROS DIAGRAM ADATOK ELŐÁLLÍTÁSA (BEVÉTELEK ÉS KÖLTSÉGEK NETTÓ NAPI EREDŐJE)
  const incomeSheet = ss.getSheetByName(APP.SHEETS.INCOME);
  const expenseSheet = ss.getSheetByName(APP.SHEETS.EXPENSES);

  // Napra bontott nettó egyenlegváltozások gyűjtése (yyyy-MM-dd -> net delta)
  const dailyDeltas = {};
  let totalNetDelta = 0;

  // Bevételek feldolgozása (+ delta)
  if (incomeSheet && incomeSheet.getLastRow() >= 2) {
    const incomeData = incomeSheet.getRange(2, 1, incomeSheet.getLastRow() - 1, incomeSheet.getLastColumn()).getValues();
    for (let i = 0; i < incomeData.length; i++) {
      if (String(incomeData[i][0]).trim() === csoportId) {
        const rawDate = incomeData[i][1];
        const amount = parseAmount(incomeData[i][3]);
        if (rawDate && amount > 0) {
          const dateKey = formatDateKey(rawDate);
          if (dateKey) {
            dailyDeltas[dateKey] = (dailyDeltas[dateKey] || 0) + amount;
            totalNetDelta += amount;
          }
        }
      }
    }
  }

  // Költségek feldolgozása (- delta)
  if (expenseSheet && expenseSheet.getLastRow() >= 2) {
    const expenseData = expenseSheet.getRange(2, 1, expenseSheet.getLastRow() - 1, expenseSheet.getLastColumn()).getValues();
    for (let i = 0; i < expenseData.length; i++) {
      if (String(expenseData[i][0]).trim() === csoportId) {
        const rawDate = expenseData[i][1];
        const amount = parseAmount(expenseData[i][3]);
        if (rawDate && amount > 0) {
          const dateKey = formatDateKey(rawDate);
          if (dateKey) {
            dailyDeltas[dateKey] = (dailyDeltas[dateKey] || 0) - amount;
            totalNetDelta -= amount;
          }
        }
      }
    }
  }

  // Dátumok sorbarendezése emelkedő időrendben
  const sortedDates = Object.keys(dailyDeltas).sort();

  // Kezdő bázis egyenleg kiszámítása (a tranzakciók előtti nyitó érték)
  const initialBalance = totalBalance - totalNetDelta;

  // Göngyölt (kumulált) összegegyenleg kiszámítása napról napra
  let runningBalance = initialBalance;
  const timeSeries = sortedDates.map(dateStr => {
    runningBalance += dailyDeltas[dateStr];
    
    // yyyy-MM-dd dátum helyi objektummá alakítása eltolódás nélkül
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

  // Időszak szűrés (timeFilter) kiszámítása
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

  // Szűrés a kiválasztott időszakra
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
    // Ha nem volt tranzakció a kért időszakban, az aktuális összegegyenleg jelenik meg
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    labels.push(`${month}.${day}.`);
    values.push(totalBalance);
  }

  return {
    totalBalance: totalBalance,
    cashBalance: cashBalance,
    bankBalance: bankBalance,
    labels: labels,
    values: values
  };
}

/**
 * Segédfüggvény a dátumok egységes YYYY-MM-DD formátumra alakításához
 */
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

/**
 * Segédfüggvény számmá alakításhoz (kezeli a formázott szöveges számokat is)
 */
function parseAmount(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (typeof val === 'string') {
    const cleaned = val.replace(/[^0-9.-]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}