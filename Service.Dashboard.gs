/**
 * Service.Dashboard.gs
 * PayKal Dashboard Adatkezelő
 * Csoportra szabott aktuális egyenlegek és idősoros vagyon-alakulás lekérése.
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
        cashBalance = Number(balanceData[i][1]) || 0;
        bankBalance = Number(balanceData[i][2]) || 0;
        totalBalance = Number(balanceData[i][3]) || (cashBalance + bankBalance);
        break;
      }
    }
  }

  // 2. IDŐSOROS DIAGRAM ADATOK ELŐÁLLÍTÁSA (BEVÉTELEK ÉS KÖLTSÉGEK ALAPJÁN)
  const incomeSheet = ss.getSheetByName(APP.SHEETS.INCOME);
  const expenseSheet = ss.getSheetByName(APP.SHEETS.EXPENSES);

  // Napra bontott egyenlegváltozások gyűjtése
  const dailyDeltas = {};

  // Bevételek feldolgozása (+ delta)
  if (incomeSheet && incomeSheet.getLastRow() >= 2) {
    const incomeData = incomeSheet.getRange(2, 1, incomeSheet.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < incomeData.length; i++) {
      if (String(incomeData[i][0]).trim() === csoportId) {
        const rawDate = incomeData[i][1];
        const amount = Number(incomeData[i][3]) || 0;
        if (rawDate && amount > 0) {
          const dateKey = formatDateKey(rawDate);
          if (dateKey) {
            dailyDeltas[dateKey] = (dailyDeltas[dateKey] || 0) + amount;
          }
        }
      }
    }
  }

  // Költségek feldolgozása (- delta)
  if (expenseSheet && expenseSheet.getLastRow() >= 2) {
    const expenseData = expenseSheet.getRange(2, 1, expenseSheet.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < expenseData.length; i++) {
      if (String(expenseData[i][0]).trim() === csoportId) {
        const rawDate = expenseData[i][1];
        const amount = Number(expenseData[i][3]) || 0;
        if (rawDate && amount > 0) {
          const dateKey = formatDateKey(rawDate);
          if (dateKey) {
            dailyDeltas[dateKey] = (dailyDeltas[dateKey] || 0) - amount;
          }
        }
      }
    }
  }

  // Dátumok sorbarendezése emelkedő időrendben
  const sortedDates = Object.keys(dailyDeltas).sort();

  // Göngyölt (kumulált) összegyenleg kiszámítása napról napra
  let runningBalance = 0;
  const timeSeries = sortedDates.map(dateStr => {
    runningBalance += dailyDeltas[dateStr];
    return {
      dateStr: dateStr,
      dateObj: new Date(dateStr),
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

  const filteredPoints = timeSeries.filter(pt => pt.dateObj >= startDate && pt.dateObj <= now);

  const labels = [];
  const values = [];

  filteredPoints.forEach(pt => {
    const month = String(pt.dateObj.getMonth() + 1).padStart(2, '0');
    const day = String(pt.dateObj.getDate()).padStart(2, '0');
    labels.push(`${month}.${day}.`);
    values.push(pt.balance);
  });

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
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    return Utilities.formatDate(rawDate, APP.TIMEZONE, "yyyy-MM-dd");
  } else if (typeof rawDate === 'string' && rawDate.trim() !== '') {
    const d = new Date(rawDate);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, APP.TIMEZONE, "yyyy-MM-dd");
    }
  }
  return null;
}