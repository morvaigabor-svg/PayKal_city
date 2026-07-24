/**
 * PayKal Dashboard Adatkezelő
 * Lekéri az N2, N4, N6 cellákból az aktuális egyenlegeket,
 * és az A-H oszlopokból felépíti az idősoros vonaldiagramot.
 */
function getPayKalDashboardDataImpl(timeFilter) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var balanceSheet = ss.getSheetByName('Egyenleg');

  if (!balanceSheet) {
    return { 
      totalBalance: 0, 
      cashBalance: 0, 
      bankBalance: 0, 
      labels: [], 
      values: [] 
    };
  }

  // 1. AKTUÁLIS EGYENLEGEK KÖZVETLEN LEKÉRÉSE AZ N CELLÁKBÓL
  var totalBalance = Number(balanceSheet.getRange('N2').getValue()) || 0;
  var cashBalance  = Number(balanceSheet.getRange('N4').getValue()) || 0;
  var bankBalance  = Number(balanceSheet.getRange('N6').getValue()) || 0;

  // 2. IDŐSOROS DIAGRAM ADATOK LEKÉRÉSE (A: Dátum, H: Egyenleg)
  var lastRow = balanceSheet.getLastRow();
  if (lastRow < 2) {
    return {
      totalBalance: totalBalance,
      cashBalance: cashBalance,
      bankBalance: bankBalance,
      labels: [],
      values: []
    };
  }

  var rawData = balanceSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  var now = new Date();
  var startDate = new Date();

  switch (timeFilter) {
    case '1H': 
      startDate.setMonth(now.getMonth() - 1); 
      break;
    case '3H': 
      startDate.setMonth(now.getMonth() - 3); 
      break;
    case '6H': 
      startDate.setMonth(now.getMonth() - 6); 
      break;
    case '1E': 
      startDate.setFullYear(now.getFullYear() - 1); 
      break;
    case '2E': 
      startDate.setFullYear(now.getFullYear() - 2); 
      break;
    case 'O':  
      startDate = new Date(2000, 0, 1); 
      break;
    default:   
      startDate.setMonth(now.getMonth() - 1); 
      break;
  }

  var seenDates = {};
  var filteredPoints = [];

  for (var i = 0; i < rawData.length; i++) {
    var rawDate = rawData[i][0];
    var amount  = Number(rawData[i][7]);

    if (!rawDate || !(rawDate instanceof Date) || isNaN(rawDate.getTime())) {
      continue;
    }

    if (rawDate >= startDate && rawDate <= now) {
      var dateKey = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

      if (!seenDates[dateKey]) {
        seenDates[dateKey] = true;
        var displayLabel = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MM.dd.");

        filteredPoints.push({
          label: displayLabel,
          value: isNaN(amount) ? 0 : amount
        });
      }
    }
  }

  filteredPoints.reverse();

  var labels = [];
  var values = [];

  for (var j = 0; j < filteredPoints.length; j++) {
    labels.push(filteredPoints[j].label);
    values.push(filteredPoints[j].value);
  }

  return {
    totalBalance: totalBalance,
    cashBalance: cashBalance,
    bankBalance: bankBalance,
    labels: labels,
    values: values
  };
}