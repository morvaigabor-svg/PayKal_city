/**
 * Service.Dashboard.gs
 * PayKal Dashboard Adatkezelő - Javított, hibatűrő backend verzió
 */

function getPayKalDashboardDataImpl(timeFilter, selectedProject, targetCsoportId, viewType) {
  try {
    const auth = getUserAuth();
    if (!auth || !auth.authorized) {
      throw new Error("Jogosulatlan hozzáférés!");
    }

    let activeTarget = targetCsoportId || auth.csoportId;
    let activeViewType = viewType || "LEADERSHIP_GROUP";

    if (!auth.isCoordinator) {
      activeTarget = auth.csoportId;
      activeViewType = "STANDARD";
    }

    const currentProject = selectedProject ? String(selectedProject).trim() : "ALL";
    const sheetId = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_ID) ? CONFIG.SHEET_ID : SpreadsheetApp.getActiveSpreadsheet().getId();
    const ss = SpreadsheetApp.openById(sheetId);

    let cityGroupIds = [];
    const userCity = auth.userCity || "";
    
    if (userCity) {
      const groupsSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.GROUPS) ? APP.SHEETS.GROUPS : "Csoportok";
      const groupsSheet = ss.getSheetByName(groupsSheetName);
      if (groupsSheet && groupsSheet.getLastRow() >= 2) {
        const groupsData = groupsSheet.getRange(2, 1, groupsSheet.getLastRow() - 1, groupsSheet.getLastColumn()).getValues();
        for (let i = 0; i < groupsData.length; i++) {
          const gId = String(groupsData[i][0] || "").trim();
          const gCity = String(groupsData[i][1] || "").trim();
          if (gCity.toLowerCase() === userCity.toLowerCase()) {
            cityGroupIds.push(gId);
          }
        }
      }
    }

    // 1. NÉZET: VÁROSI ÖSSZESÍTŐ (SÁVDIAGRAM)
    if (activeViewType === "CITY_SUMMARY") {
      let cityTotalCash = 0;
      let cityTotalBank = 0;
      let cityTotal = 0;

      const groupBalances = {};
      cityGroupIds.forEach(id => { groupBalances[id] = { cash: 0, bank: 0, total: 0 }; });

      const balanceSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.BALANCE) ? APP.SHEETS.BALANCE : "Egyenleg";
      const balanceSheet = ss.getSheetByName(balanceSheetName);

      if (balanceSheet && balanceSheet.getLastRow() >= 2) {
        const balanceData = balanceSheet.getRange(2, 1, balanceSheet.getLastRow() - 1, balanceSheet.getLastColumn()).getValues();
        for (let i = 0; i < balanceData.length; i++) {
          const rowGId = String(balanceData[i][0] || "").trim();
          if (cityGroupIds.includes(rowGId)) {
            const c = parseAmount(balanceData[i][1]);
            const b = parseAmount(balanceData[i][2]);
            const t = parseAmount(balanceData[i][3]) || (c + b);

            groupBalances[rowGId] = { cash: c, bank: b, total: t };
            cityTotalCash += c;
            cityTotalBank += b;
            cityTotal += t;
          }
        }
      }

      const barLabels = [];
      const barCashData = [];
      const barBankData = [];

      Object.keys(groupBalances).forEach(gId => {
        barLabels.push(gId);
        barCashData.push(groupBalances[gId].cash);
        barBankData.push(groupBalances[gId].bank);
      });

      return {
        viewType: "CITY_SUMMARY",
        csoportId: activeTarget || "",
        selectedProject: "ALL",
        displayTitle: `${userCity.toUpperCase()} - ÖSSZVAGYON`,
        totalBalance: cityTotal,
        cashBalance: cityTotalCash,
        bankBalance: cityTotalBank,
        chartType: "BAR",
        barData: {
          labels: barLabels,
          cash: barCashData,
          bank: barBankData
        },
        projects: [],
        labels: [],
        values: [],
        readOnly: true
      };
    }

    // 2. ÉS 3. NÉZET: EGYEDI CSOPORTNÉZET (VONALDIAGRAM)
    let totalBalance = 0;
    let cashBalance = 0;
    let bankBalance = 0;

    if (currentProject === "ALL") {
      const balanceSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.BALANCE) ? APP.SHEETS.BALANCE : "Egyenleg";
      const balanceSheet = ss.getSheetByName(balanceSheetName);
      
      if (balanceSheet && balanceSheet.getLastRow() >= 2) {
        const balanceData = balanceSheet.getRange(2, 1, balanceSheet.getLastRow() - 1, balanceSheet.getLastColumn()).getValues();
        for (let i = 0; i < balanceData.length; i++) {
          if (String(balanceData[i][0] || "").trim().toLowerCase() === activeTarget.toLowerCase()) {
            cashBalance = parseAmount(balanceData[i][1]);
            bankBalance = parseAmount(balanceData[i][2]);
            totalBalance = parseAmount(balanceData[i][3]) || (cashBalance + bankBalance);
            break;
          }
        }
      }
    }

    const calcSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.CALC) ? APP.SHEETS.CALC : "Számolótábla";
    const calcSheet = ss.getSheetByName(calcSheetName);

    const dailyDeltas = {};
    let totalNetDelta = 0;
    let projCashDelta = 0;
    let projBankDelta = 0;
    let projTotalDelta = 0;

    if (calcSheet && calcSheet.getLastRow() >= 2 && calcSheet.getLastColumn() >= 1) {
      const calcData = calcSheet.getRange(2, 1, calcSheet.getLastRow() - 1, calcSheet.getLastColumn()).getValues();
      
      for (let i = 0; i < calcData.length; i++) {
        const row = calcData[i];
        const rowCsoport = String(row[0] || "").trim();
        const rawDate = row[2];
        const rowProjId = String(row[3] || "").trim();
        const rowProjName = String(row[4] || "").trim();

        // Biztonságos tömbindexelés az M (12), N (13) és O (14) oszlopokra
        const netDelta = row.length > 12 ? parseAmount(row[12]) : 0;
        const cashDelta = row.length > 13 ? parseAmount(row[13]) : 0;
        const bankDelta = row.length > 14 ? parseAmount(row[14]) : 0;

        if (rowCsoport.toLowerCase() === activeTarget.toLowerCase()) {
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

    if (currentProject !== "ALL") {
      cashBalance = projCashDelta;
      bankBalance = projBankDelta;
      totalBalance = projTotalDelta;
    }

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
    startDate.setHours(0, 0, 0, 0);
    const startDateStr = formatDateKey(startDate);

    const sortedDates = Object.keys(dailyDeltas).sort();
    const initialBalance = totalBalance - totalNetDelta;

    let runningBalance = initialBalance;
    const filteredPoints = [];

    sortedDates.forEach(dateStr => {
      const delta = dailyDeltas[dateStr];
      runningBalance += delta;

      if (dateStr >= startDateStr) {
        filteredPoints.push({
          dateStr: dateStr,
          balance: runningBalance
        });
      }
    });

    const labels = [];
    const values = [];

    if (filteredPoints.length > 0) {
      filteredPoints.forEach(pt => {
        const parts = pt.dateStr.split('-');
        const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        labels.push(formatChartLabel(d));
        values.push(pt.balance);
      });
    } else {
      labels.push(formatChartLabel(startDate));
      values.push(totalBalance);
      
      labels.push(formatChartLabel(now));
      values.push(totalBalance);
    }

    const isOwnGroup = (activeTarget.toLowerCase() === (auth.csoportId || "").toLowerCase());

    let groupProjects = [];
    try {
      if (typeof getActiveProjectsData === "function") {
        groupProjects = getActiveProjectsData(activeTarget, ss);
      }
    } catch (e) {
      groupProjects = [];
    }

    return {
      viewType: activeViewType,
      csoportId: activeTarget,
      displayTitle: isOwnGroup ? `${activeTarget.toUpperCase()} CSOPORT VAGYONA` : `${activeTarget.toUpperCase()} (CSOPORTNÉZET)`,
      selectedProject: currentProject,
      totalBalance: totalBalance,
      cashBalance: cashBalance,
      bankBalance: bankBalance,
      chartType: "LINE",
      labels: labels,
      values: values,
      projects: groupProjects || [],
      readOnly: !isOwnGroup
    };

  } catch (err) {
    Logger.log("Hiba a getPayKalDashboardDataImpl rendszermagban: " + err.toString());
    throw new Error("Műszerfal adatok betöltési hibája: " + err.message);
  }
}

function formatChartLabel(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  return `${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}.`;
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
  let str = String(val).trim().replace(/[\s\u200B-\u200D\uFEFF]/g, '').replace(',', '.').replace(/[−–]/g, '-').replace(/[^0-9.-]/g, '');
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}