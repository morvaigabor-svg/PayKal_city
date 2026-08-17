/**
 * @file Config.gs
 * @description Az alkalmazás globális konfigurációs beállításai.
 */

const APP = Object.freeze({
  NAME: "PayKal - Wallet of Calasanz",
  VERSION: "0.2.0",
  TIMEZONE: "Europe/Budapest",
  SHEETS: Object.freeze({
    USERS: "Felhasználók",
    GROUPS: "Csoportok",
    MEMBERS: "KM Tagok", // Frissítve a pontos munkalap névre!
    CITIES: "Városok",
    SETTINGS: "Beállítások",
    BALANCE: "Egyenleg",
    EXPENSES: "Költségek",
    INCOME: "Bevételek",
    TRANSFERS: "Pénzmozgások",
    SUMMARY: "Összesítő",
    LOG: "Log"
  })
});

const CONFIG = {
  SHEET_ID: "1vjjkOriMEJVKBPzY_k3NNJItmLQySWZTQCCZHiovDYE"
};

/* Egyetlen központi menü az összes adminisztrációs funkcióhoz*/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PayKal Admin')
    .addItem('➕ Új csoport létrehozása', 'openNewGroupModal')
    .addSeparator()
    .addItem('📊 Kimutatások generálása', 'openReportGeneratorDialog')
    .addSeparator()
    .addItem('Blokkok frissítése', 'frissitsBlokkLinkeket')
    .addToUi();
}

function frissitsBlokkLinkeket() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var forrasLap = ss.getSheetByName('Költségek');
  var celLap = ss.getSheetByName('Számolótábla') || ss.getSheetByName('Számolólap');
  
  if (!forrasLap || !celLap) return;

  // 1. Csoportok mappáinak feltérképezése a Drive kereséshez
  var groupsMap = {};
  var groupsSheet = ss.getSheetByName('Csoportok');
  if (groupsSheet && groupsSheet.getLastRow() >= 2) {
    var gData = groupsSheet.getRange(2, 1, groupsSheet.getLastRow() - 1, 3).getValues();
    for (var g = 0; g < gData.length; g++) {
      var gId = String(gData[g][0] || "").trim();
      var fId = String(gData[g][2] || "").trim();
      if (gId && fId) groupsMap[gId] = fId;
    }
  }

  // 2. Költségek lap beolvasása (A: Csoport_ID [1], F: Blokk [6], K: Kiadás_ID [11])
  var forrasUtolsoSor = forrasLap.getLastRow();
  var szotar = {};
  var uresRichText = SpreadsheetApp.newRichTextValue().setText("").build();
  
  if (forrasUtolsoSor >= 2) {
    var forrasCsoportok = forrasLap.getRange(2, 1, forrasUtolsoSor - 1, 1).getValues();
    var forrasRichTextek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getRichTextValues();
    var forrasKepletek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getFormulas();
    var forrasNyersErtekek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getValues();
    var forrasIdek = forrasLap.getRange(2, 11, forrasUtolsoSor - 1, 1).getValues();

    for (var i = 0; i < forrasIdek.length; i++) {
      var id = String(forrasIdek[i][0] || "").trim();
      if (!id) continue;

      var csoportId = String(forrasCsoportok[i][0] || "").trim();
      var rt = forrasRichTextek[i][0];
      var keplet = String(forrasKepletek[i][0] || "").trim();
      var nyers = String(forrasNyersErtekek[i][0] || "").trim();
      var szoveg = (rt && rt.getText()) ? String(rt.getText()).trim() : nyers;

      // --- A) TÖBB-BLOKKOS ESET (pl. "1 | 2" vagy "1 | 2 | 3") ---
      if (szoveg.indexOf("|") !== -1) {
        if (rt) {
          szotar[id] = rt;
          continue;
        }
      }

      // --- B) URL KERESÉSE A CELLÁBÓL ---
      var talaltUrl = null;

      // 1. Képletből
      if (keplet) {
        var m = keplet.match(/https?:\/\/[^\s"',;)]+/i);
        if (m) talaltUrl = m[0];
      }

      // 2. RichText-ből
      if (!talaltUrl && rt) {
        if (rt.getLinkUrl()) {
          talaltUrl = rt.getLinkUrl();
        } else {
          var runs = rt.getRuns();
          for (var r = 0; r < runs.length; r++) {
            if (runs[r].getLinkUrl()) {
              talaltUrl = runs[r].getLinkUrl();
              break;
            }
          }
        }
      }

      // 3. Nyers szövegből ha URL
      if (!talaltUrl && nyers.match(/^https?:\/\//i)) {
        talaltUrl = nyers;
      }

      // --- C) DRIVE FALLBACK: Ha még mindig nincs link, megkeressük a Drive mappában! ---
      if (!talaltUrl && csoportId && groupsMap[csoportId]) {
        try {
          var folderId = groupsMap[csoportId];
          var folder = DriveApp.getFolderById(folderId);
          var files = folder.searchFiles("title contains '" + id + "' and trashed = false");
          
          var driveLinks = [];
          var fileIdx = 1;
          while (files.hasNext()) {
            var f = files.next();
            driveLinks.push({ label: String(fileIdx), url: f.getUrl() });
            fileIdx++;
          }

          if (driveLinks.length > 0) {
            var fullTxt = driveLinks.map(function(d) { return d.label; }).join(" | ");
            var dBuilder = SpreadsheetApp.newRichTextValue().setText(fullTxt);
            var offset = 0;
            driveLinks.forEach(function(d) {
              dBuilder.setLinkUrl(offset, offset + d.label.length, d.url);
              offset += d.label.length + 3;
            });
            szotar[id] = dBuilder.build();
            continue;
          }
        } catch(e) {
          // Ha nincs Drive jogosultság a mappához, megyünk tovább
        }
      }

      // --- D) VÉGLEGES ÖSSZEÁLLÍTÁS AZ 1 DB BLOKKHOZ ---
      if (talaltUrl) {
        var b = SpreadsheetApp.newRichTextValue().setText(szoveg || "1");
        b.setLinkUrl(0, (szoveg || "1").length, talaltUrl);
        szotar[id] = b.build();
      } else if (szoveg) {
        szotar[id] = SpreadsheetApp.newRichTextValue().setText(szoveg).build();
      } else {
        szotar[id] = uresRichText;
      }
    }
  }

  // 3. Számolótábla feldolgozása (B: ID, L: Típus)
  var celUtolsoSor = celLap.getLastRow();
  if (celUtolsoSor < 2) return;

  var celIdek = celLap.getRange(2, 2, celUtolsoSor - 1, 1).getValues();
  var celTipusok = celLap.getRange(2, 12, celUtolsoSor - 1, 1).getValues();

  var kimenet = [];
  var count = 0;

  for (var j = 0; j < celIdek.length; j++) {
    var tipus = String(celTipusok[j] ? celTipusok[j][0] : "").trim();
    var celId = String(celIdek[j] ? celIdek[j][0] : "").trim();

    if (tipus.toLowerCase() === "kiadás" && celId && szotar[celId]) {
      kimenet.push([szotar[celId]]);
      count++;
    } else {
      kimenet.push([uresRichText]);
    }
  }

  // 4. Beírás a Q OSZLOPBA (17. oszlop)
  celLap.getRange(2, 17, kimenet.length, 1).setRichTextValues(kimenet);

  if (SpreadsheetApp.getUi) {
    try {
      SpreadsheetApp.getUi().alert("Kész! Összesen " + count + " db sor frissült a Q oszlopban.");
    } catch(e) {}
  }
}