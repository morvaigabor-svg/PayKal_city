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
/**
 * FŐ TÁBLÁZAT (PayKal Admin) - Egyetlen egyeztetett onOpen()
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PayKal Admin')
    .addItem('➕ Új csoport létrehozása', 'openNewGroupModal')
    .addSeparator()
    .addItem('📊 Kimutatások generálása', 'openReportGeneratorDialog')
    .addSeparator()
    .addItem('🔄 Blokkok frissítése', 'frissitsBlokkLinkeket')
    .addToUi();
}


function frissitsBlokkLinkeket() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var celLap = ss.getSheetByName('Számolótábla') || ss.getSheetByName('Számolólap');
  var forrasLap = ss.getSheetByName('Költségek');
  
  if (!celLap || !forrasLap) {
    if (SpreadsheetApp.getUi) SpreadsheetApp.getUi().alert("Hiba: 'Számolótábla' vagy 'Költségek' munkalap nem található!");
    return;
  }

  var celUtolsoSor = celLap.getLastRow();
  if (celUtolsoSor < 2) return;

  // Beolvassuk a Számolótábla oszlopait:
  // B oszlop (2): Tranzakció ID | L oszlop (12): Típus | Q oszlop (17): Meglévő blokk érték
  var celIdek = celLap.getRange(2, 2, celUtolsoSor - 1, 1).getValues();
  var celTipusok = celLap.getRange(2, 12, celUtolsoSor - 1, 1).getValues();
  var celMeglevoRichTextek = celLap.getRange(2, 17, celUtolsoSor - 1, 1).getRichTextValues();

  // 1. LÉPÉS: Végigmegyünk a Számolótáblán, és kigyűjtjük a HIÁNYZÓ sorokat
  var hianyzoSorok = []; // Eltároljuk, melyik sorokhoz kell linket keresni
  var hianyzoIdek = {};

  for (var i = 0; i < celIdek.length; i++) {
    var tipus = String(celTipusok[i] ? celTipusok[i][0] : "").trim().toLowerCase();
    var id = String(celIdek[i] ? celIdek[i][0] : "").trim();
    var meglevoRt = celMeglevoRichTextek[i] ? celMeglevoRichTextek[i][0] : null;
    var meglevoSzoveg = meglevoRt ? String(meglevoRt.getText() || "").trim() : "";

    // HA NEM KIADÁS -> Megy tovább!
    if (tipus !== "kiadás" || !id) {
      continue;
    }

    // HA A Q OSZLOPBAN MÁR VAN ÉRTÉK -> Megy tovább!
    if (meglevoSzoveg !== "") {
      continue;
    }

    // HA KIADÁS ÉS A Q OSZLOP ÜRES -> Feljegyezzük, hogy ezt meg kell keresni!
    hianyzoSorok.push({
      sorIndex: i,       // 0-alapú index a tömbben
      sorSzam: i + 2,    // Tényleges sorszám a Sheetben
      id: id
    });
    hianyzoIdek[id] = true;
  }

  // Ha egyetlen ilyen sor sincs, azonnal készen vagyunk!
  if (hianyzoSorok.length === 0) {
    if (SpreadsheetApp.getUi) {
      try {
        SpreadsheetApp.getUi().alert("Minden 'Kiadás' típusú sor Q oszlopa már ki van töltve, nincs mit frissíteni!");
      } catch (e) {}
    }
    return;
  }

  // 2. LÉPÉS: Csak a hiányzó ID-khoz keressük meg a linkeket a 'Költségek' lapról
  var forrasUtolsoSor = forrasLap.getLastRow();
  var szotar = {};
  
  if (forrasUtolsoSor >= 2) {
    var forrasCsoportok = forrasLap.getRange(2, 1, forrasUtolsoSor - 1, 1).getValues(); // A: Csoport_ID
    var forrasRichTextek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getRichTextValues(); // F: Blokk link
    var forrasKepletek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getFormulas();
    var forrasNyersErtekek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getValues();
    var forrasIdek = forrasLap.getRange(2, 11, forrasUtolsoSor - 1, 1).getValues(); // K: Kiadás ID

    for (var k = 0; k < forrasIdek.length; k++) {
      var fId = String(forrasIdek[k][0] || "").trim();
      
      // Ha ez az ID nem hiányzik a Számolótáblában, átugorjuk
      if (!fId || !hianyzoIdek[fId]) continue;

      var rt = forrasRichTextek[k][0];
      var keplet = String(forrasKepletek[k][0] || "").trim();
      var nyers = String(forrasNyersErtekek[k][0] || "").trim();
      var szoveg = (rt && rt.getText()) ? String(rt.getText()).trim() : nyers;

      // A) Több-blokkos eset ("1 | 2")
      if (szoveg.indexOf("|") !== -1 && rt) {
        szotar[fId] = rt;
        continue;
      }

      // B) URL keresése a cellából
      var talaltUrl = null;
      if (keplet) {
        var m = keplet.match(/https?:\/\/[^\s"',;)]+/i);
        if (m) talaltUrl = m[0];
      }
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
      if (!talaltUrl && nyers.match(/^https?:\/\//i)) {
        talaltUrl = nyers;
      }

      // C) Drive Fallback keresés (ha nincs a cellában link)
      if (!talaltUrl) {
        try {
          var csoportId = String(forrasCsoportok[k][0] || "").trim();
          var groupsSheet = ss.getSheetByName('Csoportok');
          var folderId = null;
          if (groupsSheet && groupsSheet.getLastRow() >= 2) {
            var gData = groupsSheet.getRange(2, 1, groupsSheet.getLastRow() - 1, 3).getValues();
            for (var g = 0; g < gData.length; g++) {
              if (String(gData[g][0] || "").trim() === csoportId) {
                folderId = String(gData[g][2] || "").trim();
                break;
              }
            }
          }
          if (folderId) {
            var folder = DriveApp.getFolderById(folderId);
            var files = folder.searchFiles("title contains '" + fId + "' and trashed = false");
            var driveLinks = [];
            var fIdx = 1;
            while (files.hasNext()) {
              driveLinks.push({ label: String(fIdx), url: files.next().getUrl() });
              fIdx++;
            }
            if (driveLinks.length > 0) {
              var fullTxt = driveLinks.map(function(d) { return d.label; }).join(" | ");
              var dBuilder = SpreadsheetApp.newRichTextValue().setText(fullTxt);
              var off = 0;
              driveLinks.forEach(function(d) {
                dBuilder.setLinkUrl(off, off + d.label.length, d.url);
                off += d.label.length + 3;
              });
              szotar[fId] = dBuilder.build();
              continue;
            }
          }
        } catch (e) {}
      }

      // D) Összeállítás
      if (talaltUrl) {
        var b = SpreadsheetApp.newRichTextValue().setText(szoveg || "1");
        b.setLinkUrl(0, (szoveg || "1").length, talaltUrl);
        szotar[fId] = b.build();
      } else if (szoveg) {
        szotar[fId] = SpreadsheetApp.newRichTextValue().setText(szoveg).build();
      }
    }
  }

  // 3. LÉPÉS: Beírjuk az új linkeket a Számolótáblába
  var kimenet = celMeglevoRichTextek; // Kiindulunk a meglévő állapotból
  var frissitettDb = 0;

  hianyzoSorok.forEach(function(elem) {
    if (szotar[elem.id]) {
      kimenet[elem.sorIndex] = [szotar[elem.id]];
      frissitettDb++;
    }
  });

  // Beírás a Q oszlopba (17. oszlop)
  celLap.getRange(2, 17, kimenet.length, 1).setRichTextValues(kimenet);

  if (SpreadsheetApp.getUi) {
    try {
      SpreadsheetApp.getUi().alert("Kész! " + frissitettDb + " db üres sorhoz sikeresen hozzárendeltük a blokk hiperlinket.");
    } catch (e) {}
  }
}