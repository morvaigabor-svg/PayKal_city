/**
 * GenerateReports.gs
 * Automatizált olvasási kimutatások, Drive parancsikonok és blokk-linkek generálása
 */

const REPORT_SYSTEM_CONFIG = {
  // IDE MÁSOLD A FŐ GYŰJTŐMAPPA ID-JÁT, AHOL A SHEETS FÁJLOK LÉTREJÖNNEK:
  MASTER_REPORTS_FOLDER_ID: "19yARIVLE08f8ZO1ZYxsdqhZyfdC27s0Q"
};

/**
 * Megnyitja a megemelt méretű párbeszédablakot a Picker miatt
 */
function openReportGeneratorDialog() {
  const html = HtmlService.createTemplateFromFile('ReportDialog')
    .evaluate()
    .setWidth(1150)
    .setHeight(720)
    .setTitle('Olvasási Kimutatások Generálása');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Olvasási Kimutatások Generálása');
}

/**
 * Biztonsági OAuth Token átadása a Google Picker API-nak
 */
function getOAuthToken() {
  DriveApp.getRootFolder(); // Kikényszeríti a Drive jogosultságot
  return ScriptApp.getOAuthToken();
}

/**
 * Adatok beolvasása az aktív táblázatból (Csoportok és Városok)
 */
function getReportOptionsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Csoportok beolvasása ('Csoportok' lap A2:A)
  const csSheet = ss.getSheetByName('Csoportok') || ss.getSheetByName('csoportok');
  const csoportok = [];
  if (csSheet && csSheet.getLastRow() >= 2) {
    const csVals = csSheet.getRange(2, 1, csSheet.getLastRow() - 1, 1).getValues();
    csVals.forEach(row => { if (row[0]) csoportok.push(String(row[0]).trim()); });
  }

  // 2. Városok beolvasása ('Városok' lap A2:A)
  const vSheet = ss.getSheetByName('Városok') || ss.getSheetByName('városok');
  const varosok = [];
  if (vSheet && vSheet.getLastRow() >= 2) {
    const vVals = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 1).getValues();
    vVals.forEach(row => { if (row[0]) varosok.push(String(row[0]).trim()); });
  }

  return { csoportok: csoportok, varosok: varosok };
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

/**
 * Kimutatások, parancsikonok ÉS blokk-linkek generálása
 */
function processReportGeneration(payload) {
  try {
    // 1. ELŐ-LÉPÉS: Blokk-linkek frissítése a Fő Táblázat L oszlopában
    frissitsBlokkLinkeket();

    const masterSS = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheetId = masterSS.getId();
    const masterFolder = DriveApp.getFolderById(REPORT_SYSTEM_CONFIG.MASTER_REPORTS_FOLDER_ID);
    
    // 2. Költségek blokk-linkjeinek beolvasása szótárba (ID -> RichText)
    const koltsegekLap = masterSS.getSheetByName('Költségek');
    const blokkSzotar = {};
    if (koltsegekLap && koltsegekLap.getLastRow() >= 2) {
      const kIdek = koltsegekLap.getRange(2, 11, koltsegekLap.getLastRow() - 1, 1).getValues();
      const kRich = koltsegekLap.getRange(2, 6, koltsegekLap.getLastRow() - 1, 1).getRichTextValues();
      for (let i = 0; i < kIdek.length; i++) {
        const id = String(kIdek[i][0]).trim();
        if (id) blokkSzotar[id] = kRich[i][0];
      }
    }

    // 3. Számolótábla adatainak beolvasása a sorrend tartásához (Col1..Col10)
    const szamoloLap = masterSS.getSheetByName('Számolótábla');
    let szamoloAdatok = [];
    if (szamoloLap && szamoloLap.getLastRow() >= 2) {
      szamoloAdatok = szamoloLap.getRange(2, 1, szamoloLap.getLastRow() - 1, 10).getValues();
    }

    const results = [];
    const uresRichText = SpreadsheetApp.newRichTextValue().setText("").build();

    // 4. Riportok generálása ciklusban
    payload.items.forEach(item => {
      try {
        const isGroup = (payload.type === 'GROUP');
        const titleName = isGroup ? `Kimutatás - ${item.id}` : `Városi Kimutatás - ${item.id}`;
        
        // A) Új Google Sheet létrehozása
        const newSS = SpreadsheetApp.create(titleName);
        const newFile = DriveApp.getFileById(newSS.getId());
        newFile.moveTo(masterFolder);
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        // B) Munkalap és Fejlécek (12 oszlop: A-L)
        const sheet = newSS.getActiveSheet();
        sheet.setName("Számolótábla");

        const headers = [[
          "Csoport_ID", "Költési azonosító", "Dátum", "Tranzakció célja", 
          "Tranzakció típusa", "Fizetési mód", "Egyenleg változás (Ft)", 
          "Készpénz változás", "Számla változás", "Város", "Egyéb / K oszlop", "Blokkok"
        ]];
        sheet.getRange(1, 1, 1, 12).setValues(headers).setFontWeight("bold").setBackground("#eef2f7");

        // C) IMPORTRANGE + QUERY képlet beírása (A2:K tartományra)
        const filterCol = isGroup ? "Col1" : "Col10";
        const formulaLocal = `=QUERY(IMPORTRANGE("${masterSheetId}"; "Számolótábla!A2:K"); "SELECT * WHERE ${filterCol} = '${item.id}'"; 0)`;
        sheet.getRange("A2").setFormulaLocal(formulaLocal);

        // D) Blokk-linkek másolása pontosan az L oszlopba (12. oszlop)
        const riportBlokkok = [];
        const itemIdClean = String(item.id).trim();

        szamoloAdatok.forEach(row => {
          // Csoport esetén Col1 (index 0), Város esetén Col10 (index 9)
          const matchVal = isGroup ? String(row[0]).trim() : String(row[9]).trim();
          if (matchVal === itemIdClean) {
            const tranzakcioId = String(row[1]).trim(); // Col2 = Költési azonosító
            if (tranzakcioId && blokkSzotar[tranzakcioId]) {
              riportBlokkok.push([blokkSzotar[tranzakcioId]]);
            } else {
              riportBlokkok.push([uresRichText]);
            }
          }
        });

        if (riportBlokkok.length > 0) {
          sheet.getRange(2, 12, riportBlokkok.length, 1).setRichTextValues(riportBlokkok); // 12. oszlop (L oszlop)
        }

        // E) Drive Parancsikon létrehozása
        if (item.targetFolderId && item.targetFolderId !== 'root') {
          const targetFolder = DriveApp.getFolderById(item.targetFolderId);
          targetFolder.createShortcut(newSS.getId());
        }

        results.push({ name: item.id, status: 'OK' });
      } catch (err) {
        results.push({ name: item.id, status: 'HIBA: ' + err.message });
      }
    });

    return { success: true, results: results };
  } catch (globalErr) {
    return { success: false, message: globalErr.toString() };
  }
}