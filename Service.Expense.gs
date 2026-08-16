/* ==========================================
   KIADÁS KEZELŐ SZERVIZ (Service.Expense.gs)
   ========================================== */

/**
 * Kliens oldali (google.script.run) belépési pontok
 */
function saveExpense(expense, imageUrls = [], customId = null, gpsCoords = null) {
  return saveExpenseData(expense, imageUrls, customId, gpsCoords);
}

function getExpenseId(costCenter) {
  return generateExpenseId(costCenter);
}

/**
 * Kiadás adatok elmentése a Google Sheetbe
 */
function saveExpenseData(expense, imageUrls = [], customId = null, gpsCoords = null) {
  // Párhuzamos mentések kezelése (konkurrencia-védelem)
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Max 10 mp várakozás

    // 1. Jogosultság ellenőrzése
    const auth = getUserAuth();
    if (!auth || !auth.authorized) {
      throw new Error("Jogosulatlan hozzáférés!");
    }

    // 2. Bemeneti adatok tisztítása és ellenőrzése
    if (!expense) throw new Error("A kiadás adatai hiányoznak!");

    const amountNum = Number(String(expense.amount).replace(",", "."));
    if (isNaN(amountNum) || amountNum <= 0) {
      throw new Error("Érvénytelen összeget adott meg!");
    }

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.EXPENSES);
    if (!sheet) throw new Error(`A(z) '${APP.SHEETS.EXPENSES}' munkalap nem található!`);

    const id = customId || generateExpenseId(expense.costCenter);

    // 3. Projekt attribútumok keresése
    let projectType = "";
    let projectId = "";
    if (typeof getActiveProjectsData === "function") {
      const activeProjects = getActiveProjectsData(auth.csoportId, ss) || [];
      const matchedProj = activeProjects.find(p => p.name === expense.costCenter);
      if (matchedProj) {
        projectType = matchedProj.type || "";
        projectId = matchedProj.id || "";
      }
    }

    // Számla mező meghatározása ("Van" / "Nincs")
    const szamlaErtek = (expense.hasInvoice === true || expense.hasInvoice === "Van") ? "Van" : "Nincs";

    // 4. Sor hozzáadása a munkalaphoz (A - O oszlopok)
    sheet.appendRow([
      auth.csoportId || "",            // A: Csoport_ID
      expense.date || "",              // B: Dátum
      expense.costCenter || "",        // C: Költséghely
      amountNum,                       // D: Összeg (számként tárolva)
      expense.paymentMethod || "",     // E: Fizetési mód
      "",                              // F: Blokk link (RichText)
      expense.comment || "",           // G: Megjegyzés
      auth.email || "",                // H: Beküldő email
      "",                              // I: GPS (RichText)
      new Date(),                      // J: Feltöltés ideje
      id,                              // K: Egyedi azonosító
      projectType,                     // L: Projekt típusa
      projectId,                       // M: Projekt_ID
      expense.extraCategory || "",     // N: Egyedi beállítások (F2:F)
      szamlaErtek                      // O: Számla (Van / Nincs) 👈 ÚJ OSZLOP!
    ]);

    const lastRow = sheet.getLastRow();

    // 5. Blokk hivatkozások RichText formázása (F oszlop)
    if (Array.isArray(imageUrls) && imageUrls.length > 0) {
      const labels = imageUrls.map((_, i) => String(i + 1));
      const fullText = labels.join(" | ");
      const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);

      let currentOffset = 0;
      imageUrls.forEach((url, i) => {
        const label = String(i + 1);
        const start = currentOffset;
        const end = start + label.length;
        if (url) richTextBuilder.setLinkUrl(start, end, url);
        currentOffset = end + 3; // " | " elválasztó hossza
      });

      sheet.getRange(lastRow, 6).setRichTextValue(richTextBuilder.build());
    }

    // 6. GPS koordináta térkép link formázása (I oszlop)
    if (gpsCoords && gpsCoords.lat && gpsCoords.lng) {
      const mapUrl = `https://maps.google.com/?q=${gpsCoords.lat},${gpsCoords.lng}`;
      const gpsRichText = SpreadsheetApp.newRichTextValue()
        .setText("📍 Térkép")
        .setLinkUrl(mapUrl)
        .build();
      sheet.getRange(lastRow, 9).setRichTextValue(gpsRichText);
    }

    if (typeof writeLog === "function") {
      writeLog("INFO", "Expense", "Kiadás mentve: " + id + " | Számla: " + szamlaErtek);
    }

    return { success: true, id: id, message: "Mentés sikeres" };

  } catch (error) {
    if (typeof writeLog === "function") {
      writeLog("ERROR", "Expense", error.message);
    }
    throw error; // Így kapja meg a kliens oldali .withFailureHandler()
  } finally {
    lock.releaseLock();
  }
}

/**
 * Egyedi kiadás-azonosító generálása
 */
function generateExpenseId(costCenter) {
  const timeZone = (typeof APP !== "undefined" && APP.TIMEZONE) ? APP.TIMEZONE : Session.getScriptTimeZone();
  const timeStamp = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");
  const randomSuffix = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return "EXP-" + timeStamp + "-" + randomSuffix;
}