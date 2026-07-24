/**
 * ExpenseTracker
 * Google Drive szolgáltatás
 */

/**
 * ExpenseTracker mappa elérése
 */
function getExpenseFolder() {
  const folder = DriveApp.getFolderById(APP.DRIVE.FOLDER_ID);
  return folder;
}

/**
 * Drive kapcsolat teszt
 */
function testDriveConnection() {
  const folder = getExpenseFolder();
  Logger.log("Drive mappa neve: " + folder.getName());
}

/**
 * Drive feltöltés teszt
 */
function testDriveUpload() {
  const text = "ExpenseTracker teszt fájl";
  const blob = Utilities.newBlob(text, "text/plain", "teszt.txt");
  const folder = getExpenseFolder();
  const file = folder.createFile(blob);

  Logger.log("Fájl létrehozva: " + file.getName());
}

/**
 * Kép feltöltése Drive-ba
 */
function uploadExpenseImage(imageData, expenseId, imageNumber) {
  const folder = getExpenseFolder();
  
  const contentType = imageData.match(/data:(.*);base64/)[1];
  const base64 = imageData.split(",")[1];
  const bytes = Utilities.base64Decode(base64);
  const extension = contentType.split("/")[1];

  const fileName = expenseId + "_" + imageNumber + "." + extension;
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = folder.createFile(blob);

  // Megtekintési jogosultság beállítása a link birtokosainak
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}