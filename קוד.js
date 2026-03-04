function ping() { return "pong"}

function crm_setupCore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Clients");
  if (!sh) sh = ss.insertSheet("Clients");
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([["CLIENT_ID", "FULL_NAME", "PHONE"]]);
    sh.setFrozenRows(1);
  }
  return "ok";
}