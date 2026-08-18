

/*************************************************
 * GET USER SESSION BY TOKEN
 *************************************************/
function getUserSession(sessionToken) {

  Logger.log("Incoming Token = [" + sessionToken + "]");
  Logger.log("Length = " + sessionToken.length);

  const sheet = getSessionSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    Logger.log("Sheet Token = [" + data[i][0] + "]");
    Logger.log("Length = " + String(data[i][0]).length);

    if (String(data[i][0]) === String(sessionToken)) {

      Logger.log("MATCH FOUND");

      return {
        userId: data[i][1],
        name: data[i][2],
        email: data[i][3],
        role: data[i][4]
      };
    }
  }

  Logger.log("NO MATCH");
  return null;
}

function logoutUser(sessionToken) {

  if (!sessionToken) {
    return true;
  }

  const sheet = getSessionSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == sessionToken) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return true;
}
/*************************************************
 * GET USER SESSION SHEET
 *************************************************/
function getSessionSheet() {

  return SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("User_Sessions");

}
/*************************************************
 * GENERATE SESSION TOKEN
 *************************************************/
function generateSessionToken() {

  return Utilities.getUuid() + "-" + new Date().getTime();

}
/*************************************************
 * CREATE NEW SESSION
 *************************************************/
function createSession(user) {

  const sheet = getSessionSheet();

  const token = generateSessionToken();

  const now = new Date();

  sheet.appendRow([
    token,
    user.userId,
    user.name,
    user.email,
    user.role,
    now,
    now
  ]);

  return token;

}











