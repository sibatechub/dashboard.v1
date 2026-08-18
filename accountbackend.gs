/*************************************************
 * ACCOUNT BACKEND
 *************************************************/


/**
 * Get logged-in user's account details
 */
function getAccountDetails(sessionToken) {

  const user = getUserSession(sessionToken);

  if (!user) {
    throw new Error("Session expired. Please login again.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");

  const data = sheet.getDataRange().getValues();

  const loggedUserId = String(user.userId).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {

    const rowUserId =
      String(data[i][1]).trim().toUpperCase();

    if (rowUserId === loggedUserId) {

      return {

        userId: data[i][1],
        name: data[i][2],
        email: data[i][3],
        mobile: data[i][4],

        role: data[i][6],
        status: data[i][7]

      };

    }

  }

  throw new Error("User account not found.");
}


/**
 * Check whether User ID is already used
 *
 * excludeUserId = current logged-in User ID
 */
function checkAccountUserId(sessionToken, newUserId) {

  const user = getUserSession(sessionToken);

  if (!user) {
    throw new Error("Session expired. Please login again.");
  }

  newUserId = String(newUserId).trim();

  if (!newUserId) {

    return {
      available: false,
      message: "User ID cannot be empty."
    };

  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");

  const data = sheet.getDataRange().getValues();

  const currentUserId =
    String(user.userId).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {

    const existingUserId =
      String(data[i][1]).trim().toUpperCase();

    if (
      existingUserId === newUserId.toUpperCase() &&
      existingUserId !== currentUserId
    ) {

      return {

        available: false,

        message:
          "This User ID is already in use."

      };

    }

  }

  return {

    available: true,

    message:
      "User ID is available."

  };

}


/**
 * Check whether Email is already used
 *
 * excludeEmail = current logged-in Email
 */
function checkAccountEmail(sessionToken, newEmail) {

  const user = getUserSession(sessionToken);

  if (!user) {
    throw new Error("Session expired. Please login again.");
  }

  newEmail = String(newEmail).trim();

  if (!newEmail) {

    return {

      available: false,

      message:
        "Email cannot be empty."

    };

  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");

  const data = sheet.getDataRange().getValues();

  const currentEmail =
    String(user.email).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {

    const existingEmail =
      String(data[i][3]).trim().toLowerCase();

    if (
      existingEmail === newEmail.toLowerCase() &&
      existingEmail !== currentEmail
    ) {

      return {

        available: false,

        message:
          "This email address is already registered."

      };

    }

  }

  return {

    available: true,

    message:
      "Email address is available."

  };

}


/**
 * Update Profile
 *
 * Editable:
 * UserID
 * Name
 * Email
 * Mobile
 *
 * Not editable:
 * Role
 * Status
 * Password
 */
function updateAccountProfile(sessionToken, profile) {

  const user = getUserSession(sessionToken);

  if (!user) {
    throw new Error("Session expired. Please login again.");
  }

  if (!profile) {
    throw new Error("Invalid profile data.");
  }

  const newUserId =
    String(profile.userId || "").trim();

  const newName =
    String(profile.name || "").trim();

  const newEmail =
    String(profile.email || "").trim();

  const newMobile =
    String(profile.mobile || "").trim();


  if (!newUserId) {
    throw new Error("User ID is required.");
  }

  if (!newName) {
    throw new Error("Name is required.");
  }

  if (!newEmail) {
    throw new Error("Email is required.");
  }

  if (!newMobile) {
    throw new Error("Mobile number is required.");
  }


  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");

  const data = sheet.getDataRange().getValues();

  const currentUserId =
    String(user.userId).trim().toUpperCase();

  let userRow = -1;


  // Find current user
  for (let i = 1; i < data.length; i++) {

    const rowUserId =
      String(data[i][1]).trim().toUpperCase();

    if (rowUserId === currentUserId) {

      userRow = i + 1;

      break;

    }

  }


  if (userRow === -1) {

    throw new Error(
      "Current user account not found."
    );

  }


  /*
   * Check duplicate User ID
   */

  for (let i = 1; i < data.length; i++) {

    const existingUserId =
      String(data[i][1]).trim().toUpperCase();

    if (
      existingUserId === newUserId.toUpperCase() &&
      i + 1 !== userRow
    ) {

      throw new Error(
        "This User ID is already in use."
      );

    }

  }


  /*
   * Check duplicate Email
   */

  for (let i = 1; i < data.length; i++) {

    const existingEmail =
      String(data[i][3]).trim().toLowerCase();

    if (
      existingEmail === newEmail.toLowerCase() &&
      i + 1 !== userRow
    ) {

      throw new Error(
        "This email address is already registered."
      );

    }

  }


  /*
   * Update ONLY editable fields
   */

  sheet.getRange(userRow, 2)
       .setValue(newUserId);

  sheet.getRange(userRow, 3)
       .setValue(newName);

  sheet.getRange(userRow, 4)
       .setValue(newEmail);

  sheet.getRange(userRow, 5)
       .setValue(newMobile);


  return {

    success: true,

    message:
      "Profile updated successfully.",

    userId: newUserId,

    name: newName,

    email: newEmail,

    mobile: newMobile

  };

}


/**
 * Change Password
 */
function updateAccountPassword(
  sessionToken,
  currentPassword,
  newPassword
) {

  const user = getUserSession(sessionToken);

  if (!user) {
    throw new Error(
      "Session expired. Please login again."
    );
  }


  currentPassword =
    String(currentPassword || "");

  newPassword =
    String(newPassword || "");


  if (!currentPassword) {

    throw new Error(
      "Please enter your current password."
    );

  }

  if (!newPassword) {

    throw new Error(
      "Please enter a new password."
    );

  }

  if (newPassword === currentPassword) {

    throw new Error(
      "New password must be different from current password."
    );

  }


  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");

  const data = sheet.getDataRange().getValues();

  const currentUserId =
    String(user.userId).trim().toUpperCase();


  for (let i = 1; i < data.length; i++) {

    const rowUserId =
      String(data[i][1]).trim().toUpperCase();


    if (rowUserId === currentUserId) {

      const storedPassword =
        String(data[i][5] || "");


      /*
       * Verify current password
       */

      if (storedPassword !== currentPassword) {

        throw new Error(
          "Current password is incorrect."
        );

      }


      /*
       * Save new password
       */

      sheet.getRange(i + 1, 6)
           .setValue(newPassword);


      return {

        success: true,

        message:
          "Password changed successfully."

      };

    }

  }


  throw new Error(
    "User account not found."
  );

}
