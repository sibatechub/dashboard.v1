/**
 * =====================================================
 * ALL TICKETS BACKEND
 * DRIEMS IT Dashboard
 * =====================================================
 */

/*************************************************
 * FORMAT DATE & TIME
 *************************************************/
function formatDateTime(value) {

  if (!value) return "";

  return Utilities.formatDate(
    new Date(value),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy hh:mm:ss a"
  );

}


/*************************************************
 * GET ALL TICKETS
 *************************************************/
function getAllTickets(sessionToken) {

  try {

    // Open Spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Open Master Sheet
    const sheet = ss.getSheetByName(MASTER_SHEET);

    if (!sheet) {
      throw new Error("Master sheet not found : " + MASTER_SHEET);
    }

    // Read all data
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    // -------------------------------
    // Create Header Map
    // -------------------------------

    const headers = data.shift();

    const COL = {};

    headers.forEach(function(header, index) {

      COL[String(header).trim()] = index;

    });

    // -------------------------------
    // Convert rows into JSON
    // -------------------------------

    const tickets = data.map(function(row) {

      return {

        // Ticket Information
        timestamp: formatDateTime(row[COL["Timestamp"]]),
        ticketNo: row[COL["Ticket No"]],
        status: row[COL["Status"]],
        priority: row[COL["Issue Priority"]],

        // Faculty Details
        faculty: row[COL["Faculty / Staff Name"]],
        mobile: row[COL["Mobile Number"]],
        email: row[COL["Email Address"]],

        // Location
        building: row[COL["Building Name"]],
        floor: row[COL["Floor"]],
        room: row[COL["Room / Lab Name"]],

        // Asset
        assetType: row[COL["Asset Type"]],
        assetNo: row[COL["Asset Number (Enter the exact Asset QR Code Number, e.g., DRM/XXX/XXX/XXX)"]],

        // Problem
        problem: row[COL["Problem Description"]],
        upload: row[COL["Upload Photo / Video"]],

        // Technician
        technician: row[COL["Technician Name"]],
        technicianEmail: row[COL["Technician Email"]],
        technicianMobile: row[COL["Technician Mobile No"]],

        // Timeline
        acceptedTime: formatDateTime(row[COL["Ticket Accepted Time"]]),
        closeTime: formatDateTime(row[COL["Ticket Close Time"]]),
        duration: row[COL["Duration"]],

        // Comments
        comments: row[COL["Comments / Notes (Only Use for Closed or Pending)"]],

        // Feedback
        satisfaction: row[COL["Overall Satisfaction"]],
        behaviour: row[COL["Technician Behaviour"]],
        resolutionSatisfaction: row[COL["Resolution Time Satisfaction"]],
        additionalComments: row[COL["Additional Comments / Suggestions"]],

        feedbackTime: formatDateTime(row[COL["Feedback Submit Time"]]),

        // Technician Change
        newTechnician: row[COL["New Technician Name"]],
        newTechnicianEmail: row[COL["New Technician Email"]],
        newTechnicianMobile: row[COL["New Technician Mobile"]],

        // IT Head
        itHeadRemarks: row[COL["IT Head Remarks / Instructions"]]

      };

    });

    // ======================================
// Get Logged-in User Session
// ======================================

const session = getUserSession(sessionToken);

if (!session) {
  throw new Error("User session not found. Please login again.");
}

// ======================================
// Admin can view all tickets
// ======================================

if (String(session.role).toUpperCase() === "ADMIN") {
  return tickets;
}

// ======================================
// Technician can view only assigned tickets
// ======================================

return tickets.filter(ticket => {

  const assignedEmail =
    ticket.newTechnicianEmail || ticket.technicianEmail;

  return String(assignedEmail || "")
    .trim()
    .toLowerCase() ===
    String(session.email || "")
    .trim()
    .toLowerCase();

});

  } catch (error) {

    Logger.log(error);

    throw new Error(error.message);

  }

}
