function exportTicketsToExcel(ticketData) {

  if (!ticketData || ticketData.length === 0) {
    throw new Error("No ticket data received.");
  }

  // Create Temporary Spreadsheet
  const ss = SpreadsheetApp.create(
    "DRIEMS IT Tickets - " +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd-MMM-yyyy HH:mm:ss"
    )
  );

  const sheet = ss.getSheets()[0];
  sheet.setName("Tickets");

  // ===========================
  // Headers
  // ===========================

  const headers = [

    "Ticket No",
    "Status",
    "Priority",
    "Faculty / Staff",
    "Mobile Number",
    "Building",
    "Floor",
    "Room",
    "Asset Type",
    "Asset Number",
    "Problem Description",
    "Technician",
    "Created Date",
    "Ticket Accepted Time",
    "Ticket Close Time",
    "Duration",
    "Comments",
    "Overall Satisfaction",
    "Technician Behaviour",
    "Resolution Time Satisfaction",
    "Additional Comments / Suggestions"

  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ===========================
  // Header Formatting
  // ===========================

  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  headerRange
    .setBackground("#2E2278")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Freeze Header Row
  sheet.setFrozenRows(1);

  // ===========================
  // Write Data
  // ===========================

  const rows = ticketData.map(ticket => [

    ticket.ticketNo,
    ticket.status,
    ticket.priority,
    ticket.faculty,
    ticket.mobile,
    ticket.building,
    ticket.floor,
    ticket.room,
    ticket.assetType,
    ticket.assetNo,
    ticket.problem,
    ticket.technician,
    ticket.timestamp,
    ticket.acceptedTime,
    ticket.closeTime,
    ticket.duration,
    ticket.comments,
    ticket.satisfaction,
    ticket.behaviour,
    ticket.resolutionSatisfaction,
    ticket.additionalComments

  ]);

  if (rows.length > 0) {

    sheet.getRange(
      2,
      1,
      rows.length,
      headers.length
    ).setValues(rows);

  }

  // ===========================
  // Auto Resize Columns
  // ===========================

  for (let col = 1; col <= headers.length; col++) {

    sheet.autoResizeColumn(col);

  }

  // ===========================
  // Apply Filter
  // ===========================

  sheet
    .getRange(1, 1, sheet.getLastRow(), headers.length)
    .createFilter();

  // ===========================
  // Alternate Row Colors
  // ===========================

  if (rows.length > 0) {

    sheet
      .getRange(1, 1, rows.length + 1, headers.length)
      .applyRowBanding(
        SpreadsheetApp.BandingTheme.LIGHT_GREY
      );

  }

  // ===========================
  // Return Spreadsheet ID
  // ===========================

  // ===========================================
  // Convert Google Spreadsheet to Excel (.xlsx)
  // ===========================================

const url =
  "https://www.googleapis.com/drive/v3/files/" +
  ss.getId() +
  "/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const token = ScriptApp.getOAuthToken();

const response = UrlFetchApp.fetch(url, {
  headers: {
    Authorization: "Bearer " + token
  }
});

const excelBlob = response.getBlob();

excelBlob.setName(
  "DRIEMS_IT_Tickets_" +
  Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd_HH-mm"
  ) +
  ".xlsx"
);

// ===========================================
// CONVERT EXCEL TO BASE64
// ===========================================

const base64Data = Utilities.base64Encode(
  excelBlob.getBytes()
);

// Delete temporary Google Spreadsheet
DriveApp.getFileById(ss.getId()).setTrashed(true);

// Return Excel directly to browser
return {

  success: true,

  data: base64Data,

  fileName: excelBlob.getName(),

  mimeType:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

};

}
