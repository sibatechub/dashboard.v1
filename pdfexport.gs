function exportTicketsToPDF(ticketData) {

  if (!ticketData || ticketData.length === 0) {
    throw new Error("No ticket data received.");
  }

  // Create Temporary Spreadsheet
  const ss = SpreadsheetApp.create(
    "DRIEMS IT Ticket PDF Report - " +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd-MMM-yyyy HH:mm:ss"
    )
  );

  const sheet = ss.getSheets()[0];
  sheet.setName("Ticket Report");

  // =====================================
  // REPORT TITLE
  // =====================================

  sheet.getRange("A1:M1").merge();
  sheet.getRange("A1")
      .setValue("DRIEMS UNIVERSITY")
      .setFontSize(18)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontColor("#2E2278");

  sheet.getRange("A2:M2").merge();
  sheet.getRange("A2")
      .setValue("IT BREAKDOWN TICKET REPORT")
      .setFontSize(14)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontColor("#2E2278");

  sheet.getRange("A4")
      .setValue("Generated On :");

  sheet.getRange("B4")
      .setValue(
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "dd-MMM-yyyy hh:mm a"
        )
      );

  sheet.getRange("F4")
      .setValue("Total Tickets :");

  sheet.getRange("G4")
      .setValue(ticketData.length);

  // =====================================
  // TABLE HEADER
  // =====================================

  const headers = [

    "Ticket No",
    "Status",
    "Priority",
    "Faculty",
    "Mobile",
    "Building",
    "Floor",
    "Room",
    "Created Date",
    "Problem",
    "Technician",
    "Comments",
    "Close Time"

  ];

  sheet
    .getRange(6,1,1,headers.length)
    .setValues([headers]);

sheet
    .getRange(6,1,1,headers.length)
    .setBackground("#2E2278")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // =====================================
  // WRITE DATA
  // =====================================

  const rows = ticketData.map(ticket => [

    ticket.ticketNo,
    ticket.status,
    ticket.priority,
    ticket.faculty,
    ticket.mobile,
    ticket.building,
    ticket.floor,
    ticket.room,
    ticket.timestamp,
    ticket.problem,
    ticket.technician,
    ticket.comments,
    ticket.closeTime

  ]);

  if(rows.length){

    sheet
      .getRange(
        7,
        1,
        rows.length,
        headers.length
      )
      .setValues(rows);

  }

  // =====================================
  // FORMAT
  // =====================================

  sheet.setFrozenRows(6);

// Set fixed column widths
sheet.setColumnWidth(1, 170);   // Ticket No
sheet.setColumnWidth(2, 90);    // Status
sheet.setColumnWidth(3, 90);    // Priority
sheet.setColumnWidth(4, 180);   // Faculty
sheet.setColumnWidth(5, 120);   // Mobile
sheet.setColumnWidth(6, 160);   // Building
sheet.setColumnWidth(7, 70);    // Floor
sheet.setColumnWidth(8, 90);    // Room
sheet.setColumnWidth(9, 130);   // Created Date
sheet.setColumnWidth(10, 300);  // Problem
sheet.setColumnWidth(11, 150);  // Technician
sheet.setColumnWidth(12, 300);  // Comments
sheet.setColumnWidth(13, 150);  // Close Time
sheet.getRange(7, 1, rows.length, headers.length)
     .setWrap(true)
     .setVerticalAlignment("top");
     for (let r = 7; r < rows.length + 7; r++) {
  sheet.setRowHeight(r, 60);
}

  sheet
    .getRange(
      6,
      1,
      rows.length+1,
      headers.length
    )
    .applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY
    );
    // =====================================
// EXPORT TO PDF
// =====================================

SpreadsheetApp.flush();

const pdfUrl =
  "https://docs.google.com/spreadsheets/d/" +
  ss.getId() +
  "/export?" +
  "format=pdf" +
  "&size=A4" +
  "&portrait=false" +
  "&fitw=true" +
  "&sheetnames=false" +
  "&printtitle=false" +
  "&pagenumbers=true" +
  "&gridlines=false" +
  "&fzr=true" +
  "&gid=" + sheet.getSheetId();

const token = ScriptApp.getOAuthToken();

const response = UrlFetchApp.fetch(pdfUrl, {
  headers: {
    Authorization: "Bearer " + token
  }
});

const pdfBlob = response.getBlob().setName(
  "IT_Breakdown_Report_" +
  Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "ddMMyyyy_HHmmss"
  ) +
  ".pdf"
);

// Convert PDF to Base64
const base64Data = Utilities.base64Encode(
  pdfBlob.getBytes()
);

// Delete temporary spreadsheet
DriveApp.getFileById(ss.getId()).setTrashed(true);

// Return PDF directly to browser
return {
  success: true,
  data: base64Data,
  fileName: pdfBlob.getName(),
  mimeType: "application/pdf"
};

}
