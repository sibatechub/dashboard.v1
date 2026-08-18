function getTechnicianUpdateLink(sessionToken, ticketNo) {

  try {

    // ============================================
    // CHECK LOGIN SESSION
    // ============================================

    const session = getUserSession(sessionToken);

    if (!session) {
      throw new Error("Session expired. Please login again.");
    }

    // ============================================
    // VALIDATE TICKET NUMBER
    // ============================================

    ticketNo = String(ticketNo || "").trim();

    if (!ticketNo) {
      throw new Error("Ticket number is missing.");
    }

    // ============================================
    // GET TICKET DATA
    // ============================================

    const tickets = getAllTickets(sessionToken);

    if (!Array.isArray(tickets)) {
      throw new Error("Unable to load ticket information.");
    }

    const ticket = tickets.find(function(t) {

      return String(t.ticketNo || "").trim() === ticketNo;

    });

    if (!ticket) {

      throw new Error(
        "Ticket not found."
      );

    }

    // ============================================
    // DO NOT UPDATE CLOSED TICKET
    // ============================================

    if (
      String(ticket.status || "")
        .trim()
        .toUpperCase() === "CLOSED"
    ) {

      throw new Error(
        "This ticket is already closed."
      );

    }

    // ============================================
    // EFFECTIVE TECHNICIAN
    // ============================================
    // If reassigned, use new technician details.
    // Otherwise use original technician.

    const techName =
      ticket.newTechnician ||
      ticket.technician ||
      "";

    const techEmail =
      ticket.newTechnicianEmail ||
      ticket.technicianEmail ||
      "";

    const techMobile =
      ticket.newTechnicianMobile ||
      ticket.technicianMobile ||
      "";

    // ============================================
    // SAME GOOGLE FORM USED IN EMAIL
    // ============================================

    const formBaseUrl =
      "https://docs.google.com/forms/d/e/1FAIpQLSfBBacWI4YWI8x3Q7RbLZ6ssf1cmknqXYnHWPxuPGFDmkiCwg/viewform?usp=pp_url";

    // Google Form Entry IDs

    const ticketFieldId =
      "entry.327371081";

    const techNameFieldId =
      "entry.748529779";

    const techMobileFieldId =
      "entry.1950349758";

    const techEmailFieldId =
      "entry.1640227390";

    // ============================================
    // CREATE PREFILLED FORM LINK
    // ============================================

    const preFilledLink =
      formBaseUrl +
      "&" +
      ticketFieldId +
      "=" +
      encodeURIComponent(ticketNo) +

      "&" +
      techNameFieldId +
      "=" +
      encodeURIComponent(techName) +

      "&" +
      techMobileFieldId +
      "=" +
      encodeURIComponent(techMobile) +

      "&" +
      techEmailFieldId +
      "=" +
      encodeURIComponent(techEmail);

    // ============================================
    // RETURN
    // ============================================

    return {

      success: true,

      ticketNo: ticketNo,

      status: ticket.status,

      technician: techName,

      formUrl: preFilledLink

    };

  } catch (error) {

    Logger.log(
      "Technician Update Link Error: " +
      error.message
    );

    throw new Error(error.message);

  }

}
