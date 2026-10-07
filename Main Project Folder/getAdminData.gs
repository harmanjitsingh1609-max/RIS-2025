function getAdminData() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const data =
    sheet.getDataRange().getValues();


  // ============================================================
  // ADMIN EMAILS
  // ============================================================

  const ADMIN_EMAILS = [
    "rankings@thapar.edu"
  ];


  const userEmail =
    Session.getActiveUser().getEmail()
      .trim()
      .toLowerCase();


  // ============================================================
  // SECURITY CHECK
  // ============================================================

  if (!ADMIN_EMAILS.includes(userEmail)) {

    return {
      authorized: false,
      email: userEmail
    };

  }


  // ============================================================
  // COLUMN INDEXES (ZERO-BASED)
  // ============================================================

  const AUTHOR_COL = 6;           // G
  const ECODE_COL = 9;            // J
  const DISBURSEMENT_COL = 18;    // S
  const EMAIL_COL = 20;           // U
  const CONFIRMATION_COL = 21;    // V
  const DATE_COL = 22;            // W
  const CONFIRM_EMAIL_COL = 23;   // X


  // ============================================================
  // GROUP AUTHORS
  // ============================================================

  const authors = {};


  for (let i = 1; i < data.length; i++) {

    const row = data[i];


    const email =
      String(row[EMAIL_COL] || "")
        .trim()
        .toLowerCase();


    // Ignore rows without email
    if (!email) {
      continue;
    }


    const amount =
      Number(row[DISBURSEMENT_COL]) || 0;


    // Ignore zero-value publications
    if (amount <= 0) {
      continue;
    }


    // Create author record
    if (!authors[email]) {

      authors[email] = {

        email: email,

        authorName:
          String(row[AUTHOR_COL] || ""),

        ecode:
          String(row[ECODE_COL] || ""),

        papers: 0,

        totalAmount: 0,

        confirmedPapers: 0,

        confirmationDate: "",

        confirmationEmail: ""

      };

    }


    const author =
      authors[email];


    // Count publication
    author.papers++;


    // Add amount
    author.totalAmount += amount;


    // ==========================================================
    // CONFIRMATION
    // ==========================================================

    const confirmation =
      String(row[CONFIRMATION_COL] || "")
        .trim()
        .toUpperCase();


    if (confirmation === "CONFIRMED") {

      author.confirmedPapers++;


      if (!author.confirmationDate && row[DATE_COL]) {

        author.confirmationDate =
          Utilities.formatDate(
            new Date(row[DATE_COL]),
            Session.getScriptTimeZone(),
            "dd MMMM yyyy, hh:mm a"
          );

      }


      if (
        !author.confirmationEmail &&
        row[CONFIRM_EMAIL_COL]
      ) {

        author.confirmationEmail =
          String(row[CONFIRM_EMAIL_COL]);

      }

    }

  }


  // ============================================================
  // CREATE AUTHOR LIST
  // ============================================================

  const authorList =
    Object.values(authors);


  // ============================================================
  // CALCULATE STATUS
  // ============================================================

  authorList.forEach(function(author) {

    author.confirmed =
      author.papers > 0 &&
      author.confirmedPapers === author.papers;

  });


  // ============================================================
  // SUMMARY
  // ============================================================

  const totalAuthors =
    authorList.length;


  const confirmedAuthors =
    authorList.filter(function(author) {

      return author.confirmed;

    }).length;


  const pendingAuthors =
    totalAuthors - confirmedAuthors;


  const totalAmount =
    authorList.reduce(function(sum, author) {

      return sum + author.totalAmount;

    }, 0);


  const confirmedAmount =
    authorList
      .filter(function(author) {

        return author.confirmed;

      })
      .reduce(function(sum, author) {

        return sum + author.totalAmount;

      }, 0);


  const pendingAmount =
    totalAmount - confirmedAmount;


  // ============================================================
  // RETURN DATA
  // ============================================================

  return {

    authorized: true,

    adminEmail: userEmail,

    summary: {

      totalAuthors: totalAuthors,

      confirmedAuthors: confirmedAuthors,

      pendingAuthors: pendingAuthors,

      totalAmount: totalAmount,

      confirmedAmount: confirmedAmount,

      pendingAmount: pendingAmount

    },

    authors: authorList

  };

}
