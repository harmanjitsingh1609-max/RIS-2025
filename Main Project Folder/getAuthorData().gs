function getAuthorData() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const data =
    sheet.getDataRange().getValues();

  const userEmail =
    Session.getActiveUser().getEmail()
      .trim()
      .toLowerCase();


  // ============================================================
  // COLUMN INDEXES (ZERO-BASED)
  // ============================================================

  const MASTER_COL = 0;          // A
  const TITLE_COL = 2;           // C
  const EDITION_COL = 3;         // D
  const IMPACT_COL = 4;          // E
  const QUARTILE_COL = 5;        // F
  const AUTHOR_COL = 6;          // G
  const ECODE_COL = 9;           // J
  const DISBURSEMENT_COL = 18;   // S
  const EMAIL_COL = 20;          // U

  const CONFIRMATION_COL = 21;   // V
  const DATE_COL = 22;           // W
  const CONFIRM_EMAIL_COL = 23;  // X


  // ============================================================
  // BUILD PUBLICATION INFORMATION
  //
  // A publication starts whenever Master Sr. No. is present.
  // All following author rows belong to that publication until
  // the next Master Sr. No.
  // ============================================================

  const publicationByRow = {};

  let currentPublication = null;

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const master =
      String(row[MASTER_COL] || "").trim();

    const title =
      String(row[TITLE_COL] || "").trim();

    const edition =
      String(row[EDITION_COL] || "").trim();

    const impactFactor =
      String(row[IMPACT_COL] || "").trim();

    const quartile =
      String(row[QUARTILE_COL] || "").trim();


    // ----------------------------------------------------------
    // New publication
    // ----------------------------------------------------------

    if (master !== "") {

      currentPublication = {
        master: master,
        title: title,
        edition: edition,
        impactFactor: impactFactor,
        quartile: quartile,
        authorCount: 0
      };

    }


    // ----------------------------------------------------------
    // Store publication information against this row
    // ----------------------------------------------------------

    if (currentPublication) {

      currentPublication.authorCount++;

      publicationByRow[i] = currentPublication;

    }

  }


  // ============================================================
  // VARIABLES
  // ============================================================

  const papers = [];

  let authorName = "";
  let ecode = "";

  let eligibleRows = 0;
  let confirmedRows = 0;

  let confirmationDate = "";
  let confirmationEmail = "";


  // ============================================================
  // READ AUTHOR RECORDS
  // ============================================================

  for (let i = 1; i < data.length; i++) {

    const row = data[i];


    // ==========================================================
    // READ EMAIL
    // ==========================================================

    const rowEmail =
      String(row[EMAIL_COL] || "")
        .trim()
        .toLowerCase();


    // Ignore rows without email

    if (!rowEmail) {
      continue;
    }


    // ==========================================================
    // ONLY RETRIEVE AUTHENTICATED USER'S ROWS
    // ==========================================================

    if (rowEmail !== userEmail) {
      continue;
    }


    // ==========================================================
    // READ AUTHOR INFORMATION
    // ==========================================================

    if (!authorName) {
      authorName =
        String(row[AUTHOR_COL] || "");
    }


    if (!ecode) {
      ecode =
        String(row[ECODE_COL] || "");
    }


    // ==========================================================
    // READ DISBURSEMENT AMOUNT
    // ==========================================================

    const amount =
      Number(row[DISBURSEMENT_COL]) || 0;


    // ==========================================================
    // DO NOT SHOW ZERO-VALUE PUBLICATIONS
    // ==========================================================

    if (amount <= 0) {
      continue;
    }


    // ==========================================================
    // GET PUBLICATION INFORMATION
    // ==========================================================

    const publication =
      publicationByRow[i];


    if (!publication) {
      continue;
    }


    // ==========================================================
    // IMPACT FACTOR / QUARTILE
    //
    // SCOPUS papers should have these fields blank.
    // ==========================================================

    let impactFactor =
      publication.impactFactor;

    let quartile =
      publication.quartile;


    const editionUpper =
      publication.edition.toUpperCase();


    if (editionUpper.includes("SCOPUS")) {

      impactFactor = "";
      quartile = "";

    }


    // ==========================================================
    // THIS IS AN ELIGIBLE PUBLICATION
    // ==========================================================

    eligibleRows++;


    // ==========================================================
    // CHECK CONFIRMATION
    // ==========================================================

    const confirmation =
      String(row[CONFIRMATION_COL] || "")
        .trim()
        .toUpperCase();


    if (confirmation === "CONFIRMED") {

      confirmedRows++;


      // --------------------------------------------------------
      // Store confirmation information
      // --------------------------------------------------------

      if (!confirmationDate && row[DATE_COL]) {

        confirmationDate =
          Utilities.formatDate(
            new Date(row[DATE_COL]),
            Session.getScriptTimeZone(),
            "dd MMMM yyyy, hh:mm a"
          );

      }


      if (!confirmationEmail && row[CONFIRM_EMAIL_COL]) {

        confirmationEmail =
          String(row[CONFIRM_EMAIL_COL]);

      }

    }


    // ==========================================================
    // ADD PAPER
    // ==========================================================

    papers.push({

      master:
        publication.master,

      title:
        publication.title,

      edition:
        publication.edition,

      authorCount:
        publication.authorCount,

      impactFactor:
        impactFactor,

      quartile:
        quartile,

      amount:
        amount

    });

  }


  // ============================================================
  // NO ELIGIBLE PAPERS
  // ============================================================

  if (papers.length === 0) {

    return {

      authorFound: false,

      email: userEmail

    };

  }


  // ============================================================
  // TOTAL
  // ============================================================

  const total =
    papers.reduce(function(sum, paper) {

      return sum + paper.amount;

    }, 0);


  // ============================================================
  // CONFIRMATION STATUS
  // ============================================================

  const fullyConfirmed =
    eligibleRows > 0 &&
    confirmedRows === eligibleRows;


  // ============================================================
  // RETURN DATA
  // ============================================================

  return {

    authorFound: true,

    email: userEmail,

    authorName: authorName,

    ecode: ecode,

    papers: papers,

    total: total,

    confirmed: fullyConfirmed,

    confirmationDate: confirmationDate,

    confirmationEmail: confirmationEmail

  };

}
