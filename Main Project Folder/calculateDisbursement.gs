function calculateDisbursement() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // ============================================================
  // COLUMN DEFINITIONS
  // ============================================================

  const COL = {
    MASTER: 0,        // A
    LINK: 1,          // B
    TITLE: 2,         // C
    EDITION: 3,       // D
    IMPACT: 4,        // E
    QUARTILE: 5,      // F
    AUTHOR: 6,        // G
    AFFILIATION: 7,   // H
    CATEGORY: 8,      // I
    ECODE: 9,         // J
    ROLE: 10,         // K
    REMARKS: 11       // L
  };

  // M = 13
  const startOutputCol = 13;

  const OUTPUT_HEADERS = [
    "Award Amount",
    "Total Authors",
    "TIET Student Count",
    "TIET Student Claimants",
    "TIET Faculty Claimants",
    "Per Share Amount",
    "Disbursement Amount",
    "Calculation Remark"
  ];

  // ============================================================
  // HELPER FUNCTIONS
  // ============================================================

  function clean(value) {
    return String(value || "").trim();
  }

  function normalizeCategory(value) {

    const cat = clean(value).toUpperCase();

    if (cat === "TIET") {
      return "TIET";
    }

    if (
      cat === "TIET_WOST" ||
      cat === "TIET WOST" ||
      cat === "TIET-WOST"
    ) {
      return "TIET_WOST";
    }

    if (cat === "STUDENT") {
      return "STUDENT";
    }

    if (
      cat.includes("OUTSIDE") ||
      cat.includes("OUSIDE")
    ) {
      return "OUTSIDE";
    }

    if (cat === "RETRACT") {
      return "RETRACT";
    }

    return cat;
  }

  function isTIET(category) {

    const cat = normalizeCategory(category);

    return (
      cat === "TIET" ||
      cat === "TIET_WOST" ||
      cat === "STUDENT"
    );
  }

  function isStudent(role) {

    const r = clean(role).toUpperCase();

    return (
      r.includes("STUDENT") ||
      r.includes("PHD") ||
      r.includes("UG") ||
      r.includes("PG")
    );
  }

  function isFaculty(role) {

    const r = clean(role).toUpperCase();

    return r.includes("FACULTY");
  }

  function hasClaim(ecode) {

    const value = clean(ecode).toUpperCase();

    if (!value) {
      return false;
    }

    if (
      value === "X" ||
      value === "N/A" ||
      value === "N.A." ||
      value === "NA"
    ) {
      return false;
    }

    return true;
  }

  // ============================================================
  // AWARD AMOUNT
  // ============================================================

  function getAwardAmount(row) {

    const edition =
      clean(row[COL.EDITION]).toUpperCase();

    const quartile =
      clean(row[COL.QUARTILE]).toUpperCase();

    // SCOPUS / ESCI
    if (
      edition.includes("SCOPUS") ||
      edition.includes("ESCI")
    ) {
      return 3000;
    }

    // WOS / SCIE / SCI / SSCI
    if (
      edition.includes("WOS") ||
      edition.includes("SCIE") ||
      edition.includes("SCI") ||
      edition.includes("SSCI")
    ) {

      if (quartile === "Q1") {
        return 20000;
      }

      if (quartile === "Q2") {
        return 15000;
      }

      if (quartile === "Q3") {
        return 8000;
      }

      if (quartile === "Q4") {
        return 5000;
      }
    }

    return 0;
  }

  // ============================================================
  // GROUP AUTHORS BY PUBLICATION
  // ============================================================

  const publications = [];

  let currentPublication = null;

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const master = clean(row[COL.MASTER]);

    // New publication
    if (master !== "") {

      currentPublication = {
        master: master,
        rows: [],
        startRow: i
      };

      publications.push(currentPublication);
    }

    // Add row to current publication
    if (currentPublication) {

      currentPublication.rows.push({
        rowIndex: i,
        data: row
      });
    }
  }

  // ============================================================
  // OUTPUT HEADERS
  // ============================================================

  sheet
    .getRange(
      1,
      startOutputCol,
      1,
      OUTPUT_HEADERS.length
    )
    .setValues([OUTPUT_HEADERS]);

  // ============================================================
  // CLEAR PREVIOUS OUTPUT
  // ============================================================

  if (data.length > 1) {

    sheet
      .getRange(
        2,
        startOutputCol,
        data.length - 1,
        OUTPUT_HEADERS.length
      )
      .clearContent();
  }

  // ============================================================
  // OUTPUT ARRAY
  // ============================================================

  const output = [];

  // ============================================================
  // PROCESS EACH PUBLICATION
  // ============================================================

  publications.forEach(pub => {

    // ----------------------------------------------------------
    // CREATE AUTHOR OBJECTS
    // ----------------------------------------------------------

    const authors = pub.rows.map(item => {

      const row = item.data;

      return {

        rowIndex: item.rowIndex,

        author: clean(row[COL.AUTHOR]),

        affiliation:
          clean(row[COL.AFFILIATION]),

        category:
          normalizeCategory(row[COL.CATEGORY]),

        ecode:
          clean(row[COL.ECODE]),

        role:
          clean(row[COL.ROLE]),

        claimed:
          hasClaim(row[COL.ECODE])
      };
    });

    if (authors.length === 0) {
      return;
    }

    // ==========================================================
    // PUBLICATION INFORMATION
    // ==========================================================

    const totalAuthors = authors.length;

    const firstAuthor = authors[0];

    const awardAmount =
      getAwardAmount(pub.rows[0].data);

    // ----------------------------------------------------------
    // ALL TIET AUTHORS
    // ----------------------------------------------------------

    const tietAuthors =
      authors.filter(author =>
        isTIET(author.category)
      );

    // ----------------------------------------------------------
    // TIET STUDENTS
    // ----------------------------------------------------------

    const tietStudents =
      tietAuthors.filter(author =>
        isStudent(author.role)
      );

    // ----------------------------------------------------------
    // REMAINING TIET AUTHORS
    // ----------------------------------------------------------

    const remainingTIETAuthors =
      tietAuthors.filter(author =>
        !isStudent(author.role)
      );

    // ----------------------------------------------------------
    // TIET FACULTY
    // ----------------------------------------------------------

    const tietFaculty =
      tietAuthors.filter(author =>
        isFaculty(author.role)
      );

    // ----------------------------------------------------------
    // CLAIMANTS
    // ----------------------------------------------------------

    const studentClaimants =
      tietStudents.filter(author =>
        author.claimed
      );

    const facultyClaimants =
      tietFaculty.filter(author =>
        author.claimed
      );

    // ==========================================================
    // FIRST AUTHOR TYPE
    // ==========================================================

    const firstAuthorIsOutside =
      normalizeCategory(firstAuthor.category) ===
      "OUTSIDE";

    const firstAuthorIsTIETStudent =
      isTIET(firstAuthor.category) &&
      isStudent(firstAuthor.role);

    const firstAuthorIsTIETFaculty =
      isTIET(firstAuthor.category) &&
      isFaculty(firstAuthor.role);

    // ==========================================================
    // PROCESS EVERY AUTHOR
    // ==========================================================

    authors.forEach(author => {

      let perShare = 0;
      let disbursement = 0;
      let calculationRemark = "";

      const category =
        normalizeCategory(author.category);

      // ========================================================
      // CASE 1: RETRACT
      // ========================================================

      if (category === "RETRACT") {

        perShare = 0;
        disbursement = 0;

        calculationRemark =
          "Retracted publication - no disbursement.";
      }

      // ========================================================
      // CASE 2: OUTSIDE TIET
      // ========================================================

      else if (category === "OUTSIDE") {

        perShare = 0;
        disbursement = 0;

        calculationRemark =
          "Outside TIET author - no TIET disbursement.";
      }

      // ========================================================
      // CASE 3: TIET AUTHOR
      // ========================================================

      else if (isTIET(category)) {

        // ------------------------------------------------------
        // 3A. FIRST AUTHOR IS OUTSIDE TIET
        // ------------------------------------------------------

        if (firstAuthorIsOutside) {

          perShare =
            totalAuthors > 0
              ? awardAmount / totalAuthors
              : 0;

          disbursement =
            author.claimed
              ? perShare
              : 0;

          calculationRemark =
            "1st author is outside TIET: award divided by total authors. Claimant receives own share; unclaimed shares are not redistributed.";
        }

        // ------------------------------------------------------
        // 3B. FIRST AUTHOR IS TIET STUDENT
        // ------------------------------------------------------

        else if (firstAuthorIsTIETStudent) {

          // 30% STUDENT POOL
          if (isStudent(author.role)) {

            if (tietStudents.length > 0) {

              perShare =
                (awardAmount * 0.30) /
                tietStudents.length;

            } else {

              perShare = 0;
            }

            disbursement =
              author.claimed
                ? perShare
                : 0;

            calculationRemark =
              "1st author is TIET student: 30% of award divided equally among all TIET student authors. Unclaimed shares are not redistributed.";
          }

          // 70% REMAINING TIET AUTHORS
          else if (isTIET(category)) {

            if (remainingTIETAuthors.length > 0) {

              perShare =
                (awardAmount * 0.70) /
                remainingTIETAuthors.length;

            } else {

              perShare = 0;
            }

            disbursement =
              author.claimed
                ? perShare
                : 0;

            calculationRemark =
              "1st author is TIET student: 70% of award divided equally among all remaining TIET authors. Unclaimed shares are not redistributed.";
          }
        }

        // ------------------------------------------------------
        // 3C. FIRST AUTHOR IS TIET FACULTY
        // ------------------------------------------------------

        else if (firstAuthorIsTIETFaculty) {

          perShare =
            tietAuthors.length > 0
              ? awardAmount / tietAuthors.length
              : 0;

          disbursement =
            author.claimed
              ? perShare
              : 0;

          calculationRemark =
            "1st author is TIET faculty: award divided equally among TIET authors only. Unclaimed shares are not redistributed.";
        }

        // ------------------------------------------------------
        // 3D. FALLBACK
        // ------------------------------------------------------

        else {

          perShare =
            tietAuthors.length > 0
              ? awardAmount / tietAuthors.length
              : 0;

          disbursement =
            author.claimed
              ? perShare
              : 0;

          calculationRemark =
            "Award divided equally among TIET authors. Unclaimed shares are not redistributed.";
        }
      }

      // ========================================================
      // CASE 4: UNKNOWN CATEGORY
      // ========================================================

      else {

        perShare = 0;
        disbursement = 0;

        calculationRemark =
          "Category not recognized - no disbursement.";
      }

      // ========================================================
      // ADD OUTPUT
      // ========================================================

      output.push([
        awardAmount,
        totalAuthors,
        tietStudents.length,
        studentClaimants.length,
        facultyClaimants.length,
        perShare,
        disbursement,
        calculationRemark
      ]);

    });
  });

  // ============================================================
  // WRITE OUTPUT
  // ============================================================

  if (output.length > 0) {

    sheet
      .getRange(
        2,
        startOutputCol,
        output.length,
        OUTPUT_HEADERS.length
      )
      .setValues(output);

    // Currency formatting
    sheet
      .getRange(
        2,
        startOutputCol,
        output.length,
        1
      )
      .setNumberFormat('₹#,##0.00');

    sheet
      .getRange(
        2,
        startOutputCol + 5,
        output.length,
        2
      )
      .setNumberFormat('₹#,##0.00');
  }

  // ============================================================
  // COMPLETION MESSAGE
  // ============================================================

  SpreadsheetApp.getUi().alert(
    "Disbursement calculation completed successfully."
  );
}
