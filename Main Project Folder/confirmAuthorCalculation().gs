function confirmAuthorCalculation() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const data =
    sheet.getDataRange().getValues();

  // Authenticated Thapar account
  const userEmail =
    Session.getActiveUser().getEmail()
      .trim()
      .toLowerCase();

  // Column indexes (zero-based)
  const DISBURSEMENT_COL = 18;  // S
  const EMAIL_COL = 20;         // U
  const CONFIRMATION_COL = 21;  // V
  const DATE_COL = 22;          // W
  const CONFIRM_EMAIL_COL = 23; // X

  const confirmationTime = new Date();

  let updatedRows = 0;

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const rowEmail =
      String(row[EMAIL_COL] || "")
        .trim()
        .toLowerCase();

    // Only process the logged-in author's records
    if (rowEmail !== userEmail) {
      continue;
    }

    // Only confirm publications having a positive
    // disbursement amount
    const amount =
      Number(row[DISBURSEMENT_COL]) || 0;

    if (amount <= 0) {
      continue;
    }

    // Write confirmation information
    sheet.getRange(i + 1, CONFIRMATION_COL + 1)
      .setValue("CONFIRMED");

    sheet.getRange(i + 1, DATE_COL + 1)
      .setValue(confirmationTime);

    sheet.getRange(i + 1, CONFIRM_EMAIL_COL + 1)
      .setValue(userEmail);

    updatedRows++;
  }

  if (updatedRows === 0) {

    return {
      success: false,
      message:
        "No eligible publication records were found for confirmation."
    };

  }

  return {
    success: true,
    message:
      "Your RIS 2025 calculation has been confirmed successfully.",
    updatedRows: updatedRows
  };
}
