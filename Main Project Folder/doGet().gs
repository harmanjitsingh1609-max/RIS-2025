function doGet(e) {

  const page =
    e &&
    e.parameter &&
    e.parameter.page
      ? e.parameter.page.toLowerCase()
      : "author";


  if (page === "admin") {

    return HtmlService
      .createHtmlOutputFromFile("AdminDashboard")
      .setTitle("RIS 2025 – Admin Confirmation Dashboard");

  }


  return HtmlService
    .createHtmlOutputFromFile("AuthorPortal")
    .setTitle("RIS 2025 – Publication Award Statement");

}
