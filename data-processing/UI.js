/** Server side functions the program UI can trigger client side */
const reportDataFunctions = {
  openRawDataFolder: (function() {
    openDigestFolder(`TMI Data Raw`);
  }),
  openDigestFolder: (function() {
    openDigestFolder(`TMI Data Digested`);
  }),
  openSummaryFolder: (function() {
    openDigestFolder(`StatsResults`);
  }),
  sanitizeTMIRawData: (function() {
    mainSanitizeTMIRawData();
  }),
  digestTMIRawData: (function() {
    mainDigestTMIData();
  }),
  downloadTMIData: (function() {
    downloadTMIDigests();
  }),
  insertReports: (function() {
    addSummaryReports();
  }),
  help: (function() {
    openHelpDocuments();
  }),
};

/** Creates then serves the Financial Report Spreadsheet program UI */
function financialReportUi() {
  SpreadsheetApp.getUi().createMenu(`Financial Report Data`)
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Open')
          .addItem('Open Raw Data Folder', `FinancialReportManager.reportDataFunctions.openRawDataFolder`)
          .addItem('Open Digest Folder', `FinancialReportManager.reportDataFunctions.openDigestFolder`)
          .addItem('Open Stats Summary Folder', `FinancialReportManager.reportDataFunctions.openSummaryFolder`))
      .addSeparator()
      .addItem(`Sanitize Raw Data`, 'FinancialReportManager.reportDataFunctions.sanitizeTMIRawData')
      .addSeparator()
      .addItem(`Digest Raw Data`, 'FinancialReportManager.reportDataFunctions.digestTMIRawData')
      .addSeparator()
      .addItem(`Download Digest`, `FinancialReportManager.reportDataFunctions.downloadTMIData`)
      .addSeparator()
      .addItem(`Insert User Stat Summary Data`, `FinancialReportManager.reportDataFunctions.insertReports`)
      .addSeparator()
      .addItem(`Help`, `FinancialReportManager.reportDataFunctions.help`)
      .addToUi();
}
