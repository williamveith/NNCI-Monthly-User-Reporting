/**
 * Creates the financial report selector window
 * @author William Veith <williamveith@gmail.com>
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('open-selector');
  DocumentApp.getUi()
      .showModalDialog(
          HtmlService.createHtmlOutput(template.evaluate().getContent())
              .setHeight(500)
              .setWidth(700)
              .setTitle('NNCI Annual User Reports'),
          'Open...',
      );
}

/**
 * Retrieves information on all currently existing financial reports
 * @author William Veith <williamveith@gmail.com>
 * @return {object[]} Array of objects each containing the names & url of a Financial Report file
 */
function getFinancialReports() {
  const rootFolder = DriveApp.getFileById(DocumentApp.getActiveDocument().getId()).getParents().next();
  const reportFoldersIter = rootFolder.getFoldersByName('Reports').next().getFolders();
  const reports = [];
  while (reportFoldersIter.hasNext()) {
    const currentFolder = reportFoldersIter.next();
    const filesIter = currentFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (filesIter.hasNext()) {
      const currentFile = filesIter.next();
      const formattedName = currentFile.getName().replace(`-`, `/`).replace(`-`, `/`).replace(`_`, ` - `);
      reports.push({name: formattedName, url: currentFile.getUrl()});
    }
  }
  return reports;
}

/**
 * Creates and displays a modal dialog box so a user can select a financial report
 * @author William Veith <williamveith@gmail.com>
 * @param {object[]} reports Array of objects each containing the names & url of a Financial Report file
 */
function openReportSelection(reports) {
  const htmlTemplate = HtmlService.createTemplateFromFile('open-selector');
  htmlTemplate.reports = reports;
  const html = htmlTemplate.evaluate().getContent();
  DocumentApp.getUi()
      .showModalDialog(
          HtmlService.createHtmlOutput(html).setHeight(50),
          'Select Which open-report You Wish To View',
      );
}

/**
 * Opens selected financial report file
 * @author William Veith <williamveith@gmail.com>
 * @param {GoogleAppsScript.Drive.File} reportFile NNCI financial report file
 */
function openFinancialReport(reportFile) {
  reportFile.getUrl() === undefined ? noReportFound(reportFile) : (function() {
    const htmlVariables = {
      reportUrl: reportFile.getUrl(),
      reportTitle: `open-report: ${reportFile.getName()}`,
    };
    const htmlTemplate = HtmlService.createTemplateFromFile('open-report');
    htmlTemplate.htmlVariables = htmlVariables;
    const html = htmlTemplate.evaluate().getContent();
    DocumentApp.getUi()
        .showModalDialog(
            HtmlService.createHtmlOutput(html).setHeight(50),
            htmlVariables.reportTitle,
        );
  })();
}

/**
 * Alerts user that selector found no existing NNCI financial reports
 * @author William Veith <williamveith@gmail.com>
 * @param {GoogleAppsScript.Drive.File} reportFile If no report file exists
 */
function noReportFound(reportFile) {
  const ui = DocumentApp.getUi();
  ui.alert(`No report was found for ${reportFile.getName()}`);
}
