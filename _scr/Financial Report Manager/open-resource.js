/**
 * Opens the folder with a redirect based on the name provided
 * @author William Veith <williamveith@gmail.com>
 * @param {string} folderName Name of the folder to open in a window redirect
 */
function openDigestFolder(folderName) {
  const folder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const popupName = folder.getName();
  const dataFolder = folderName === 'StatsResults' ?
    (() => {
      const parentFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
      return parentFolder.getFoldersByName('LabSentry').next().getFoldersByName('StatsResults').next();
    })() :
    folder.getParents().next().getFoldersByName(folderName).next();
  const htmlVariables = {
    url: dataFolder.getUrl(),
    windowTitle: `${[popupName]}: ${folderName}`,
  };
  openModalDialog(htmlVariables);
}

/**
 * Opens sheet containing all raw TMI data that requires manual reformatting
 * @author William Veith <williamveith@gmail.com>
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 */
function openSheet(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Fix These Rows');
  const spreadsheetId = spreadsheet.getId();
  const sheetId = sheet.getSheetId();
  const htmlVariables = {
    url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`,
    windowTitle: `Fix These Values`,
  };
  openModalDialog(htmlVariables);
}

/**
 * Opens NNCI Monthly User Reporting Tool word document from a Spreadsheet report
 * @author William Veith <williamveith@gmail.com>
 */
function openHelpDocuments() {
  const file = DriveApp.getFilesByName(`NNCI Monthly User Reporting Tool`).next();
  const htmlVariables = {
    url: `https://docs.google.com/document/d/${file.getId()}/edit`,
    windowTitle: file.getName(),
  };
  openModalDialog(htmlVariables);
}

/**
 * Open a modal box informing user selected resource is opening in a new window
 * @author William Veith <williamveith@gmail.com>
 * @param {object} htmlVariables
 * @param {string} templateName
 */
function openModalDialog(htmlVariables) {
  const htmlTemplate = HtmlService.createTemplateFromFile(`open`);
  htmlTemplate.htmlVariables = htmlVariables;
  const html = htmlTemplate.evaluate().getContent();
  SpreadsheetApp.getUi()
      .showModalDialog(
          HtmlService.createHtmlOutput(html).setHeight(1),
          `Opening... ${htmlVariables.windowTitle}`,
      );
}
