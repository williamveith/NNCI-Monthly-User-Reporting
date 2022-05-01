/** Directory structure for NNCI report */
const directoryMap = {
  name: 'Root',
  children: [
    {
      name: 'LabSentry',
      children: [
        {
          name: 'BillCodes',
          children: [],
        },
        {
          name: 'StatsResults',
          children: [
            {
              name: 'Monthly Multistat Summary',
              children: [],
            },
            {
              name: 'Cumulative Multistat Summary',
              children: [],
            },
          ],
        },
        {
          name: 'StatsXML',
          children: [],
        },
      ],
    },
    {
      name: 'TMI Data Digested',
      children: [],
    },
    {
      name: 'TMI Data Raw',
      children: [],
    },
  ],
};

/** 3 letter abbreviated month names */
const monthNames = ['Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'];

/** Financial report cell indices for month header rows */
const firstRowMonthHeaders = [7, 21, 38, 47, 61, 78, 87, 101, 118, 127, 141, 158, 169, 184];

/**
 * Creates files and directories for a new NNCI reporting year
 * @param {string} startYear Year report starts on
 * Ex. 2021-09_2022-10 report starts 2022
 */
function initializeYear(startYear) {
  const rootFolder = createFolders(startYear);
  folderStructure(rootFolder);
  const reportFile = initializeReport(rootFolder);
  setupReport(reportFile, startYear);
  openFinancialReport(reportFile);
}

/**
 * Creates the root folder for all NNCI files and data for a year
 * @param {string} startYear Year report starts on
 * @return {GoogleAppsScript.Drive.Folder} Root folder for year report and data
 */
function createFolders(startYear) {
  const parentFolder = DriveApp.getFileById(DocumentApp.getActiveDocument().getId()).getParents().next();
  const reportsFolder = parentFolder.getFoldersByName('Reports').next();
  return DriveApp.createFolder(`${startYear + 1}-09_${startYear}-10`)
      .moveTo(reportsFolder);
}

/**
 * Creates the directory structure for the NNCI report
 * @param {GoogleAppsScript.Drive.Folder} parentFolder Year NNCI Folder
 * @param {object} currentEntry Name of next directory to be created
 */
function folderStructure(parentFolder, currentEntry = directoryMap) {
  const currentFolder = currentEntry.name === 'Root' ?
    parentFolder : DriveApp.createFolder(currentEntry.name).moveTo(parentFolder);
  currentEntry.children.forEach((entry) => folderStructure(currentFolder, entry));
  if (currentEntry.children.length === 0) {
    return;
  }
}

/**
 * Creates new NNCI Report File
 * @param {GoogleAppsScript.Drive.Folder} rootFolder Year NNCI Folder
 * @return {GoogleAppsScript.Drive.File} Newly created NNCI Report File
 */
function initializeReport(rootFolder) {
  const parentFolder = DriveApp.getFileById(DocumentApp.getActiveDocument().getId()).getParents().next();
  const financialReportTemplate = parentFolder.getFoldersByName('Templates').next().getFilesByName('Template').next();
  return financialReportTemplate
      .makeCopy()
      .setName(rootFolder.getName())
      .moveTo(rootFolder);
}

/**
 * Updates elements of the new NNCI report to reflect the Reporting Year
 * @param {GoogleAppsScript.Drive.File} reportFile Newly created NNCI Report File
 * @param {number} startYear Year in 4 digit number form
 */
function setupReport(reportFile, startYear) {
  const firstColumnMonthHeaders = 2;
  const numberOfRowsMonthHeaders = 1;
  const year = parseInt(startYear.toString().slice(2, 4), 10);
  const sheet = SpreadsheetApp.openById(reportFile.getId())
      .getSheetByName(`Report`);
  sheet.getRange(1, 1).setValue([[`${sheet.getRange(1, 1).getValue()} 10/1/${startYear} - 09/30/${startYear + 1}`]]);
  const monthHeaderValues = monthNames.map((month, index) => {
    return index > 2 ? `${month}-${year + 1}` : `${month}-${year}`;
  });
  const numberOfColumns = monthHeaderValues.length;
  firstRowMonthHeaders.forEach((firstRow) => {
    sheet.getRange(firstRow, firstColumnMonthHeaders, numberOfRowsMonthHeaders, numberOfColumns)
        .setValues([monthHeaderValues]);
  });
}
