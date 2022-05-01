/**
 * Collects all digested TMI Spreadsheet files, turns them into Tab Separated Value text files, then downloads them to the local computer
 */
function downloadTMIDigests() {
  mainDigestTMIData(); // Makes sure all raw files are digested
  const tsvFiles = [];
  const tmiDigestFiles = DriveApp
      .getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
      .getParents().next()
      .getFoldersByName(`TMI Data Digested`).next()
      .getFiles();
  while (tmiDigestFiles.hasNext()) {
    const currentFile = tmiDigestFiles.next();
    tsvFiles.push(
        {
          name: `TMI_${currentFile.getName().slice(0, 7)}_bill-code.txt`,
          data: convertSheet2TSV(SpreadsheetApp.openById(currentFile.getId()).getSheets()[0]),
        });
  }
  downloadTSVFiles(tsvFiles);
}

/**
 * Concatenates an array of arrays together with new lines.
 * Delimits all values in the array with tabs
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Spreadsheet sheet containing digested TMI Data to download
 * @return {*[]} Digested TMI Data as a tab separated values array
 */
function convertSheet2TSV(sheet) {
  const data = sheet.getDataRange().getValues();
  return data.map((row) => row.join('\t')).join('\n');
}

/**
 * Turn array contains arrays of tab separated values into files, then download the files to the local computer
 * @param {*[]} tsvFiles Array of arrays containing digested TMI data in tab separated value form
 */
function downloadTSVFiles(tsvFiles) {
  const folder = (() => {
    const parentFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
    return parentFolder.getFoldersByName('LabSentry').next().getFoldersByName('BillCodes').next();
  })();
  const blobs = tsvFiles.map((file) => {
    const previousBillCodeVersion = folder.getFilesByName(file.name);
    previousBillCodeVersion.hasNext() ? previousBillCodeVersion.next().setTrashed(true) : undefined;
    return folder.createFile(file.name, file.data).getBlob();
  });
  const zipBlob = Utilities.zip(blobs, `${folder.getName()}.zip`);
  const fileId = DriveApp.createFile(zipBlob).getId();
  const htmlVariables = {
    zipFileUrl: `https://drive.google.com/uc?export=download&id=${fileId}`,
    title: zipBlob.getName(),
  };
  const htmlTemplate = HtmlService.createTemplateFromFile(`download-tmi-billcode`);
  htmlTemplate.htmlVariables = htmlVariables;
  const html = htmlTemplate
      .evaluate()
      .getContent();
  SpreadsheetApp.getUi()
      .showModalDialog(HtmlService.createHtmlOutput(html).setHeight(50), htmlVariables.title);
  DriveApp.getFileById(fileId).setTrashed(true);
}
