/**
 * Global variable that defines the digested excel sheet's column headers
 */
const tmiDigestedDataHeaders = [[
  'Date', 'User\'s full name (Last First)', 'User\'s EID',
  'Email address', 'Advisor (last, first) or Company', 'Department',
  'Account Number', 'TIME IN', 'TIME OUT',
  'Total Time (hrs.)', 'Charge ($)', 'System',
  'Code ( 2= training, 1= usage)',
]];

/**
 * Adds, removes, and checks flags located in a Google Drive file's description
 * @author William Veith <williamveith@gmail.com>
 */
class FileFlags {
  /**
   * Flag keys and values used across all scripts and files
   */
  constructor() {
    this.converted = 'A copy of this file has already been saved as a Google Sheet.';
    this.sanitized = 'This Google Sheet contains a sheet with sanitized raw data.';
    this.errors = 'Errors in the sanitized data need to be corrected before data can be digested';
    this.digested = 'A copy of the digested sanitized data has already been saved as a Google Sheet.';
  }
  /**
   * Checks if a specified file has a specified flag in its description
   * @param {GoogleAppsScript.Drive.File} file File to check description flag for
   * @param {string} searchFlag Flag to search for
   * @return {boolean} Whether file description contains flag
   */
  hasFlag(file, searchFlag) {
    const description = file.getDescription();
    return description === '' || description === null ? false : JSON.parse(description).includes(this[searchFlag]);
  }
  /**
   * Place specified flag inside the file description of the specified file
   * @param {GoogleAppsScript.Drive.File} file File located in Google Drive
   * @param {string} flag Flag to mark file with
   */
  setFlag(file, flag) {
    const currentDescription = file.getDescription();
    if (currentDescription === '' || currentDescription === null) {
      file.setDescription(JSON.stringify([this[flag]]));
    } else {
      const parsedCurrentDescription = JSON.parse(currentDescription);
      parsedCurrentDescription.push(this[flag]);
      file.setDescription(JSON.stringify(parsedCurrentDescription));
    }
  }

  /**
   * Remove specified flag inside the file description of the specified file
   * @param {GoogleAppsScript.Drive.File} file Class File https://developers.google.com/apps-script/reference/drive/file. File located in Google Drive
   * @param {string} flag Flag to remove from file
   */
  unsetFlag(file, flag) {
    const currentDescription = file.getDescription();
    if (currentDescription !== '' || currentDescription !== null) {
      const parsedCurrentDescription = JSON.parse(currentDescription);
      const newDescription = parsedCurrentDescription.filter((element) => element !== this[flag]);
      newDescription.length === 0 ? file.setDescription('') : file.setDescription(JSON.stringify(newDescription));
    }
  }
}

/**
 * Processes all raw, sanitized, TMI excel files located in the directory. Saves the digested TMI data to a google sheet file in the directory
 * @author William Veith <williamveith@gmail.com>
 */
function mainDigestTMIData() {
  const fileFlags = new FileFlags();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tmiDataFileIds = getRawTMIDataFiles();
  const rawSpreadsheets = getSpreadSheets(tmiDataFileIds);
  /**
 * @typedef sheetInfo
 * @type {object}
 * @property {string} spreadsheetId Spreadsheet file id
 * @property {string} spreadsheetName Future digested data spreadsheet file name
 * @property {string[]} headers New digested sheet headers
 * @property {*[]} rawData 2D array of undigested data
 */
  /** @type {sheetInfo[]} */
  const sheetInfos = rawSpreadsheets.map((spreadsheet) => {
    const sheet = spreadsheet.getSheets().filter((sheet) => {
      return sheet.getLastColumn() > 17;
    });
    return {spreadsheetId: spreadsheet.getId(), spreadsheetName: spreadsheet.getName().replace(`raw`, `digested`), headers: [], rawData: sheet[0].getDataRange().getValues()};
  });
  sheetInfos.forEach((sheetInfo) => sanitizeData(sheetInfo));
  sheetInfos.forEach((sheetInfo) => sortData(sheetInfo));
  sheetInfos.forEach((sheetInfo) => formatNames(sheetInfo));
  sheetInfos.forEach((sheetInfo) => addDictionaryValues(sheetInfo));
  sheetInfos.forEach((sheetInfo) => {
    if (SpreadsheetApp.openById(sheetInfo.spreadsheetId).getSheetByName('Fix These Rows') === null) {
      const digestSheet = saveDigestedData(sheetInfo, spreadsheet);
      setStyleFormatDigestedData(digestSheet);
      const file = DriveApp.getFileById(sheetInfo.spreadsheetId);
      fileFlags.setFlag(file, 'digested');
    } else {
      openSheet(SpreadsheetApp.openById(sheetInfo.spreadsheetId));
    }
  });
}

/**
 * Sanitized all raw TMI excel files located in the directory
 * @author William Veith <williamveith@gmail.com>
 */
function mainSanitizeTMIRawData() {
  const tmiDataFileIds = getRawTMIDataFiles();
  const rawSpreadsheets = getSpreadSheets(tmiDataFileIds);
  const sheetInfos = rawSpreadsheets.map((spreadsheet) => {
    const sheet = spreadsheet.getSheets().filter((sheet) => {
      return sheet.getLastColumn() > 17;
    });
    return {spreadsheetId: spreadsheet.getId(), spreadsheetName: spreadsheet.getName().replace(`raw`, `digested`), headers: [], rawData: sheet[0].getDataRange().getValues()};
  });
  sheetInfos.forEach((sheetInfo) => sanitizeData(sheetInfo));
}

/**
 * Retrieves all raw TMI files from the directory that have not already been digested
 * @author William Veith <williamveith@gmail.com>
 * @return {string[]} An array of file ids belonging to raw tmi data files
 */
function getRawTMIDataFiles() {
  const fileFlags = new FileFlags();
  const tmiDataFiles = DriveApp
      .getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
      .getParents().next()
      .getFoldersByName('TMI Data Raw').next()
      .getFiles();
  const fileIds = [];
  while (tmiDataFiles.hasNext()) {
    const currentFile = tmiDataFiles.next();
    fileFlags.hasFlag(currentFile, 'digested') ? undefined : fileIds.push(currentFile.getId());
  }
  return fileIds;
}

/**
 * Retrieves all undigested, raw TMI data spreadsheets. Converts any excel files into google sheets files if necessary
 * @author William Veith <williamveith@gmail.com>
 * @param {string[]} fileIds
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet[]} Array of Spreadsheet objects
 */
function getSpreadSheets(fileIds) {
  const fileFlags = new FileFlags();
  const fileIdsAndUndefineds = fileIds.map((fileId) => {
    const file = DriveApp.getFileById(fileId);
    switch (file.getMimeType()) {
      case MimeType.GOOGLE_SHEETS:
        return SpreadsheetApp.openById(fileId);
      case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
      case 'application/vnd.ms-excel':
        if (fileFlags.hasFlag(file, 'converted')) {
          return;
        }
        const blob = file.getBlob();
        const config = {
          title: file.getName(),
          parents: [{id: file.getParents().next().getId()}],
          description: '',
          mimeType: MimeType.GOOGLE_SHEETS,
        };
        const id = Drive.Files.insert(config, blob).id;
        fileFlags.setFlag(file, 'converted');
        return SpreadsheetApp.openById(id);
    }
  });
  return fileIdsAndUndefineds.filter((fileId) => fileId !== undefined);
}

/**
 * Removes raw TMI data headers, trailing/leading whitespace, blank rows, and formula errors. Makes data safe for the digest function
 * @author William Veith <williamveith@gmail.com>
 * @param {sheetInfo} sheetInfo TMI raw data spreadsheet properties
 */
function sanitizeData(sheetInfo) {
  const fileFlags = new FileFlags();
  const file = DriveApp.getFileById(sheetInfo.spreadsheetId);
  const spreadsheet = SpreadsheetApp.open(file);
  if (fileFlags.hasFlag(file, 'sanitized')) {
    sheetInfo.rawData = spreadsheet.getSheetByName('Sanitized Data').getDataRange().getValues();
    return;
  }
  const data = sheetInfo.rawData;
  data.shift(); /* Removes Extra Header Row      */
  data.shift(); /* Removes Raw Header            */
  const formulaErrorRegex = new RegExp(/^#(.*)[!?]$/gm);
  const whiteSpaceRegex = new RegExp(/^\s+|\s+$/gm);
  const dataWithoutFormulaErrors = data.filter((row) => {/* Removes Formula Errors        */
    if (!row.some((element) => formulaErrorRegex.test(element))) {
      return row;
    }
  });
  const sanitizedData = dataWithoutFormulaErrors.map((row) => {/* Removes Trailing White Spaces */
    return row.map((element) => {
      return element.toString().replace(whiteSpaceRegex, '');
    });
  });
  const sheet = spreadsheet.insertSheet('Sanitized Data');
  sheet.getRange(1, 1, sanitizedData.length, sanitizedData[0].length).setValues(sanitizedData);
  fileFlags.setFlag(file, 'sanitized');
  sheetInfo.rawData = sanitizedData;
}

/**
 * Rearranges raw TMI data so it works with the LabAccess Database
 * @author William Veith <williamveith@gmail.com>
 * @param {sheetInfo} sheetInfo TMI raw data spreadsheet properties
 */
function sortData(sheetInfo) {
  const data = sheetInfo.rawData;
  const isTimeInOutRegex = new RegExp(/^Time: \d{2}:\d{2} [P|A]{1}M{1} - \d{2}:\d{2} [P|A]{1}M{1}/gm);
  const timeInRegex = new RegExp(/(?<=Time: )\d{2}:\d{2} [P|A]{1}M{1}/gm);
  const timeOutRegex = new RegExp(/(?<=Time: \d{2}:\d{2} [P|A]{1}M{1} - )\d{2}:\d{2} [P|A]{1}M{1}/gm);
  const doorOpenRegex = new RegExp(/(?<=Door Opened at )\d{2}:\d{2}/gm);
  const sorted = {
    data: data.map((row) => {
      const trainingCode = row[10].toLowerCase().includes('training') ? `2` : `1`;
      const time = row[10].match(doorOpenRegex) !== null ?
        {'TIME IN': row[10].match(doorOpenRegex), 'TIME OUT': row[10].match(doorOpenRegex)} :
        row[10].match(isTimeInOutRegex) !== null ?
          {'TIME IN': row[10].match(timeInRegex), 'TIME OUT': row[10].match(timeOutRegex)} :
          {'TIME IN': ['12:00 PM'], 'TIME OUT': ['12:00 PM']};
      return {
        'Date': row[11],
        'User\'s full name (Last First)': row[5],
        'User\'s EID': 'na',
        'Email address': '',
        'Advisor (last, first) or Company': row[4],
        'Department': '',
        'Account Number': row[1],
        'TIME IN': time['TIME IN'][0],
        'TIME OUT': time['TIME OUT'][0],
        'Total Time (hrs.)': row[6],
        'Charge ($)': row[13],
        'System': row[8],
        'Code ( 2= training, 1= usage)': trainingCode,
      };
    }),
  };
  sheetInfo.rawData = sorted.data;
}

/**
 * Modifies and restructures TMI raw data names to meet LabAccess database format
 * @author William Veith <williamveith@gmail.com>
 * @param {sheetInfo} sheetInfo TMI raw data spreadsheet properties
 */
function formatNames(sheetInfo) {
  const data = sheetInfo.rawData;
  const typeSetDataArray = setRawDataTypeAndCase(data);
  const nameKeysArray = ['User\'s full name (Last First)', 'Advisor (last, first) or Company'];
  const rowsWithErrors = [];
  nameKeysArray.forEach((nameKey) => {
    rowsWithErrors.push(splitTransformConcatNames(typeSetDataArray, nameKey));
  });
  const sheetValues = rowsWithErrors[0].filter((rowsByNameKey) => rowsByNameKey.length !== 0);
  const spreadsheet = SpreadsheetApp.openById(sheetInfo.spreadsheetId);
  if (spreadsheet.getSheetByName(`Fix These Rows`) === null && sheetValues.length > 0) {
    const sheet = spreadsheet.insertSheet(`Fix These Rows`);
    const headers = ['Fix', 'Row Number', ...tmiDigestedDataHeaders[0]];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, sheetValues.length, sheetValues[0].length).setValues(sheetValues);
    setStyleFormatMisshapenData(sheet);
    sheetInfo.rawData = typeSetDataArray;
  }
  if (spreadsheet.getSheetByName(`Fix These Rows`) !== null && sheetValues.length === 0) {
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(`Fix These Rows`));
  }
}

/**
 * Turns all names into lowercase to prevent errors
 * @author William Veith <williamveith@gmail.com>
 * @param {*[]} data  Google sheet header data
 * @return {*[]} Google sheet header data correctly formatted
 */
function setRawDataTypeAndCase(data) {
  return data.map((row) => {
    row['User\'s full name (Last First)'] = row['User\'s full name (Last First)'].toLowerCase();
    row['Advisor (last, first) or Company'] = row['Advisor (last, first) or Company'].toLowerCase();
    return row;
  });
}

/**
 * Reorders and reforms names from raw TMI data. Returns names that could not be fixed
 * @author William Veith <williamveith@gmail.com>
 * @param {*[]} typeSetDataArray 2D array of Google Sheet data typeset
 * @param {string[]} nameKey An array of header to split data at
 * @return {*[]} 2D array with google sheet data rows that need fixing manually
 */
function splitTransformConcatNames(typeSetDataArray, nameKey) {
  const recordsToFix = [];
  const splitNamesAtPeriodDashMultinameRegex = new RegExp(/\b[\w][a-zA-Z]+\b|[a-zA-Z](?=\.)/gm);
  const getSpaceDashPeriodRegex = new RegExp(/[\s-.]/gm);
  typeSetDataArray.forEach((row, rowIndex) => {
    try {
      let splitName = row[nameKey].match(splitNamesAtPeriodDashMultinameRegex);
      splitName = splitName.map((name) => {
        return name.length === 1 ? name.toUpperCase() : `${name.charAt(0).toUpperCase()}${name.slice(1)}`;
      });
      const splitDelimiterArray = row[nameKey].match(getSpaceDashPeriodRegex);
      if (splitDelimiterArray.every((delim) => delim === ' ')) {
        const lastName = splitName.pop();
        splitName.unshift(`${lastName},`);
        typeSetDataArray[rowIndex][nameKey] = splitName.join(' ');
      } else {
        splitDelimiterArray.forEach((delimit, delimIndex) => {
          if (delimit === `-`) {
            splitName.splice(delimIndex, 2, `${splitName[delimIndex]}-${splitName[delimIndex + 1]}`);
          }
          if (row[nameKey].match(/,/gm)) {
            splitName.splice(0, 1, `${splitName[0]},`);
          }
        });
        const last = splitName.pop();
        splitName.unshift(`${last},`);
        let finalCheck = splitName.join(' ');
        while (finalCheck.charAt(finalCheck.length - 1) === ',') {
          finalCheck = finalCheck.slice(0, -1);
        }
        typeSetDataArray[rowIndex][nameKey] = finalCheck;
      }
    } catch (e) {
      const rowWithError = Object.values(row);
      const errorToFix = row[nameKey];
      recordsToFix.push([errorToFix, rowIndex + 3, ...rowWithError]);
      typeSetDataArray.splice([rowIndex], 1);
    }
  });
  return recordsToFix;
}

/**
 * Adds Professor research field and student EIDs from dictionary files
 * @author William Veith <williamveith@gmail.com>
 * @param {sheetInfo} sheetInfo TMI raw data spreadsheet properties
 */
function addDictionaryValues(sheetInfo) {
  const eidDictionaryId = '1rXYosSuFLYDu0BRx51_bR2TsJHlISTpE';
  const professorDictionaryId = '1ZzhNvDSKZp7hwm6MWg2iRLNndywv04HZ';
  const eidDictionary = JSON.parse(
      DriveApp
          .getFileById(eidDictionaryId)
          .getBlob()
          .getDataAsString(),
  );
  const professorDictionary = JSON.parse(
      DriveApp
          .getFileById(professorDictionaryId)
          .getBlob()
          .getDataAsString(),
  );
  sheetInfo.rawData.forEach((row, rowIndex) => {
    const eidKey = row['User\'s full name (Last First)'].toLowerCase();
    const professorKey = row['Advisor (last, first) or Company'].toLowerCase();
    if (eidDictionary.hasOwnProperty(eidKey) && row['User\'s EID'] === 'na') {
      sheetInfo.rawData[rowIndex]['User\'s EID'] = eidDictionary[eidKey];
    }
    if (professorDictionary.hasOwnProperty(professorKey)) {
      sheetInfo.rawData[rowIndex]['Department'] = professorDictionary[professorKey];
    }
  });
}

/**
 * Saves digested Raw TMI data to a new spreadsheet
 * @author William Veith <williamveith@gmail.com>
 * @param {sheetInfo} sheetInfo
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} reportSpreadsheet Digested data to be saved
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Google Sheet object containing digested TMI data
 */
function saveDigestedData(sheetInfo, reportSpreadsheet) {
  const digestFolder = DriveApp
      .getFileById(reportSpreadsheet.getId())
      .getParents().next()
      .getFoldersByName(`TMI Data Digested`).next();
  const digestedData = sheetInfo.rawData.map((object) => Object.values(object));
  const digestSpreadsheet = SpreadsheetApp.create(sheetInfo.spreadsheetName);
  DriveApp.getFileById(digestSpreadsheet.getId()).moveTo(digestFolder);
  digestSpreadsheet.getSheets()[0].activate();
  const sheet = digestSpreadsheet.getActiveSheet();
  const headerRange = sheet.getRange(1, 1, 1, tmiDigestedDataHeaders[0].length);
  const dataRange = sheet.getRange(2, 1, digestedData.length, digestedData[0].length);
  headerRange.setValues(tmiDigestedDataHeaders);
  dataRange.setValues(digestedData);
  const sheetName = sheetInfo.spreadsheetName.replace('_digested', '');
  sheet.setName(sheetName);
  return sheet;
}

/**
 * Styles the new sheet containing the digested TMI data
 * @author William Veith <williamveith@gmail.com>
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sheet Google Sheet object
 */
function setStyleFormatDigestedData(sheet) {
  const rows = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  const columns = sheet.getLastColumn();
  const maxColumns = sheet.getMaxColumns();
  const firstColumnRange = sheet.getRange(1, 1, rows, 1);
  const timeInOutColumnRange = sheet.getRange(1, 8, rows, 2);
  const headerRowRange = sheet.getRange(1, 1, 1, columns);
  sheet.getDataRange().setFontFamily('Times New Roman').setFontSize(12);
  headerRowRange.setNumberFormat('@STRING@').setFontWeight('bold');
  firstColumnRange.setNumberFormat('@STRING@');
  timeInOutColumnRange.setNumberFormat('@STRING@');
  sheet.autoResizeColumns(1, columns);
  maxRows > rows ? sheet.deleteRows(rows + 1, (sheet.getMaxRows() - 1) - rows) : undefined;
  maxColumns > columns ? sheet.deleteColumns(columns + 1, (sheet.getMaxColumns() - 1) - columns) : undefined;
  sheet.deleteColumns(columns + 1, sheet.getMaxColumns() - columns);
  sheet.protect().setDescription('Only Edit With Script');
}

/**
 * Styles the sheet containing the TMI data that needs to be manually reformated
 * @author William Veith <williamveith@gmail.com>
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sheet Google Sheet object
 */
function setStyleFormatMisshapenData(sheet) {
  const rows = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  const columns = sheet.getLastColumn();
  const maxColumns = sheet.getMaxColumns();

  const dataRange = sheet.getDataRange();
  const header = sheet.getRange(1, 1, 1, columns);
  const charges = sheet.getRange(1, 13, rows, 1);
  const protectedData = sheet.getRange(1, 2, rows, columns - 1);

  dataRange.setFontFamily('Times New Roman').setFontSize(12).setNumberFormat('@STRING@');
  header.setFontWeight('bold');
  charges.setNumberFormat('$#,##0.00;$(#,##0.00)');

  sheet.autoResizeColumns(1, columns);
  maxRows > rows ? sheet.deleteRows(rows + 1, (sheet.getMaxRows() - 1) - rows) : undefined;
  maxColumns > columns ? sheet.deleteColumns(columns + 1, (sheet.getMaxColumns() - 1) - columns) : undefined;
  sheet.deleteColumns(columns + 1, sheet.getMaxColumns() - columns);
  header.protect().setDescription(`The headers do not need to be edited`);
  protectedData.protect().setDescription('These values only reference the record with the incorrect data. Edit the values in the first column');
}
