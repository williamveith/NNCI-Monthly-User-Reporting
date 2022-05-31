/** The number and 3 letter abbreviated formatted month names used in this files */
const monthColumns = {
  'Oct': 2,
  'Nov': 3,
  'Dec': 4,
  'Jan': 5,
  'Feb': 6,
  'Mar': 7,
  'Apr': 8,
  'May': 9,
  'Jun': 10,
  'Jul': 11,
  'Aug': 12,
  'Sep': 13,
};

/** The string number and 3 letter abbreviated formatted month names used in this files */
const monthNumber2Abbreviation = {
  '01': 'Jan',
  '02': 'Feb',
  '03': 'Mar',
  '04': 'Apr',
  '05': 'May',
  '06': 'Jun',
  '07': 'Jul',
  '08': 'Aug',
  '09': 'Sep',
  '10': 'Oct',
  '11': 'Nov',
  '12': 'Dec',
};

/** Index number for first spreadsheet cell in each data block */
const blockRow = [8, 22, 39, 48, 62, 79, 88, 102, 128, 142, 159, 170, 185];

/**
 * Structures a block of Stats Results rows that correspond with a block of rows in the financial report excel file
 * @author William Veith <williamveith@gmail.com>
 */
class Block {
  /**
   * Initializes the block
   * @param {number} startIndex Cell index of block's first row
   */
  constructor(startIndex) {
    this.start = startIndex;
    this.end = undefined;
  }
  /**
   * Splits Stats Results at new lines. Type casts string data to numbers
   * @param {string[]} blockArray Stats Results results data in string form
   * @return {number[]} Stats Results data in numeric form
   */
  data(blockArray) {
    const values = blockArray.slice(this.start, this.end);
    return values.map((value) => {
      return [parseFloat(value.slice(value.indexOf('\t') + 1))];
    });
  }
}

/**
 * Adds cumulative and monthly Stats Reports in the directory to the financial report spreadsheet
 * @author William Veith <williamveith@gmail.com>
 */
function addSummaryReports() {
  const activeSpreadSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parentFolder = DriveApp.getFileById(activeSpreadSheetId).getParents().next();
  const statsResultParentFolder = parentFolder
      .getFoldersByName('LabSentry').next()
      .getFoldersByName('StatsResults').next();
  const statsResultsFolders = [
    statsResultParentFolder.getFoldersByName('Cumulative Multistat Summary').next(),
    statsResultParentFolder.getFoldersByName('Monthly Multistat Summary').next(),
  ];
  statsResultsFolders.forEach((folder) => {
    const invoiceSummaryFiles = folder.getFiles();
    const invoiceType = folder.getName() === 'Cumulative Multistat Summary' ?
      'Cumulative' : 'Monthly';
    while (invoiceSummaryFiles.hasNext()) {
      const invoice = invoiceSummaryFiles.next();
      if (invoice.getDescription() !== 'Added To Report') {
        const month = monthNumber2Abbreviation[invoice.getName().slice(5, 7)];
        const fileId = invoice.getId();
        const contentArray = loadFile(fileId);
        setSheet(activeSpreadSheetId, contentArray, month, invoiceType);
        invoice.setDescription('Added To Report');
      }
    }
  });
}

/**
 * Loads Stats Result file and transforms it to a form that can be inserted into the financial report spreadsheet
 * @author William Veith <williamveith@gmail.com>
 * @param {string} fileId Stats Result file Id
 * @return {string[]} Stats Results reorganized and structured for insertion into a spreadsheet
 */
function loadFile(fileId) {
  const content = DriveApp.getFileById(fileId).getBlob().getDataAsString();
  let contentArray = content.split('\r\n');
  removeHeaders(contentArray);
  contentArray = removeIrrelevantData(contentArray);
  contentArray = breakIntoBlocks(contentArray);
  contentArray = removeIrrelevantBlocks(contentArray);
  contentArray = reorderBlocks(contentArray);
  return breakIntoSubBlocks(contentArray);
}

/**
 * Removes elements containing text file header information from the Stats Result array
 * @author William Veith <williamveith@gmail.com>
 * @param {string[]} contentArray Array contains arrays of Stat Results for each month
 * @param {number} numberOfHeaderRows Number of header rows in the Stats Result text file
 */
function removeHeaders(contentArray, numberOfHeaderRows = 7) {
  for (let i = 0; i < numberOfHeaderRows; i++) {
    contentArray.shift();
  }
}

/**
 * Removes elements containing Stats Result data not needed in the financial report
 * @author William Veith <williamveith@gmail.com>
 * @param {string[]} contentArray contains Stats Result as an array
 * @return {string[]} contains array of Stats Result data needed for the report
 */
function removeIrrelevantData(contentArray) {
  const irrelevantData = [
    'Standard Use\t', 'Train Use\t', 'Contract Use\t',
    'New Train Use\t', 'Trainer Use\t', 'Contractor Use\t',
    'All Codes Use\t', 'On-Site Use\t',
    'Total\t', '=', '*'];
  return contentArray.filter((row) => {
    if (!irrelevantData.some((irrelevant) => row.includes(irrelevant))) {
      return row;
    }
  });
}

/**
 * Separates Stats Result data, by category, into separate sections
 * @author William Veith <williamveith@gmail.com>
 * @param {string[]} contentArray contains arrays of Stat Results for each month
 * @return {string[]} contentArray contains arrays of Stat Results for each month
 */
function breakIntoBlocks(contentArray) {
  const blockArray = [];
  const blockNames = contentArray.filter((row) => {
    if (row.includes('------')) {
      return row;
    }
  });
  const numberOfBlocks = blockNames.length;
  for (let start = 0; start + 1 < numberOfBlocks; start++) {
    blockArray.push(contentArray.slice(contentArray.indexOf(blockNames[start]), contentArray.indexOf(blockNames[start + 1])));
  }
  return blockArray;
}

/**
 * Removes Stats Result data catagories no longer needed in the financial spreadsheet report
 * @author William Veith <williamveith@gmail.com>
 * @param {object[]} contentArray contains arrays of Stat Results for each month
 * @return {object[]} contains arrays of Stat Results for each month
 */
function removeIrrelevantBlocks(contentArray) {
  const irrelevantData = [
    'Contract Time', 'Training Fees', 'Standard Fees', 'Training Time',
    'Standard Time', 'Trained Users', 'Contract Users',
  ];
  return contentArray.filter((block) => {
    if (!irrelevantData.some((irrelevant) => block[0].includes(irrelevant))) {
      return block;
    }
  });
}

/**
 * Reorders Stats Result array, so array rows are ordered identically to financial report rows
 * @author William Veith <williamveith@gmail.com>
 * @param {string[]} contentArray contains arrays of Stat Results for each month
 * @return {string[]} contains arrays of Stat Results for each month
 */
function reorderBlocks(contentArray) {
  const desiredOrder = [
    '------ Lab Time ------',
    '------ Monthly Users ------',
    '------ Standard Users ------',
    '------ Fees ------',
    '------ New Users Trained on site (once per new users) ------',
  ];
  return desiredOrder.map((blockName) => {
    return contentArray.filter((block) => {
      if (block[0].includes(blockName)) {
        block.shift();
        return block;
      }
    })[0];
  });
}

/**
 * Places each Stats Result data category into its own block object
 * @author William Veith <williamveith@gmail.com>
 * @param {object[]} contentArray contains arrays of Stat Results for each month
 * @return {object[]} contains arrays of Stat Results for each month
 */
function breakIntoSubBlocks(contentArray) {
  const subBlockArray = [];
  const valueArray = [];
  const blockArray = contentArray.flat();
  blockArray.forEach((row, index) => {
    if (row.includes('(By Affiliation)') || row.includes('(By Discipline)')) {
      subBlockArray.length === 0 ? subBlockArray.push(new Block(index + 1)) : (function() {
        subBlockArray[subBlockArray.length - 1].end = index;
        subBlockArray.push(new Block(index + 1));
      })();
    }
    if (row.includes('Remote Use\t')) {
      subBlockArray.length === 0 ? subBlockArray.push(new Block(index)) : (function() {
        subBlockArray[subBlockArray.length - 1].end = index;
        subBlockArray.push(new Block(index));
      })();
    }
  });
  subBlockArray[subBlockArray.length - 1].end = blockArray.length;
  subBlockArray.forEach((block) => valueArray.push(block.data(blockArray)));
  return valueArray;
}

/**
 * Inserts Stats Result data into the financial report spreadsheet
 * @author William Veith <williamveith@gmail.com>
 * @param {string} activeSpreadSheetId
 * @param {string[]} contentArray Array containing Stats Result
 * @param {string} month 3 letter month abbreviation for the Stats Result
 * @param {string} invoiceType Cumulative or monthly Stats Result
 */
function setSheet(activeSpreadSheetId, contentArray, month, invoiceType) {
  const firstColumn = monthColumns[month];
  const numberOfColumns = 1;
  const spreadSheet = SpreadsheetApp.openById(activeSpreadSheetId);
  const sheet = spreadSheet.getSheetByName('Report');
  contentArray.forEach((block, index) => {
    const firstRow = blockRow[index];
    if (invoiceType === 'Cumulative') {
      if (firstRow === 88 || firstRow === 102) {
        const numberOfRows = block.length;
        sheet.getRange(firstRow, firstColumn, numberOfRows, numberOfColumns).setValues(block);
      }
    }
    if (invoiceType === 'Monthly') {
      if (firstRow != 88 && firstRow != 102) {
        const numberOfRows = block.length;
        sheet.getRange(firstRow, firstColumn, numberOfRows, numberOfColumns).setValues(block);
      }
    }
  });
}
