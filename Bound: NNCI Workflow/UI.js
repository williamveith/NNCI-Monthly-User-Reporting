/**
 * Creates Financial Report Tab in NNCI Monthly User Reporting Tool on file open
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Financial Report')
      .addItem('Create New Report', 'initializeYearMenuFunction')
      .addSeparator()
      .addItem('Open Financial Report', 'doGet')
      .addToUi();
}

/**
 * Adds UI button functionality to "Create New Report" located in tab "Financial Report"
 */
function initializeYearMenuFunction() {
  const ui = DocumentApp.getUi();
  const response = ui.prompt('Enter the year this report starts on', ui.ButtonSet.OK_CANCEL);
  const startYear = validateYearInput(response.getResponseText());
  switch (response.getSelectedButton()) {
    case ui.Button.OK:
      initializeYear(startYear);
      break;
    case ui.Button.CANCEL:
      break;
  }
}

/**
 * Typecasts Year into a number. Verifies year in tbe starting year
 * @param {string} response Text response from user input field
 * @return {number} Year in 4 digit number form
 */
function validateYearInput(response) {
  const currentYear = new Date().getFullYear();
  try {
    const inputYear = parseInt(response, 10);
    if (Math.abs(inputYear - currentYear) < 11) {
      return inputYear;
    } else {
      return new Date().getMonth() > 7 ? currentYear : currentYear - 1;
    }
  } catch {
    return currentYear;
  }
}
