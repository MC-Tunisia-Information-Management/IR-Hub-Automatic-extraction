function doPost(data) {
  const sheet = SpreadsheetApp.openById("");
  const range = String(data.parameters.range);
  sheet.getRange(range).setValues(JSON.parse(data.parameters.data));
  return ContentService.createTextOutput("Updated data");
}

function monthToNumber(month) {
  let monthNumber;
  if (month === "July") {
    monthNumber = 0;
  } else if (month === "August") {
    monthNumber = 1;
  } else if (month === "September") {
    monthNumber = 2;
  } else if (month === "October") {
    monthNumber = 3;
  } else if (month === "November") {
    monthNumber = 4;
  } else if (month === "December") {
    monthNumber = 5;
  } else if (month === "January") {
    monthNumber = 6;
  } else if (month === "February") {
    monthNumber = 7;
  } else if (month === "March") {
    monthNumber = 8;
  } else if (month === "April") {
    monthNumber = 9;
  } else if (month === "Mai") {
    monthNumber = 10;
  } else if (month === "June") {
    monthNumber = 11;
  } else {
    monthNumber = "Invalid month"; // Handle any other invalid input
  }
  return monthNumber;
}

function remplissage() {
  // Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet named "Dashboard"
  const dashboardSheet = spreadsheet.getSheetByName("Dashboard");
  const chartsDataSheet = spreadsheet.getSheetByName("ChartsData");

  // Get the start and end month names from specific cells
  const startMonthName = dashboardSheet.getRange("G56").getValue();
  const endMonthName = dashboardSheet.getRange("L56").getValue();

  // Convert the month names to month numbers
  let startMonthNumber = monthToNumber(startMonthName);
  let endMonthNumber = monthToNumber(endMonthName);

  // Starting column letter and row
  let startColumn = "AG"; // Starting from column 'D'
  let row = 10; // The fixed row for output'

  for (let j = 0; j < 6; j++) {
    // Generate the formula using the converted month numbers
    const formula = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      47,
      5
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber = getColumnNumber(startColumn) + j;
    let columnLetter = getColumnLetter(columnNumber);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter + row).setFormula(formula);
  }

  let startColumn1 = "AG"; // Starting from column 'D'
  let row1 = 11; // The fixed row for output'

  for (let j = 0; j < 6; j++) {
    // Generate the formula using the converted month numbers
    const formula1 = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      61,
      5
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber1 = getColumnNumber(startColumn1) + j;
    let columnLetter1 = getColumnLetter(columnNumber1);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter1 + row1).setFormula(formula1);
  }

  let startColumn2 = "V"; // Starting from column 'D'
  let row2 = 3; // The fixed row for output'

  for (let j = 0; j < 7; j++) {
    // Generate the formula using the converted month numbers
    const formula2 = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      30,
      5
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber2 = getColumnNumber(startColumn2) + j;
    let columnLetter2 = getColumnLetter(columnNumber2);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter2 + row2).setFormula(formula2);
  }

  let startColumn3 = "W"; // Starting from column 'D'
  let row3 = 8; // The fixed row for output'

  for (let j = 0; j < 3; j++) {
    // Generate the formula using the converted month numbers
    const formula3 = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      58,
      5
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber3 = getColumnNumber(startColumn3) + j;
    let columnLetter3 = getColumnLetter(columnNumber3);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter3 + row3).setFormula(formula3);
  }

  let startColumn4 = "AA"; // Starting from column 'D'
  let row4 = 8; // The fixed row for output'

  for (let j = 0; j < 3; j++) {
    // Generate the formula using the converted month numbers
    const formula4 = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      72,
      5
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber4 = getColumnNumber(startColumn4) + j;
    let columnLetter4 = getColumnLetter(columnNumber4);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter4 + row4).setFormula(formula4);
  }

  let startColumn5 = "V"; // Starting from column 'D'
  let row5 = 13; // The fixed row for output'

  for (let j = 0; j < 7; j++) {
    // Generate the formula using the converted month numbers
    const formula5 = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      40,
      5
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber5 = getColumnNumber(startColumn5) + j;
    let columnLetter5 = getColumnLetter(columnNumber5);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter5 + row5).setFormula(formula5);
  }

  let startColumn6 = "V"; // Starting from column 'D'
  let row6 = 14;

  for (let j = 0; j < 7; j++) {
    // Generate the formula using the converted month numbers
    const formula6 = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      40,
      6
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber6 = getColumnNumber(startColumn6) + j;
    let columnLetter6 = getColumnLetter(columnNumber6);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter6 + row6).setFormula(formula6);
  }
  let row7 = 15; // The fixed row for output'

  for (let j = 0; j < 7; j++) {
    // Generate the formula using the converted month numbers
    const formula7 = formulasCreation(
      startMonthNumber,
      endMonthNumber,
      j,
      40,
      7
    );

    // Calculate the column number to place the formula, starting from 'D'
    let columnNumber5 = getColumnNumber(startColumn5) + j;
    let columnLetter5 = getColumnLetter(columnNumber5);

    // Set the generated formula in the dynamically chosen cell of the Dashboard sheet
    chartsDataSheet.getRange(columnLetter5 + row7).setFormula(formula7);
  }
}

function getColumnLetter(columnNumber) {
  let dividend = columnNumber;
  let columnName = "";
  let modulo;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = parseInt((dividend - modulo) / 26, 10); // Ensure base-10 parsing
  }

  return columnName;
}

function getColumnNumber(columnLetter) {
  // Convert the column letter to uppercase to handle case-insensitive input
  columnLetter = columnLetter.toUpperCase();

  // Initialize the column number
  let columnNumber = 0;

  // Iterate over each character of the column letter
  for (let i = 0; i < columnLetter.length; i++) {
    // Get the ASCII code of the character
    const charCode = columnLetter.charCodeAt(i);

    // Calculate the column number by adding the value of the character
    // minus the ASCII code of 'A' + 1 (to make 'A' = 1)
    columnNumber = columnNumber * 26 + (charCode - 64); // 64 is ASCII code of 'A'
  }

  return columnNumber;
}

function formulasCreation(startMonthNumber, endMonthNumber, j, x, y) {
  // Initialize the formula with the starting part
  let formula =
    "='MC DATA'! " + getColumnLetter(x + j) + (y + startMonthNumber * 40); // Starting formula

  // Loop through the months from startMonthNumber to endMonthNumber
  for (let i = startMonthNumber + 1; i <= endMonthNumber; i++) {
    const result = x + j;
    const columnLetter = getColumnLetter(result);
    formula += " + 'MC DATA'!" + columnLetter + (y + i * 40); // Ensured spacing for readability
  }

  return formula;
}
