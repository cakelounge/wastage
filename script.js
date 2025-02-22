

let workbook;
let todaysSheetName;
const suspects = [];
const bestUsage = 5;
const result = document.querySelector('.result');
const infoDiv = document.querySelector('.info');
const history = document.querySelector('.history');

// Helper to get a cell value from a given sheet (or "" if not found)
function getCellValue(sheetName, cellRef) {
  const sheet = workbook.Sheets[sheetName];
  if (sheet && sheet[cellRef]) {
    return sheet[cellRef].v;
  }
  return "";
}

// Constructs today's sheet name based on the first sheet in the workbook.
// Assumes the first sheet’s name is in the form "day.month.year".
function getTodaysSheetName() {
  if (!workbook) return "";
  const sheetList = workbook.SheetNames;
  if (sheetList.length === 0) return "";
  // Use the first sheet as a template for month and year.
  const firstSheetName = sheetList[0];
  const dates = firstSheetName.split(".");
  const today = new Date();
  const day = today.getDate();
  return `${day}.${dates[1]}.${dates[2]}`;
}

// Sums values in a range of columns (given by start and end letters) for a specific row in a given sheet.
function rangeSum(start, end, row, sheetName) {
  let sum = 0;
  for (let c = start.charCodeAt(0); c <= end.charCodeAt(0); c++) {
    const col = String.fromCharCode(c);
    const cell = getCellValue(sheetName, col + row);
    if (cell === "" || cell === undefined) continue;
    const value = parseInt(cell, 10);
    if (!isNaN(value)) {
      sum += value;
    } else {
      console.log("Can't convert cell value into num for cell " + col + row);
    }
  }
  return sum;
}

// Formats the data for one row (suspect) on a given sheet (date)
function rowData(suspect, dateSheet) {
  const cakeName = getCellValue(dateSheet, "B" + suspect);
  const openingStr = getCellValue(dateSheet, "E" + suspect);
  let opening = parseInt(openingStr, 10);
  if (isNaN(opening)) {
    console.log("Can't convert opening into num for row " + suspect);
    opening = 0;
  }
  const stockIn = rangeSum('F', 'O', suspect, dateSheet);
  const stockOut = rangeSum('P', 'Y', suspect, dateSheet);
  let wastage = 0;
  const wastageStr = getCellValue(dateSheet, "Z" + suspect);
  if (wastageStr !== "" && wastageStr !== undefined) {
    wastage = parseInt(wastageStr, 10);
    if (isNaN(wastage)) {
      console.log("Can't convert wastage into num for row " + suspect);
      wastage = 0;
    }
  }
  const closingStr = getCellValue(dateSheet, "AA" + suspect);
  let closing = parseInt(closingStr, 10);
  if (isNaN(closing)) {
    console.log("Can't convert closing into num for row " + suspect);
    closing = 0;
  }

  colourNumber = (colour, number) => {
    if (number === 0) return number;
    return `<span style="color: ${colour};">${number}</span>`;
  }
  return `<span class="dateRow"> ${dateSheet} Opening: ${opening} , Stock In: ${colourNumber("green", stockIn)}
   ,Stock Out: ${colourNumber("red", stockOut)}, Wastage: ${colourNumber("red", wastage)}, Closing: ${closing}</span>`;
  // return `${dateSheet} Opening: ${opening}, Stock In: ${stockIn}, Stock Out: ${stockOut}, Wastage: ${wastage}, Closing: ${closing}`;
}


// Recursively checks past days to determine a suspect.
function checkLastDay(row, value, duration) {
  const cakeName = getCellValue(todaysSheetName, "B" + row);
  if (duration <= 0) {
    // console.log("suspect: ", cakeName);
    result.innerHTML += cakeName + "<br>";
    suspects.push(row);
    return;
  }
  const decrement = (bestUsage + 1) - duration;
  const today = new Date();
  const dateInt = today.getDate();
  const adjustedDay = dateInt - decrement;
  const parts = todaysSheetName.split(".");
  const lastDaySheet = `${adjustedDay}.${parts[1]}.${parts[2]}`;
  const cell = getCellValue(lastDaySheet, "AA" + row);
  const cellInt = parseInt(cell, 10);
  if (!isNaN(cellInt) && cellInt > 0) {
    checkLastDay(row, cellInt, duration - 1);
  }
}

// The main function that starts the application logic.
function startApp() {
  todaysSheetName = getTodaysSheetName();
  if (!todaysSheetName) {
    console.log("Workbook not loaded or invalid sheet name");
    return;
  }
  // Loop from row 2 to 114 in today's sheet.
  for (let i = 2; i <= 114; i++) {
    const cell = getCellValue(todaysSheetName, "AA" + i);
    const value = parseInt(cell, 10);
    if (!isNaN(value) && value > 0) {
      checkLastDay(i, value, bestUsage);
    }
  }

  infoDiv.innerHTML = "Check:\n"
  result.innerHTML = result.innerHTML;
  if (suspects.length > 0) {
    showData();
  }
}

// When a file is chosen, read it and start processing.
document.getElementById('fileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function (event) {
    const data = event.target.result;
    workbook = XLSX.read(data, { type: 'binary' });
    startApp();
  };
  reader.readAsBinaryString(file);
});