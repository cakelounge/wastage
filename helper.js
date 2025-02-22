

// Displays data for each suspect over the last 6 days.
function showData() {
  for (let j = 0; j < suspects.length; j++) {
    const suspect = suspects[j];

    for (let i = 0; i < 6; i++) {
      // Compute the sheet name for (today - i days)
      const parts = todaysSheetName.split(".");
      const d = new Date();
      d.setDate(d.getDate() - i);
      const day = d.getDate();
      const dateSheet = `${day}.${parts[1]}.${parts[2]}`;
      // console.log(dateSheet, rowData(suspect, dateSheet));

      if (i == 0) {
        const cakeName = getCellValue(dateSheet, "B" + suspect);
        history.innerHTML += `<span class="cakeName">${cakeName}</span><br>`;
      }

      history.innerHTML += rowData(suspect, dateSheet) + "<br>";
    }
    history.innerHTML += "<hr>";
  }
}