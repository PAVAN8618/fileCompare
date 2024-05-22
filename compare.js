let highlightedWorkbook;
let newFileName;

function readExcelFile(file, callback) {
  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    callback(workbook);
  };
  reader.readAsArrayBuffer(file);
}

function compareSheets(oldSheet, newSheet) {
  const oldData = XLSX.utils.sheet_to_json(oldSheet, { header: 1 });
  const newData = XLSX.utils.sheet_to_json(newSheet, { header: 1 });
  const maxRows = Math.max(oldData.length, newData.length);
  const maxCols = Math.max(
    oldData[0] ? oldData[0].length : 0,
    newData[0] ? newData[0].length : 0
  );

  const table = document.createElement("table");
  const highlightRows = new Set();

  // Create table header
  const header = document.createElement("tr");
  for (let col = 0; col < maxCols; col++) {
    const th = document.createElement("th");
    th.textContent = `Column ${col + 1}`;
    header.appendChild(th);
  }
  table.appendChild(header);

  for (let row = 0; row < maxRows; row++) {
    const tr = document.createElement("tr");
    let rowDifferent = false;

    for (let col = 0; col < maxCols; col++) {
      const td = document.createElement("td");
      const oldValue = oldData[row] ? oldData[row][col] : undefined;
      const newValue = newData[row] ? newData[row][col] : undefined;

      td.textContent = newValue !== undefined ? newValue : "";
      if (oldValue !== newValue) {
        rowDifferent = true;
      }
      tr.appendChild(td);
    }

    if (rowDifferent) {
      tr.classList.add("highlight");
      highlightRows.add(row + 1);
    }

    table.appendChild(tr);
  }

  // Highlight rows in the workbook
  for (let row of highlightRows) {
    for (let col = 0; col < maxCols; col++) {
      const cellAddress = { c: col, r: row };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      if (!newSheet[cellRef]) newSheet[cellRef] = { t: "s", v: "" };
      if (!newSheet[cellRef].s) newSheet[cellRef].s = {};
      newSheet[cellRef].s.fill = { fgColor: { rgb: "FFFF00" } };
    }
  }

  return table;
}

function compareFiles() {
  const oldFile = document.getElementById("oldFile").files[0];
  const newFile = document.getElementById("newFile").files[0];
  const output = document.getElementById("output");

  if (!oldFile || !newFile) {
    alert("Please select both files.");
    return;
  }

  newFileName = newFile.name.replace(".xlsx", "_highlight.xlsx");

  readExcelFile(oldFile, (oldWorkbook) => {
    readExcelFile(newFile, (newWorkbook) => {
      output.innerHTML = "";
      highlightedWorkbook = XLSX.utils.book_new();

      oldWorkbook.SheetNames.forEach((sheetName) => {
        if (newWorkbook.SheetNames.includes(sheetName)) {
          const oldSheet = oldWorkbook.Sheets[sheetName];
          const newSheet = newWorkbook.Sheets[sheetName];

          const table = compareSheets(oldSheet, newSheet);
          const sheetTitle = document.createElement("h2");
          sheetTitle.textContent = `Sheet: ${sheetName}`;
          output.appendChild(sheetTitle);
          output.appendChild(table);

          XLSX.utils.book_append_sheet(
            highlightedWorkbook,
            newSheet,
            sheetName
          );
        }
      });
    });
  });
}

function exportHighlightedFile() {
  if (!highlightedWorkbook) {
    alert("No comparison done yet.");
    return;
  }
  XLSX.writeFile(highlightedWorkbook, newFileName);
}
