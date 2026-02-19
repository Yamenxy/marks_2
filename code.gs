const SHEET_ID = "1gzTEmRi4CaOHvF8QdJoqwmE1PY6MJKkPkerAs6M7Yw0";
const SHEET_NAME = "Sheet1";

function doGet(e) {
  const action = e.parameter.action;

  if (action === "getData") {
    return getData();
  }
  if (action === "search") {
    const q = e.parameter.q || "";
    return searchData(q);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "invalid request" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function searchData(query) {
  query = (query || "").toString().trim();
  if (!query) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "empty query", rows: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const rows = data.slice(1);

  const normalize = v => (v == null ? "" : v.toString().trim());
  const isNumeric = s => s !== "" && !isNaN(Number(s));
  const numericEqual = (a, b) => isNumeric(a) && isNumeric(b) && Number(a) === Number(b);

  const idIndex = headers.findIndex(h => h && h.toString().trim().includes("رقم"));
  const nameIndex = headers.findIndex(h => h && h.toString().trim().includes("الاسم"));
  const codeIndex = headers.findIndex(h => h && (h.toString().trim().includes("كود") || h.toString().toLowerCase().includes("code") || h.toString().trim().includes("رمز")));

  const qLower = query.toString().toLowerCase();

  const matches = rows.filter(row => {
    const idCell = idIndex !== -1 ? normalize(row[idIndex]) : "";
    const codeCell = codeIndex !== -1 ? normalize(row[codeIndex]) : "";
    const nameCell = nameIndex !== -1 ? normalize(row[nameIndex]).toLowerCase() : "";

    // numeric-aware exact match for id or code
    if (idCell && (numericEqual(idCell, query) || idCell === query)) return true;
    if (codeCell && (numericEqual(codeCell, query) || codeCell === query)) return true;

    // partial name match
    if (nameCell && nameCell.indexOf(qLower) !== -1) return true;

    // fallback: any cell contains the query (partial, case-insensitive)
    for (let c = 0; c < row.length; c++) {
      const cell = normalize(row[c]).toLowerCase();
      if (cell.indexOf(qLower) !== -1) return true;
    }

    return false;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ headers, rows: matches }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === "updateMarks") {
    return updateMarks(data.studentId, data.marks);
  }
}

function getData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const rows = data.slice(1);

  return ContentService
    .createTextOutput(JSON.stringify({ headers, rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function updateMarks(studentId, marks) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const normalize = v => (v == null ? "" : v.toString().trim());
  const isNumeric = s => s !== "" && !isNaN(Number(s));
  const numericEqual = (a, b) => isNumeric(a) && isNumeric(b) && Number(a) === Number(b);

  const idIndex = headers.findIndex(h => h && h.toString().trim().includes("رقم"));

  for (let i = 1; i < data.length; i++) {
    const cellVal = idIndex !== -1 ? normalize(data[i][idIndex]) : "";
    const target = normalize(studentId);

    const found = (cellVal && (numericEqual(cellVal, target) || cellVal === target));
    if (found) {
      for (let key in marks) {
        // find column by exact trimmed header match (fallback to indexOf)
        const keyNorm = normalize(key);
        let colIndex = headers.findIndex(h => h && normalize(h) === keyNorm);
        if (colIndex === -1) colIndex = headers.indexOf(key);
        if (colIndex !== -1) {
          sheet.getRange(i + 1, colIndex + 1).setValue(marks[key]);
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "not found" }))
    .setMimeType(ContentService.MimeType.JSON);
}
