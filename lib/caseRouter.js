// lib/caseRouter.js
//
// All the "business logic" for this tool lives here, separate from the
// screen (app/page.js). If the layout of either spreadsheet ever changes,
// this is almost always the only file you need to touch.
//
// Everything runs in the browser (client-side). Nothing here uploads your
// files to a server -- the xlsx library just reads bytes you already have
// in memory from the <input type="file"> element.

// Monthly tabs in the Case Tracker are named like "2026.07". This matches
// that pattern so we can find them automatically instead of hard-coding
// every possible tab name.
export const MONTH_TAB_REGEX = /^\d{4}\.\d{2}$/;

// The tab we read in the Caseload file. Change this if that tab ever gets
// renamed.
export const CASELOAD_SHEET_NAME = "Caseload";

// Turns "  Sign   Date \n" into "sign date" so header comparisons don't
// break because of stray spaces, line breaks, or capitalization -- which
// is exactly what real-world Excel headers tend to have.
function normalize(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// -----------------------------------------------------------------------
// STEP 1: Read the Case Tracker file (File 1) in full.
// It's small (well under a MB), so a full parse is fast and safe.
// -----------------------------------------------------------------------
// A tab "has data" if, after its header row, at least one row has any
// non-empty cell. Some months (e.g. a future month that hasn't started
// yet) exist as empty placeholder tabs -- we don't want those picked as
// the default.
function tabHasData(XLSX, workbook, tabName) {
  const sheet = workbook.Sheets[tabName];
  if (!sheet) return false;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headerRowIdx = rows.findIndex((row) =>
    row.some((cell) => normalize(cell) === "state id")
  );
  if (headerRowIdx === -1) return false;
  return rows.slice(headerRowIdx + 1).some((row) => row.some((cell) => cell !== ""));
}

export async function readCaseTrackerFile(file) {
  const XLSX = await import("xlsx");
  const buf = await file.arrayBuffer();
  const workbook = XLSX.read(buf, { type: "array", cellDates: true });

  const monthlyTabs = workbook.SheetNames.filter((name) =>
    MONTH_TAB_REGEX.test(name)
  ).sort(); // "YYYY.MM" strings sort correctly as plain text

  if (monthlyTabs.length === 0) {
    throw new Error(
      "No monthly tabs (named like \"2026.07\") were found in that file."
    );
  }

  // Default to the most recent tab that actually has rows in it, skipping
  // empty placeholder tabs for months that haven't started yet.
  const tabsWithData = monthlyTabs.filter((t) => tabHasData(XLSX, workbook, t));
  const latestTab = tabsWithData.length > 0
    ? tabsWithData[tabsWithData.length - 1]
    : monthlyTabs[monthlyTabs.length - 1];

  return { workbook, monthlyTabs, latestTab };
}

// -----------------------------------------------------------------------
// STEP 2: Read ONLY the "Caseload" tab from File 2.
//
// HCGY_Caseload.xlsx is ~15MB because two other tabs ("Data" and
// "Data 2") have formatting applied to over a million empty rows each.
// Passing `sheets: [CASELOAD_SHEET_NAME]` tells the xlsx library to skip
// parsing those tabs entirely, which keeps this fast and avoids ever
// touching (or risking corrupting) the rest of that workbook.
// -----------------------------------------------------------------------
export async function readCaseloadSheet(file) {
  const XLSX = await import("xlsx");
  const buf = await file.arrayBuffer();
  const workbook = XLSX.read(buf, {
    type: "array",
    cellDates: true,
    sheets: [CASELOAD_SHEET_NAME],
  });

  const sheet = workbook.Sheets[CASELOAD_SHEET_NAME];
  if (!sheet) {
    throw new Error(`Couldn't find a tab named "${CASELOAD_SHEET_NAME}" in that file.`);
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const header = rows[0] || [];
  const dataRows = rows.slice(1).filter((row) => row.some((cell) => cell !== ""));

  const stateIdColIdx = header.findIndex((h) => normalize(h) === "state id");
  const existingStateIds = new Set(
    dataRows
      .map((row) => (stateIdColIdx >= 0 ? String(row[stateIdColIdx]).trim() : ""))
      .filter(Boolean)
  );

  return { header, existingStateIds };
}

// -----------------------------------------------------------------------
// Field mapping between the two files.
//
// Left side  = normalized header text we look for in the Case Tracker's
//              monthly tabs.
// Right side = a matcher function that identifies the matching column in
//              the Caseload tab's header row, however it's punctuated.
//
// Any Caseload column with no matcher listed here (App ID, Exit Date,
// PWE Placed, etc.) is simply left blank on new rows, per your call.
// -----------------------------------------------------------------------
const FIELD_MAP = [
  { key: "stateId", label: "State ID", sourceHeader: "state id", targetMatch: (h) => h === "state id" },
  { key: "lastName", label: "Last Name", sourceHeader: "last name", targetMatch: (h) => h === "last name" },
  { key: "firstName", label: "First Name", sourceHeader: "first name", targetMatch: (h) => h === "first name" },
  { key: "office", label: "Office", sourceHeader: "office", targetMatch: (h) => h === "office" },
  { key: "caseManager", label: "Case Manager", sourceHeader: "case manager", targetMatch: (h) => h.startsWith("case manager from pwe") },
  { key: "assignDate", label: "Assign Date", sourceHeader: "assign date", targetMatch: (h) => h === "date assigned" },
  { key: "signDue", label: "Sign Due", sourceHeader: "sign due", targetMatch: (h) => h === "signature due" },
  { key: "signDate", label: "Sign Date", sourceHeader: "sign date", targetMatch: (h) => h === "sign date" },
  { key: "status", label: "Status", sourceHeader: "status", targetMatch: (h) => h === "status" },
  { key: "notes", label: "Notes", sourceHeader: "notes", targetMatch: (h) => h === "notes" },
];

// Fields used only for filtering / preview, not written to the output file
// (the Caseload tab has no "Fund" column to receive it).
const FUND_SOURCE_HEADER = "fund";

// Given the Caseload tab's header row, figure out which column index each
// FIELD_MAP entry should be written to.
function buildTargetColumnMap(targetHeader) {
  const normalizedTargetHeader = targetHeader.map(normalize);
  const map = {};
  for (const field of FIELD_MAP) {
    const idx = normalizedTargetHeader.findIndex((h) => field.targetMatch(h));
    map[field.key] = idx; // -1 if not found
  }
  return map;
}

// -----------------------------------------------------------------------
// STEP 3: For each selected monthly tab, find rows where
//   Status (column O) = "APVD"
//   Fund   (column M) contains "OSY" or "ISY" (covers "A/OSY" etc. too)
// and build the mapped output row + preview info for each match.
// -----------------------------------------------------------------------
export async function filterAndMapRows({ workbook, tabNames, caseloadHeader, existingStateIds }) {
  const XLSX = await import("xlsx");
  const targetColMap = buildTargetColumnMap(caseloadHeader);
  const missingTargets = FIELD_MAP.filter((f) => targetColMap[f.key] === -1).map((f) => f.label);

  const results = [];

  for (const tabName of tabNames) {
    const sheet = workbook.Sheets[tabName];
    if (!sheet) continue;

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    // Find the header row (the row that actually contains "State ID"),
    // rather than assuming it's always row 1 -- this tab layout has a
    // blank spacer row under the header in some months.
    const headerRowIdx = rows.findIndex((row) =>
      row.some((cell) => normalize(cell) === "state id")
    );
    if (headerRowIdx === -1) continue;

    const header = rows[headerRowIdx].map(normalize);
    const colIdx = (h) => header.indexOf(h);
    const statusIdx = colIdx("status");
    const fundIdx = colIdx(FUND_SOURCE_HEADER);

    const dataRows = rows
      .slice(headerRowIdx + 1)
      .filter((row) => row.some((cell) => cell !== ""));

    for (const row of dataRows) {
      const status = normalize(row[statusIdx]);
      const fund = String(row[fundIdx] ?? "");

      const statusMatches = status === "apvd";
      const fundMatches = /osy|isy/i.test(fund);
      if (!statusMatches || !fundMatches) continue;

      // Build the output row, aligned to the Caseload tab's real columns.
      const outputRow = new Array(caseloadHeader.length).fill("");
      for (const field of FIELD_MAP) {
        const targetIdx = targetColMap[field.key];
        if (targetIdx === -1) continue;
        const sourceIdx = colIdx(field.sourceHeader);
        outputRow[targetIdx] = sourceIdx === -1 ? "" : row[sourceIdx] ?? "";
      }

      const stateId = String(row[colIdx("state id")] ?? "").trim();
      const isDuplicate = existingStateIds.has(stateId);

      results.push({
        id: `${tabName}-${stateId}-${results.length}`,
        tab: tabName,
        stateId,
        lastName: row[colIdx("last name")] ?? "",
        firstName: row[colIdx("first name")] ?? "",
        office: row[colIdx("office")] ?? "",
        caseManager: row[colIdx("case manager")] ?? "",
        status: row[statusIdx] ?? "",
        fund,
        isDuplicate,
        included: !isDuplicate, // duplicates start unchecked so you consciously opt in
        outputRow,
      });
    }
  }

  return { results, missingTargets };
}

// -----------------------------------------------------------------------
// STEP 4b: Build a FULL updated copy of HCGY_Caseload.xlsx, with the new
// rows appended directly into the Caseload tab.
//
// This does NOT use the xlsx library to read/rewrite the whole workbook
// (that was tested and rejected -- see README for why). Instead it treats
// the .xlsx as what it physically is, a zip file, and edits only the one
// internal XML file that holds the Caseload tab's rows. Every other file
// in the zip (including the two huge "Data" tabs) is copied through
// completely untouched, byte-for-byte. That's what makes this both fast
// and safe.
// -----------------------------------------------------------------------

// Converts a JS Date into the day-count format Excel stores dates as
// (days since 1899-12-30). Excel/Caseload stores dates as plain integers
// with no time component, which is what this produces.
function dateToExcelSerial(date) {
  const utcDate = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  const excelEpoch = Date.UTC(1899, 11, 30);
  return Math.round((utcDate - excelEpoch) / 86400000);
}

// Turns a 0-based column index into a spreadsheet column letter
// (0 -> A, 12 -> M, 38 -> AM, ...).
function columnLetter(index0) {
  let n = index0 + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function escapeXml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// Row height / row number metadata comes from the literal last row.
function parseLastRowMeta(xml) {
  const sheetDataEnd = xml.indexOf("</sheetData>");
  const lastRowOpen = xml.lastIndexOf("<row ", sheetDataEnd);
  const lastRowClose = xml.indexOf("</row>", lastRowOpen) + "</row>".length;
  const block = xml.slice(lastRowOpen, lastRowClose);

  const rowNumMatch = block.match(/<row r="(\d+)"/);
  const htMatch = block.match(/ ht="([^"]+)"/);
  const hasCustomHeight = / customHeight="1"/.test(block);

  return {
    rowNum: rowNumMatch ? parseInt(rowNumMatch[1], 10) : null,
    ht: htMatch ? htMatch[1] : null,
    hasCustomHeight,
  };
}

// Per-column CELL STYLE (formatting) does NOT come from just the last row --
// the very last row or two in this sheet are often near-empty placeholder
// rows, so a date column that's blank there would hand back a plain
// "General" style instead of a real date format. Instead, scan rows from
// the bottom up and, for each column we're about to write, use the style
// from the nearest row where that column actually has a real value. That
// reliably finds a properly-formatted cell (e.g. an actual date format for
// date columns) rather than whatever the last, possibly-blank row happened
// to have.
function deriveColumnStyles(xml, neededColumnLetters) {
  const sheetDataStart = xml.indexOf("<sheetData>") + "<sheetData>".length;
  const sheetDataEnd = xml.indexOf("</sheetData>");
  const body = xml.slice(sheetDataStart, sheetDataEnd);
  const rowBlocks = body.match(/<row [^>]*>[\s\S]*?<\/row>/g) || [];

  const styles = {};
  const remaining = new Set(neededColumnLetters);
  for (let i = rowBlocks.length - 1; i >= 0 && remaining.size > 0; i--) {
    const block = rowBlocks[i];
    for (const col of Array.from(remaining)) {
      const cellRegex = new RegExp(`<c r="${col}\\d+"(?: s="(\\d+)")?[^>]*>(?:<v>[^<]+</v>|<is>)`);
      const match = block.match(cellRegex);
      if (match && match[1]) {
        styles[col] = match[1];
        remaining.delete(col);
      }
    }
  }
  return styles;
}

function buildRowXml(rowNum, outputRow, styleByColumn, meta) {
  let cells = "";
  outputRow.forEach((value, idx) => {
    if (value === "" || value === null || value === undefined) return;
    const col = columnLetter(idx);
    const style = styleByColumn[col];
    const styleAttr = style ? ` s="${style}"` : "";

    if (value instanceof Date) {
      cells += `<c r="${col}${rowNum}"${styleAttr}><v>${dateToExcelSerial(value)}</v></c>`;
    } else if (typeof value === "number") {
      cells += `<c r="${col}${rowNum}"${styleAttr}><v>${value}</v></c>`;
    } else {
      cells += `<c r="${col}${rowNum}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(
        String(value)
      )}</t></is></c>`;
    }
  });

  const htAttr = meta.ht ? ` ht="${meta.ht}"` : "";
  const chAttr = meta.hasCustomHeight ? ` customHeight="1"` : "";
  return `<row r="${rowNum}" spans="1:39"${htAttr}${chAttr} x14ac:dyDescent="0.25">${cells}</row>`;
}

export async function buildUpdatedCaseloadWorkbook(caseloadFile, selectedResults, outputFilename) {
  const JSZipModule = await import("jszip");
  const JSZip = JSZipModule.default;

  const buf = await caseloadFile.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);

  // Find which physical worksheetN.xml file the "Caseload" tab actually is
  // (this is stored indirectly via an r:id, not a fixed filename).
  const workbookXml = await zip.file("xl/workbook.xml").async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");

  const sheetMatch = workbookXml.match(
    new RegExp(`<sheet[^>]*name="${CASELOAD_SHEET_NAME}"[^>]*r:id="([^"]+)"`)
  );
  if (!sheetMatch) {
    throw new Error(`Could not find the "${CASELOAD_SHEET_NAME}" tab inside that workbook.`);
  }
  const rId = sheetMatch[1];

  const relMatch = relsXml.match(new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*Target="([^"]+)"`));
  if (!relMatch) {
    throw new Error("Could not locate the Caseload tab's file inside the workbook.");
  }
  const target = relMatch[1];
  const sheetPath = target.startsWith("/") ? target.slice(1) : `xl/${target}`;

  let xml = await zip.file(sheetPath).async("string");

  const meta = parseLastRowMeta(xml);
  if (meta.rowNum === null) {
    throw new Error("Couldn't find any existing rows in the Caseload tab to match formatting from.");
  }

  // Only look up styles for columns we're actually about to write into.
  const neededColumns = new Set();
  for (const result of selectedResults) {
    result.outputRow.forEach((value, idx) => {
      if (value !== "" && value !== null && value !== undefined) {
        neededColumns.add(columnLetter(idx));
      }
    });
  }
  const styleByColumn = deriveColumnStyles(xml, Array.from(neededColumns));

  const newRowsXml = selectedResults
    .map((result, i) => buildRowXml(meta.rowNum + 1 + i, result.outputRow, styleByColumn, meta))
    .join("");

  xml = xml.replace("</sheetData>", `${newRowsXml}</sheetData>`);

  // Extend the sheet's declared dimension to cover the new rows.
  const newLastRow = meta.rowNum + selectedResults.length;
  const dimMatch = xml.match(/<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"\/>/);
  if (dimMatch) {
    xml = xml.replace(dimMatch[0], `<dimension ref="${dimMatch[1]}:${dimMatch[2]}${newLastRow}"/>`);
  }

  zip.file(sheetPath, xml);

  const outBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
  const url = URL.createObjectURL(outBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = outputFilename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// -----------------------------------------------------------------------
// STEP 4: Build and download the small "New Rows" workbook, using the
// Caseload tab's real header row so columns line up exactly for
// paste-special.
// -----------------------------------------------------------------------
export async function downloadNewRowsWorkbook(caseloadHeader, selectedResults, filename) {
  const XLSX = await import("xlsx");
  const aoa = [caseloadHeader, ...selectedResults.map((r) => r.outputRow)];
  const sheet = XLSX.utils.aoa_to_sheet(aoa);

  // Reasonable default column widths so the file is readable before pasting.
  sheet["!cols"] = caseloadHeader.map(() => ({ wch: 18 }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "New Rows");
  XLSX.writeFile(workbook, filename);
}
