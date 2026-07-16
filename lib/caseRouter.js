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
