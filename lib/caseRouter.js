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
// This is the definitive list of Case Tracker columns to carry over,
// per your spec: A, D, F, G, I, J, K, L, M, O, P, R, S, T, U.
//
// Fields with a `targetMatch` already have a home in the Caseload tab's
// existing columns. Fields marked `isNew: true` don't -- those get brand
// new columns appended to the end of the Caseload tab (same label as in
// the Case Tracker), rather than being dropped.
// -----------------------------------------------------------------------
const FIELD_MAP = [
  // Column A
  { key: "eliSpec", label: "Eli Spec.", sourceHeader: "eli spec.", isNew: true, type: "text" },
  // Column D
  { key: "stateId", label: "State ID", sourceHeader: "state id", targetMatch: (h) => h === "state id" },
  // Column F
  { key: "caseManager", label: "Case Manager", sourceHeader: "case manager", targetMatch: (h) => h.startsWith("case manager from pwe") },
  // Column G
  { key: "assessDate", label: "Assess Date", sourceHeader: "assess date", isNew: true, type: "date" },
  // Column I
  { key: "lastName", label: "Last Name", sourceHeader: "last name", targetMatch: (h) => h === "last name" },
  // Column J
  { key: "firstName", label: "First Name", sourceHeader: "first name", targetMatch: (h) => h === "first name" },
  // Column K
  { key: "office", label: "Office", sourceHeader: "office", targetMatch: (h) => h === "office" },
  // Column L
  { key: "assignDate", label: "Assign Date", sourceHeader: "assign date", targetMatch: (h) => h === "date assigned" },
  // Column M
  { key: "fund", label: "Fund", sourceHeader: "fund", isNew: true, type: "text" },
  // Column O
  { key: "status", label: "Status", sourceHeader: "status", targetMatch: (h) => h === "status" },
  // Column P
  { key: "notes", label: "Notes", sourceHeader: "notes", targetMatch: (h) => h === "notes" },
  // Column R
  { key: "completeDate", label: "Complete Date", sourceHeader: "complete date", isNew: true, type: "date" },
  // Column S
  { key: "signDue", label: "Sign Due", sourceHeader: "sign due", targetMatch: (h) => h === "signature due" },
  // Column T
  { key: "signDate", label: "Sign Date", sourceHeader: "sign date", targetMatch: (h) => h === "sign date" },
  // Column U
  { key: "pwe", label: "PWE", sourceHeader: "pwe", isNew: true, type: "text" },
];

const MATCHED_FIELDS = FIELD_MAP.filter((f) => !f.isNew);
const NEW_FIELDS = FIELD_MAP.filter((f) => f.isNew);

// Fields used only for filtering / preview, in addition to being carried
// over as the "Fund" field above.
const FUND_SOURCE_HEADER = "fund";

// Given the Caseload tab's header row, figure out which column index each
// matched FIELD_MAP entry should be written to.
function buildTargetColumnMap(targetHeader) {
  const normalizedTargetHeader = targetHeader.map(normalize);
  const map = {};
  for (const field of MATCHED_FIELDS) {
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
//
// The output row is wider than the Caseload tab's original header --
// it's [...existing Caseload columns, ...5 brand-new columns] -- since
// New fields (Eli Spec., Assess Date, Fund, Complete Date, PWE) get
// appended past the end. `outputHeader` reflects that full column set
// and should be used for anything downstream (downloads, the full-file
// writer) instead of the raw `caseloadHeader`.
// -----------------------------------------------------------------------
export async function filterAndMapRows({ workbook, tabNames, caseloadHeader, existingStateIds }) {
  const XLSX = await import("xlsx");
  const targetColMap = buildTargetColumnMap(caseloadHeader);
  const missingTargets = MATCHED_FIELDS.filter((f) => targetColMap[f.key] === -1).map((f) => f.label);

  const newFieldStartIdx = caseloadHeader.length;
  const newFieldColIdx = {};
  NEW_FIELDS.forEach((f, i) => { newFieldColIdx[f.key] = newFieldStartIdx + i; });

  const outputHeader = [...caseloadHeader, ...NEW_FIELDS.map((f) => f.label)];

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

      // Build the output row: existing Caseload columns + the 5 new ones.
      const outputRow = new Array(outputHeader.length).fill("");
      for (const field of MATCHED_FIELDS) {
        const targetIdx = targetColMap[field.key];
        if (targetIdx === -1) continue;
        const sourceIdx = colIdx(field.sourceHeader);
        outputRow[targetIdx] = sourceIdx === -1 ? "" : row[sourceIdx] ?? "";
      }
      for (const field of NEW_FIELDS) {
        const sourceIdx = colIdx(field.sourceHeader);
        outputRow[newFieldColIdx[field.key]] = sourceIdx === -1 ? "" : row[sourceIdx] ?? "";
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

  return { results, missingTargets, outputHeader, newFieldStartIdx };
}

// -----------------------------------------------------------------------
// STEP 4b: Build a FULL updated copy of HCGY_Caseload.xlsx, with the new
// rows appended directly into the Caseload tab, new columns added for
// fields that don't have an existing home, and every new cell highlighted
// with a yellow fill so the additions are easy to spot.
//
// This does NOT use the xlsx library to read/rewrite the whole workbook
// (that was tested and rejected -- see README for why). Instead it treats
// the .xlsx as what it physically is, a zip file, and edits only the two
// internal XML files that actually need to change: the Caseload sheet's
// XML and styles.xml (for the yellow-highlight styles). Every other file
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

// The sheet's very last declared row is often a blank pre-formatted
// spacer row (sheets like this tend to have a block of them at the
// bottom), not real data. Find the actual last row that has a real value
// in it, so new rows can be placed right after real data instead of
// after a wall of blank spacer rows.
function findLastRealDataRowNum(xml) {
  const sheetDataStart = xml.indexOf("<sheetData>") + "<sheetData>".length;
  const sheetDataEnd = xml.indexOf("</sheetData>");
  const body = xml.slice(sheetDataStart, sheetDataEnd);
  const rowBlocks = body.match(/<row [^>]*>[\s\S]*?<\/row>/g) || [];

  for (let i = rowBlocks.length - 1; i >= 0; i--) {
    const block = rowBlocks[i];
    if (/<v>[^<]+<\/v>|<is>/.test(block)) {
      const m = block.match(/<row r="(\d+)"/);
      if (m) return parseInt(m[1], 10);
    }
  }
  return null;
}

// Per-column CELL STYLE (formatting) does NOT come from just the last row --
// the very last row or two in this sheet are often near-empty placeholder
// rows, so a date column that's blank there would hand back a plain
// "General" style instead of a real date format. Instead, scan rows from
// the bottom up and, for each column we're about to write, use the style
// from the nearest row where that column actually has a real value.
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

// Splits the raw <cellXfs>...</cellXfs> body into individual <xf> element
// strings. A plain regex isn't reliable here because self-closing entries
// (<xf .../>) and entries-with-children (<xf ...>...</xf>) need different
// end detection, so this walks the string manually instead.
function splitXfEntries(body) {
  const entries = [];
  let i = 0;
  while (i < body.length) {
    const start = body.indexOf("<xf", i);
    if (start === -1) break;
    const tagEnd = body.indexOf(">", start);
    if (body[tagEnd - 1] === "/") {
      entries.push(body.slice(start, tagEnd + 1));
      i = tagEnd + 1;
    } else {
      const closeTag = body.indexOf("</xf>", tagEnd);
      entries.push(body.slice(start, closeTag + "</xf>".length));
      i = closeTag + "</xf>".length;
    }
  }
  return entries;
}

// Finds an existing solid-color fill matching the given ARGB hex (e.g.
// "FFFFFF00" for yellow) in styles.xml, so we can reuse it instead of
// declaring a duplicate. Returns null if nothing matches.
function findExistingFillIndex(stylesXml, argbHex) {
  const fillsMatch = stylesXml.match(/<fills count="\d+">([\s\S]*?)<\/fills>/);
  if (!fillsMatch) return null;
  const fillList = fillsMatch[1].match(/<fill>[\s\S]*?<\/fill>/g) || [];
  const idx = fillList.findIndex((f) => f.includes(`rgb="${argbHex}"`));
  return idx === -1 ? null : idx;
}

// Given styles.xml text and a set of existing cellXfs indexes, returns
// updated styles.xml text plus a map of originalStyleIndex -> new
// yellow-highlighted style index, by cloning each needed style and
// pointing the clone at a yellow fill (reusing one if the workbook
// already has one, since HCGY_Caseload.xlsx already does).
function addYellowStyleVariants(stylesXml, neededOriginalIndexes) {
  const YELLOW_ARGB = "FFFFFF00";
  let yellowFillIdx = findExistingFillIndex(stylesXml, YELLOW_ARGB);
  let workingXml = stylesXml;

  if (yellowFillIdx === null) {
    const fillsMatch = workingXml.match(/<fills count="(\d+)">([\s\S]*?)<\/fills>/);
    const fillsCount = parseInt(fillsMatch[1], 10);
    const newFill = `<fill><patternFill patternType="solid"><fgColor rgb="${YELLOW_ARGB}"/><bgColor indexed="64"/></patternFill></fill>`;
    workingXml = workingXml.replace(
      fillsMatch[0],
      `<fills count="${fillsCount + 1}">${fillsMatch[2]}${newFill}</fills>`
    );
    yellowFillIdx = fillsCount;
  }

  const cellXfsMatch = workingXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  const origCount = parseInt(cellXfsMatch[1], 10);
  const xfList = splitXfEntries(cellXfsMatch[2]);

  function makeYellowVariant(origIdx) {
    let xf = xfList[origIdx];
    if (/fillId="\d+"/.test(xf)) {
      xf = xf.replace(/fillId="\d+"/, `fillId="${yellowFillIdx}"`);
    } else {
      xf = xf.replace("<xf ", `<xf fillId="${yellowFillIdx}" `);
    }
    if (!/applyFill="1"/.test(xf)) {
      xf = xf.replace("<xf ", `<xf applyFill="1" `);
    }
    return xf;
  }

  const variantMap = {};
  const newEntries = [];
  let nextIdx = origCount;
  for (const origIdx of neededOriginalIndexes) {
    if (origIdx === null || origIdx === undefined || Number.isNaN(origIdx)) continue;
    newEntries.push(makeYellowVariant(origIdx));
    variantMap[origIdx] = nextIdx;
    nextIdx++;
  }

  const newCount = origCount + newEntries.length;
  workingXml = workingXml.replace(
    cellXfsMatch[0],
    `<cellXfs count="${newCount}">${cellXfsMatch[2]}${newEntries.join("")}</cellXfs>`
  );

  return { stylesXml: workingXml, variantMap, xfList };
}

export async function buildUpdatedCaseloadWorkbook(caseloadFile, selectedResults, outputHeader, newFieldStartIdx, outputFilename) {
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

  // Styles for the existing (matched) columns we're about to write into.
  const neededExistingColumns = new Set();
  for (const result of selectedResults) {
    result.outputRow.forEach((value, idx) => {
      if (idx < newFieldStartIdx && value !== "" && value !== null && value !== undefined) {
        neededExistingColumns.add(columnLetter(idx));
      }
    });
  }
  const styleByColumn = deriveColumnStyles(xml, Array.from(neededExistingColumns));

  // Basis styles for the brand-new columns: reuse a known-good date style
  // (from Assign Date) and a known-good text style (from Notes), so new
  // columns look consistent with the rest of the sheet rather than using
  // completely unstyled cells. Found dynamically (not hardcoded letters)
  // in case the Caseload layout ever shifts.
  const normalizedOutputHeader = outputHeader.map(normalize);
  const assignDateField = MATCHED_FIELDS.find((f) => f.key === "assignDate");
  const notesField = MATCHED_FIELDS.find((f) => f.key === "notes");
  const assignDateColIdx = assignDateField ? normalizedOutputHeader.findIndex((h) => assignDateField.targetMatch(h)) : -1;
  const notesColIdx = notesField ? normalizedOutputHeader.findIndex((h) => notesField.targetMatch(h)) : -1;
  const assignDateStyle = assignDateColIdx >= 0 ? deriveColumnStyles(xml, [columnLetter(assignDateColIdx)])[columnLetter(assignDateColIdx)] : null;
  const notesStyle = notesColIdx >= 0 ? deriveColumnStyles(xml, [columnLetter(notesColIdx)])[columnLetter(notesColIdx)] : null;

  // Build yellow-highlight variants of every style we're about to use.
  let styles = await zip.file("xl/styles.xml").async("string");
  const neededOriginalStyleIndexes = new Set(
    Object.values(styleByColumn).map((s) => parseInt(s, 10))
  );
  if (assignDateStyle) neededOriginalStyleIndexes.add(parseInt(assignDateStyle, 10));
  if (notesStyle) neededOriginalStyleIndexes.add(parseInt(notesStyle, 10));

  const { stylesXml: updatedStyles, variantMap } = addYellowStyleVariants(
    styles,
    Array.from(neededOriginalStyleIndexes)
  );
  zip.file("xl/styles.xml", updatedStyles);

  const yellowDateStyleIdx = assignDateStyle ? variantMap[parseInt(assignDateStyle, 10)] : undefined;
  const yellowTextStyleIdx = notesStyle ? variantMap[parseInt(notesStyle, 10)] : undefined;

  // Append header cells for the new columns onto row 1.
  const row1Match = xml.match(/<row r="1"[^>]*>[\s\S]*?<\/row>/);
  if (row1Match && NEW_FIELDS.length > 0) {
    const row1Block = row1Match[0];
    const lastHeaderStyleMatch = row1Block.match(/<c r="[A-Z]+1" s="(\d+)"[^>]*>[\s\S]*?<\/c>(?=<\/row>)/);
    const headerStyle = lastHeaderStyleMatch ? lastHeaderStyleMatch[1] : null;
    const headerStyleAttr = headerStyle ? ` s="${headerStyle}"` : "";

    let newHeaderCells = "";
    NEW_FIELDS.forEach((field, i) => {
      const col = columnLetter(newFieldStartIdx + i);
      newHeaderCells += `<c r="${col}1"${headerStyleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(field.label)}</t></is></c>`;
    });

    const newRow1Block = row1Block.replace("</row>", `${newHeaderCells}</row>`);
    xml = xml.replace(row1Block, newRow1Block);
  }

  const totalCols = newFieldStartIdx + NEW_FIELDS.length;

  function buildRowXml(rowNum, outputRow) {
    let cells = "";
    outputRow.forEach((value, idx) => {
      if (value === "" || value === null || value === undefined) return;
      const col = columnLetter(idx);
      let styleIdx;
      if (idx >= newFieldStartIdx) {
        const field = NEW_FIELDS[idx - newFieldStartIdx];
        styleIdx = field.type === "date" ? yellowDateStyleIdx : yellowTextStyleIdx;
      } else {
        const origStyle = styleByColumn[col];
        styleIdx = origStyle ? variantMap[parseInt(origStyle, 10)] : undefined;
      }
      const styleAttr = styleIdx !== undefined ? ` s="${styleIdx}"` : "";

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
    return `<row r="${rowNum}" spans="1:${totalCols}"${htAttr}${chAttr} x14ac:dyDescent="0.25">${cells}</row>`;
  }

  // Find any blank pre-formatted spacer rows sitting between the last row
  // of real data and the sheet's last declared row, and reuse THOSE for
  // the new data instead of always appending after every blank row --
  // that way new rows show up right after your real data, not buried
  // under a wall of blank spacer rows.
  const lastRealDataRowNum = findLastRealDataRowNum(xml) ?? meta.rowNum;
  const blankSlotNumbers = [];
  for (let r = lastRealDataRowNum + 1; r <= meta.rowNum; r++) blankSlotNumbers.push(r);

  const replacements = []; // { rowNum, block }
  const appended = []; // { rowNum, block }
  let overflowCount = 0;

  selectedResults.forEach((result, i) => {
    if (i < blankSlotNumbers.length) {
      const rowNum = blankSlotNumbers[i];
      replacements.push({ rowNum, block: buildRowXml(rowNum, result.outputRow) });
    } else {
      overflowCount++;
      const rowNum = meta.rowNum + overflowCount;
      appended.push({ rowNum, block: buildRowXml(rowNum, result.outputRow) });
    }
  });

  // Replace each reused blank row in place (same row number, real content).
  for (const { rowNum, block } of replacements) {
    const rowRegex = new RegExp(`<row r="${rowNum}"[^>]*>[\\s\\S]*?<\\/row>`);
    if (rowRegex.test(xml)) {
      xml = xml.replace(rowRegex, block);
    } else {
      // Shouldn't happen (we just found this row number), but fall back
      // to appending rather than silently dropping the row.
      appended.push({ rowNum, block });
    }
  }

  // Anything that didn't fit in a blank slot gets appended past the end.
  if (appended.length > 0) {
    const appendedXml = appended.map((a) => a.block).join("");
    xml = xml.replace("</sheetData>", `${appendedXml}</sheetData>`);
  }

  // Extend the sheet's declared dimension to cover any appended rows AND
  // the new columns. The row-count only needs to grow if we ran past the
  // existing blank rows; reused blank rows were already inside the range.
  const newLastRow = Math.max(meta.rowNum, ...(appended.length ? appended.map((a) => a.rowNum) : [meta.rowNum]));
  const newLastCol = columnLetter(totalCols - 1);
  const dimMatch = xml.match(/<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"\/>/);
  if (dimMatch) {
    xml = xml.replace(dimMatch[0], `<dimension ref="${dimMatch[1]}:${newLastCol}${newLastRow}"/>`);
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
// full output header (existing Caseload columns + the 5 new ones) so
// columns line up exactly for paste-special.
// -----------------------------------------------------------------------
export async function downloadNewRowsWorkbook(outputHeader, selectedResults, filename) {
  const XLSX = await import("xlsx");
  const aoa = [outputHeader, ...selectedResults.map((r) => r.outputRow)];
  const sheet = XLSX.utils.aoa_to_sheet(aoa);

  // Reasonable default column widths so the file is readable before pasting.
  sheet["!cols"] = outputHeader.map(() => ({ wch: 18 }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "New Rows");
  XLSX.writeFile(workbook, filename);
}
