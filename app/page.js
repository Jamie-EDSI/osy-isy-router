"use client";

import { useMemo, useState, useCallback } from "react";
import {
  readCaseTrackerFile,
  readCaseloadSheet,
  filterAndMapRows,
  downloadNewRowsWorkbook,
  buildUpdatedCaseloadWorkbook,
} from "@/lib/caseRouter";

// Small helper so date-ish cells display nicely in the preview table
// without changing the underlying value we actually export.
function displayValue(value) {
  if (value instanceof Date) return value.toLocaleDateString("en-US");
  return String(value ?? "");
}

// One numbered step in the routing-slip header. Order is meaningful here
// (you can't filter before both files are loaded), so numbering earns
// its place.
function StepBadge({ n, done }) {
  return (
    <span
      className={`inline-flex h-7 w-7 items-center justify-center rounded-full border-2 font-display text-sm font-bold ${
        done ? "border-teal bg-teal text-white" : "border-line text-ink-soft"
      }`}
    >
      {n}
    </span>
  );
}

function FileDropCard({ title, hint, file, onFile, accept = ".xlsx" }) {
  const [dragOver, setDragOver] = useState(false);

  const handleFiles = (files) => {
    if (files && files[0]) onFile(files[0]);
  };

  return (
    <label
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        handleFiles(e.dataTransfer.files);
      }}
      className={`block cursor-pointer rounded-lg border-2 border-dashed p-6 transition-colors ${
        dragOver ? "border-teal bg-teal-tint" : "border-line bg-panel hover:border-teal/60"
      }`}
    >
      <input
        type="file"
        accept={accept}
        className="hidden"
        onChange={(e) => handleFiles(e.target.files)}
      />
      <div className="font-display text-sm font-semibold text-ink">{title}</div>
      <p className="mt-1 text-sm text-ink-soft">{hint}</p>
      {file ? (
        <div className="mt-3 inline-flex items-center gap-2 rounded-md bg-teal-tint px-3 py-1.5 text-sm font-medium text-teal-dark">
          <span className="font-mono">{file.name}</span>
        </div>
      ) : (
        <div className="mt-3 text-sm text-ink-soft">Click to browse, or drop a file here.</div>
      )}
    </label>
  );
}

export default function Home() {
  // ---- File 1: Case Tracker -------------------------------------------
  const [trackerFile, setTrackerFile] = useState(null);
  const [trackerWorkbook, setTrackerWorkbook] = useState(null);
  const [monthlyTabs, setMonthlyTabs] = useState([]);
  const [selectedTabs, setSelectedTabs] = useState([]);
  const [defaultTab, setDefaultTab] = useState(null);

  // ---- File 2: Caseload --------------------------------------------
  const [caseloadFile, setCaseloadFile] = useState(null);
  const [caseloadHeader, setCaseloadHeader] = useState(null);
  const [existingStateIds, setExistingStateIds] = useState(new Set());

  // ---- Results ----------------------------------------------------
  const [results, setResults] = useState(null); // null = not run yet
  const [missingTargets, setMissingTargets] = useState([]);

  // ---- UI state -----------------------------------------------------
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const bothFilesLoaded = trackerWorkbook && caseloadHeader;

  const handleTrackerFile = useCallback(async (file) => {
    setError("");
    setBusy(true);
    try {
      const { workbook, monthlyTabs, latestTab } = await readCaseTrackerFile(file);
      setTrackerFile(file);
      setTrackerWorkbook(workbook);
      setMonthlyTabs(monthlyTabs);
      setSelectedTabs([latestTab]);
      setDefaultTab(latestTab);
      setResults(null);
    } catch (err) {
      setError(err.message || "Couldn't read that Case Tracker file.");
    } finally {
      setBusy(false);
    }
  }, []);

  const handleCaseloadFile = useCallback(async (file) => {
    setError("");
    setBusy(true);
    try {
      const { header, existingStateIds } = await readCaseloadSheet(file);
      setCaseloadFile(file);
      setCaseloadHeader(header);
      setExistingStateIds(existingStateIds);
      setResults(null);
    } catch (err) {
      setError(err.message || "Couldn't read that Caseload file.");
    } finally {
      setBusy(false);
    }
  }, []);

  const toggleTab = (tab) => {
    setSelectedTabs((prev) =>
      prev.includes(tab) ? prev.filter((t) => t !== tab) : [...prev, tab]
    );
  };

  const runFilter = async () => {
    if (!trackerWorkbook || !caseloadHeader || selectedTabs.length === 0) return;
    setError("");
    setBusy(true);
    try {
      const { results, missingTargets } = await filterAndMapRows({
        workbook: trackerWorkbook,
        tabNames: selectedTabs,
        caseloadHeader,
        existingStateIds,
      });
      setResults(results);
      setMissingTargets(missingTargets);
    } catch (err) {
      setError(err.message || "Something went wrong while filtering.");
    } finally {
      setBusy(false);
    }
  };

  const toggleRow = (id) => {
    setResults((prev) =>
      prev.map((r) => (r.id === id ? { ...r, included: !r.included } : r))
    );
  };

  const selectedCount = useMemo(
    () => (results ? results.filter((r) => r.included).length : 0),
    [results]
  );
  const duplicateCount = useMemo(
    () => (results ? results.filter((r) => r.isDuplicate).length : 0),
    [results]
  );

  const handleDownload = async () => {
    const selected = results.filter((r) => r.included);
    if (selected.length === 0) return;
    const stamp = new Date().toISOString().slice(0, 10);
    await downloadNewRowsWorkbook(caseloadHeader, selected, `New_Rows_to_Add_${stamp}.xlsx`);
  };

  const handleDownloadFullFile = async () => {
    const selected = results.filter((r) => r.included);
    if (selected.length === 0 || !caseloadFile) return;
    setError("");
    setBusy(true);
    try {
      const stamp = new Date().toISOString().slice(0, 10);
      const baseName = caseloadFile.name.replace(/\.xlsx$/i, "");
      await buildUpdatedCaseloadWorkbook(caseloadFile, selected, `${baseName}_updated_${stamp}.xlsx`);
    } catch (err) {
      setError(err.message || "Couldn't build the updated Caseload file.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="mx-auto w-full max-w-5xl flex-1 px-6 py-10">
      {/* ---- Header / routing slip title block ---- */}
      <header className="mb-8 border-b-2 border-ink pb-4">
        <p className="font-mono text-xs uppercase tracking-[0.2em] text-ink-soft">
          Case Routing Utility
        </p>
        <h1 className="mt-1 font-display text-3xl font-bold text-ink">
          OSY / ISY Case Router
        </h1>
        <p className="mt-2 max-w-2xl text-sm text-ink-soft">
          Finds APVD cases with an OSY or ISY fund code in the Case Tracker, checks
          them against the Caseload file for duplicates, and builds a small file of
          new rows ready to paste in. Your files stay in this browser tab &mdash;
          nothing is uploaded anywhere.
        </p>
      </header>

      {error && (
        <div className="mb-6 rounded-md border-2 border-red-stamp bg-red-tint px-4 py-3 text-sm font-medium text-red-stamp">
          {error}
        </div>
      )}

      {/* ---- Step 1 & 2: uploads ---- */}
      <section className="grid gap-4 sm:grid-cols-2">
        <div>
          <div className="mb-2 flex items-center gap-2">
            <StepBadge n={1} done={!!trackerWorkbook} />
            <span className="font-display text-sm font-semibold text-ink">Case Tracker file</span>
          </div>
          <FileDropCard
            title="1_Case_Tracker_PY2026.xlsx"
            hint="The file with the monthly tabs (e.g. 2026.07) and Status / Fund columns."
            file={trackerFile}
            onFile={handleTrackerFile}
          />
        </div>
        <div>
          <div className="mb-2 flex items-center gap-2">
            <StepBadge n={2} done={!!caseloadHeader} />
            <span className="font-display text-sm font-semibold text-ink">Caseload file</span>
          </div>
          <FileDropCard
            title="HCGY_Caseload.xlsx"
            hint="The larger file. Only the Caseload tab is read; the rest is left untouched."
            file={caseloadFile}
            onFile={handleCaseloadFile}
          />
        </div>
      </section>

      {/* ---- Step 3: pick tabs + run ---- */}
      {trackerWorkbook && (
        <section className="mt-8">
          <div className="mb-2 flex items-center gap-2">
            <StepBadge n={3} done={!!results} />
            <span className="font-display text-sm font-semibold text-ink">
              Choose month tab(s) to scan
            </span>
          </div>
          <div className="rounded-lg border border-line bg-panel p-4">
            <div className="flex flex-wrap gap-2">
              {monthlyTabs.map((tab) => {
                const active = selectedTabs.includes(tab);
                return (
                  <button
                    key={tab}
                    onClick={() => toggleTab(tab)}
                    className={`rounded-md border px-3 py-1.5 font-mono text-sm transition-colors ${
                      active
                        ? "border-teal bg-teal text-white"
                        : "border-line bg-paper text-ink-soft hover:border-teal/60"
                    }`}
                  >
                    {tab}
                  </button>
                );
              })}
            </div>
            <p className="mt-3 text-xs text-ink-soft">
              Defaults to the most recent tab with data ({defaultTab}). Click to add or remove
              tabs from the scan.
            </p>
            <button
              onClick={runFilter}
              disabled={!bothFilesLoaded || selectedTabs.length === 0 || busy}
              className="mt-4 rounded-md bg-ink px-5 py-2.5 font-display text-sm font-semibold text-white transition-opacity disabled:opacity-40"
            >
              {busy ? "Working…" : "Find matching cases"}
            </button>
            {!caseloadHeader && (
              <p className="mt-2 text-xs text-amber">Upload the Caseload file first.</p>
            )}
          </div>
        </section>
      )}

      {/* ---- Step 4: results ---- */}
      {results && (
        <section className="mt-8">
          <div className="mb-2 flex items-center gap-2">
            <StepBadge n={4} done={selectedCount > 0} />
            <span className="font-display text-sm font-semibold text-ink">
              Review &amp; export
            </span>
          </div>

          <div className="mb-3 flex flex-wrap items-center gap-3 text-sm">
            <span className="rounded-md bg-teal-tint px-3 py-1 font-medium text-teal-dark">
              {results.length} match{results.length === 1 ? "" : "es"} found
            </span>
            {duplicateCount > 0 && (
              <span className="rounded-md bg-amber-tint px-3 py-1 font-medium text-amber">
                {duplicateCount} possible duplicate{duplicateCount === 1 ? "" : "s"} (unchecked by default)
              </span>
            )}
            <span className="rounded-md bg-line/60 px-3 py-1 font-medium text-ink-soft">
              {selectedCount} selected to export
            </span>
          </div>

          {missingTargets.length > 0 && (
            <div className="mb-3 rounded-md border border-amber bg-amber-tint px-4 py-2 text-xs text-amber">
              Heads up: couldn&apos;t find a matching Caseload column for: {missingTargets.join(", ")}.
              Those fields were skipped.
            </div>
          )}

          {results.length === 0 ? (
            <div className="rounded-lg border border-line bg-panel p-6 text-sm text-ink-soft">
              No rows in the selected tab(s) matched Status = APVD with an OSY/ISY fund code.
            </div>
          ) : (
            <div className="overflow-x-auto rounded-lg border border-line bg-panel">
              <table className="w-full min-w-[840px] text-sm">
                <thead>
                  <tr className="border-b border-line bg-paper text-left font-display text-xs uppercase tracking-wide text-ink-soft">
                    <th className="w-10 px-3 py-2"></th>
                    <th className="px-3 py-2">State ID</th>
                    <th className="px-3 py-2">Name</th>
                    <th className="px-3 py-2">Office</th>
                    <th className="px-3 py-2">Case Manager</th>
                    <th className="px-3 py-2">Fund</th>
                    <th className="px-3 py-2">Tab</th>
                    <th className="px-3 py-2"></th>
                  </tr>
                </thead>
                <tbody>
                  {results.map((r) => (
                    <tr
                      key={r.id}
                      className={`border-b border-line last:border-0 ${
                        r.isDuplicate ? "bg-red-tint/40" : ""
                      }`}
                    >
                      <td className="px-3 py-2">
                        <input
                          type="checkbox"
                          checked={r.included}
                          onChange={() => toggleRow(r.id)}
                          className="h-4 w-4 accent-teal"
                        />
                      </td>
                      <td className="px-3 py-2 font-mono">{r.stateId}</td>
                      <td className="px-3 py-2">
                        {displayValue(r.lastName)}, {displayValue(r.firstName)}
                      </td>
                      <td className="px-3 py-2 text-ink-soft">{displayValue(r.office)}</td>
                      <td className="px-3 py-2 text-ink-soft">{displayValue(r.caseManager)}</td>
                      <td className="px-3 py-2 font-mono text-ink-soft">{r.fund}</td>
                      <td className="px-3 py-2 font-mono text-ink-soft">{r.tab}</td>
                      <td className="px-3 py-2">
                        {r.isDuplicate && (
                          <span className="stamp text-red-stamp">On file</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          <div className="mt-5 flex flex-wrap gap-3">
            <button
              onClick={handleDownloadFullFile}
              disabled={selectedCount === 0 || busy}
              className="rounded-md bg-teal px-5 py-2.5 font-display text-sm font-semibold text-white transition-opacity disabled:opacity-40"
            >
              {busy ? "Building…" : `Download updated Caseload file (${selectedCount} new)`}
            </button>
            <button
              onClick={handleDownload}
              disabled={selectedCount === 0}
              className="rounded-md border border-ink bg-panel px-5 py-2.5 font-display text-sm font-semibold text-ink transition-opacity disabled:opacity-40"
            >
              Download new rows only (backup copy)
            </button>
          </div>
          <p className="mt-2 text-xs text-ink-soft">
            <strong>Updated Caseload file</strong>: a complete, ready-to-use copy of your Caseload
            file with the new rows already added &mdash; rename it to replace your original.
            Every other tab is left byte-for-byte untouched.
            <br />
            <strong>New rows only</strong>: just the new rows, for pasting in yourself or keeping
            as a record of what was added.
          </p>
        </section>
      )}

      <footer className="mt-16 border-t border-line pt-4 text-xs text-ink-soft">
        Matches on Status = APVD and Fund containing OSY or ISY (covers variants like A/OSY).
      </footer>
    </div>
  );
}
