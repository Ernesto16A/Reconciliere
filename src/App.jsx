import { useState, useEffect } from "react";
import { Play, Upload, Plus, Trash2, Pencil, ChevronRight, Moon, Sun, Download } from "lucide-react";
import * as XLSX from "xlsx";

// ---------- Theme ----------

function getTheme(dark) {
  return {
    dark,
    page: dark ? "bg-slate-950 text-slate-100" : "bg-slate-50 text-slate-900",
    panel: dark ? "bg-slate-800 border-slate-700" : "bg-white border-slate-300",
    panelAlt: dark ? "bg-slate-900 border-slate-700" : "bg-slate-50 border-slate-200",
    heading: dark ? "text-slate-100" : "text-slate-900",
    subtle: dark ? "text-slate-400" : "text-slate-500",
    muted: dark ? "text-slate-500" : "text-slate-400",
    input: dark
      ? "bg-slate-900 border-slate-700 text-slate-100 placeholder-slate-600"
      : "bg-white border-slate-300 text-slate-900 placeholder-slate-400",
    rowBorder: dark ? "border-slate-700" : "border-slate-100",
    rowHover: dark ? "hover:bg-slate-700/50" : "hover:bg-slate-50",
    rowActive: dark ? "bg-slate-700/50" : "bg-slate-100",
    tabBorder: dark ? "border-slate-700" : "border-slate-300",
    tabActive: dark ? "border-slate-100 text-slate-100" : "border-slate-900 text-slate-900",
    tabInactive: dark ? "text-slate-500 hover:text-slate-300" : "text-slate-500 hover:text-slate-700",
    buttonPrimary: dark ? "bg-slate-100 text-slate-900 hover:bg-white" : "bg-slate-900 text-white hover:bg-slate-800",
    buttonDisabled: dark ? "bg-slate-800 text-slate-600" : "bg-slate-200 text-slate-400",
    iconMuted: dark ? "text-slate-600 hover:text-slate-300" : "text-slate-300 hover:text-slate-600",
    danger: dark ? "text-red-400" : "text-red-600",
    dangerHover: dark ? "hover:text-red-400" : "hover:text-red-500",
  };
}

// ---------- Parsing ----------

function findKeyExact(keys, exactName) {
  if (!exactName) return null;
  const target = exactName.trim().toLowerCase();
  return keys.find((k) => k.trim().toLowerCase() === target) || null;
}

// Reads a field's value, falling back to an alternate header if the
// primary header is missing/blank for that row (e.g. "Receipt" is filled
// but "Dispatch" is empty, or vice versa).
function readField(row, keys, primaryHeader, altHeader) {
  const primaryKey = findKeyExact(keys, primaryHeader);
  let value = primaryKey ? String(row[primaryKey] ?? "").trim() : "";
  if (!value && altHeader) {
    const altKey = findKeyExact(keys, altHeader);
    if (altKey) value = String(row[altKey] ?? "").trim();
  }
  return value;
}

// Reads delivery/box/qty for one side ("client" or "factory") using the
// profile's columnMappings, plus that side's own extra fields (informational
// only — never used for matching). No auto-detection — headers must match
// exactly (with an optional alternate header as a fallback for blank cells).
function parseSheet(json, columnMappings, extraFields, side) {
  const altField = side === "client" ? "clientAlt" : "factoryAlt";
  const getMapping = (key) => columnMappings.find((m) => m.key === key);

  const deliveryMap = getMapping("delivery");
  const boxMap = getMapping("box");
  const qtyMap = getMapping("qty");
  const labeledExtraFields = extraFields.filter((f) => f.label.trim());

  return json
    .map((row) => {
      const keys = Object.keys(row);

      const delivery = deliveryMap ? readField(row, keys, deliveryMap[side], deliveryMap[altField]) : "";
      const qty = qtyMap ? readField(row, keys, qtyMap[side], qtyMap[altField]) : "";
      const box = boxMap ? readField(row, keys, boxMap[side], boxMap[altField]) : "";

      if (!delivery && !qty) return null;

      const extra = {};
      labeledExtraFields.forEach((f) => {
        extra[f.label] = readField(row, keys, f.header, f.alt);
      });

      return { id: Math.random().toString(36).slice(2), delivery, qty, box, extra };
    })
    .filter(Boolean);
}

// Sums quantities for rows sharing the same delivery number AND box number
// (so a repeated DN+Box combination adds up), while keeping every distinct
// box under a delivery as its own separate entry — different boxes are
// never combined into one blended total.
function aggregateByDeliveryAndBox(rows) {
  const byKey = new Map();

  rows.forEach((r) => {
    const delivery = r.delivery.trim();
    if (!delivery) return;
    const box = (r.box || "").trim();
    const key = `${delivery}|||${box}`;

    if (!byKey.has(key)) {
      byKey.set(key, { delivery, box, qty: 0, extra: r.extra || {} });
    }
    byKey.get(key).qty += parseFloat(r.qty) || 0;
  });

  return [...byKey.values()];
}

// Matches each client box-level row against its factory counterpart, found
// via the box mapping (client box → factory box, or assumed identical if
// unmapped). Comparison happens per (delivery, box) pair — never as a
// delivery-wide total across different boxes.
function matchByDeliveryAndBox(clientRows, factoryRows, boxMappings) {
  const factoryByKey = new Map();
  factoryRows.forEach((r) => factoryByKey.set(`${r.delivery}|||${r.box}`, r));
  const usedFactoryKeys = new Set();

  const results = [];

  clientRows.forEach((c) => {
    const mappedBox = resolveFactoryBox(c.box, boxMappings);
    const key = `${c.delivery}|||${mappedBox}`;
    const f = factoryByKey.get(key);

    let status;
    if (!f) {
      status = "missing_factory";
    } else {
      usedFactoryKeys.add(key);
      status = c.qty === f.qty ? "match" : "mismatch";
    }

    results.push({
      delivery: c.delivery,
      clientBox: c.box,
      clientQty: c.qty,
      clientExtra: c.extra,
      factoryBox: f ? f.box : mappedBox,
      factoryQty: f ? f.qty : null,
      factoryExtra: f ? f.extra : {},
      status,
    });
  });

  factoryRows.forEach((f) => {
    const key = `${f.delivery}|||${f.box}`;
    if (usedFactoryKeys.has(key)) return;
    results.push({
      delivery: f.delivery,
      clientBox: null,
      clientQty: null,
      clientExtra: {},
      factoryBox: f.box,
      factoryQty: f.qty,
      factoryExtra: f.extra,
      status: "missing_client",
    });
  });

  return results;
}

function describeDifference(status) {
  if (status === "mismatch") return "Qty mismatch";
  if (status === "missing_factory") return "Missing on factory side";
  if (status === "missing_client") return "Missing on client side";
  return "";
}

// If no explicit client→factory mapping exists for this box (either because
// the mapping table is empty, or this particular box just isn't in it), we
// assume the box is named the same on both sides rather than hiding it.
function resolveFactoryBox(clientBox, boxMappings) {
  if (!clientBox) return "";
  const match = boxMappings.find((m) => m.clientBox.trim() === clientBox.trim());
  return match ? match.factoryBox : clientBox;
}

function downloadResultsAsExcel(results, profile) {
  const clientExtraLabels = (profile?.clientExtraFields || []).map((f) => f.label).filter((l) => l.trim());
  const factoryExtraLabels = (profile?.factoryExtraFields || []).map((f) => f.label).filter((l) => l.trim());

  const data = results.map((r) => {
    const row = {
      "Delivery number": r.type,
      "Delivery note": r.delivery,
      Box: r.clientBox || r.factoryBox || "",
      "Client qty": r.clientQty ?? "",
      "Factory qty": r.factoryQty ?? "",
      Difference: describeDifference(r.status),
    };
    clientExtraLabels.forEach((label) => {
      row[`Client – ${label}`] = r.clientExtra?.[label] || "";
    });
    factoryExtraLabels.forEach((label) => {
      row[`Factory – ${label}`] = r.factoryExtra?.[label] || "";
    });
    return row;
  });

  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Reconciliation");

  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `reconciliation-results-${stamp}.xlsx`);
}

function defaultColumnMappings() {
  return [
    { id: "delivery", key: "delivery", label: "Delivery note", client: "", clientAlt: "", factory: "", factoryAlt: "" },
    { id: "box", key: "box", label: "Box number", client: "", clientAlt: "", factory: "", factoryAlt: "" },
    { id: "qty", key: "qty", label: "Quantity", client: "", clientAlt: "", factory: "", factoryAlt: "" },
  ];
}

// ---------- Shared bits ----------

function DarkModeSwitch({ dark, onToggle }) {
  return (
    <button
      onClick={onToggle}
      className="flex items-center gap-2 text-xs font-medium"
      style={{ color: dark ? "#CBD5E1" : "#475569" }}
      aria-label="Toggle dark mode"
    >
      {dark ? <Moon size={14} /> : <Sun size={14} />}
      <span
        className="relative inline-flex items-center rounded-full transition-colors"
        style={{ width: 40, height: 22, background: dark ? "#475569" : "#CBD5E1" }}
      >
        <span
          className="absolute rounded-full bg-white transition-transform"
          style={{ width: 16, height: 16, top: 3, left: 3, transform: dark ? "translateX(18px)" : "translateX(0)" }}
        />
      </span>
    </button>
  );
}

function UploadPanel({ theme, title, accent, onFile, fileName, fileError, disabled, disabledHint }) {
  const inputId = `file-${title.replace(/\s+/g, "-")}`;

  return (
    <div className={`flex-1 min-w-0 border ${theme.panel}`}>
      <div className={`px-4 py-2.5 border-b ${theme.tabBorder} flex items-center`} style={{ borderTop: `3px solid ${accent}` }}>
        <span className={`font-semibold text-sm ${theme.heading}`}>{title}</span>
      </div>

      <div className="px-4 py-2.5 flex items-center justify-between gap-3">
        <label
          htmlFor={inputId}
          className={`text-xs font-medium px-2.5 py-1.5 border flex items-center gap-1.5 shrink-0 ${
            disabled
              ? `${theme.tabBorder} ${theme.muted} cursor-not-allowed`
              : `${theme.tabBorder} ${theme.heading} ${theme.rowHover} cursor-pointer`
          }`}
        >
          <Upload size={13} /> Upload Excel
        </label>
        <input id={inputId} type="file" accept=".xlsx,.xls,.csv" onChange={onFile} disabled={disabled} className="hidden" />
        <span className={`text-xs truncate ${theme.muted}`}>
          {disabled ? disabledHint : fileError ? <span className={theme.danger}>{fileError}</span> : fileName || "No file uploaded"}
        </span>
      </div>
    </div>
  );
}

function ResultsTable({ theme, results, profile }) {
  const clientExtraFields = (profile?.clientExtraFields || []).filter((f) => f.label.trim());
  const factoryExtraFields = (profile?.factoryExtraFields || []).filter((f) => f.label.trim());
  const extraCount = clientExtraFields.length + factoryExtraFields.length;
  const gridStyle = { gridTemplateColumns: `repeat(6, 1fr) ${extraCount ? `repeat(${extraCount}, 1fr)` : ""}`.trim() };

  return (
    <div className="mb-6">
      <div className="flex items-center justify-between mb-2">
        <span className={`text-xs ${theme.muted}`}>
          {results.length} difference{results.length === 1 ? "" : "s"}
        </span>
        <button
          onClick={() => downloadResultsAsExcel(results, profile)}
          disabled={results.length === 0}
          className={`text-xs font-medium px-2.5 py-1.5 border flex items-center gap-1.5 ${
            results.length === 0 ? `${theme.tabBorder} ${theme.muted} cursor-not-allowed` : `${theme.tabBorder} ${theme.heading} ${theme.rowHover}`
          }`}
        >
          <Download size={13} /> Download as Excel
        </button>
      </div>
      <div className={`border ${theme.panel} overflow-x-auto`}>
        <div className={`grid gap-3 px-4 py-2.5 border-b ${theme.tabBorder} text-[11px] uppercase tracking-wide ${theme.muted} font-medium`} style={gridStyle}>
          {/* "Delivery number" is a placeholder header for now — several delivery
              notes may share one delivery number, to be wired up later. */}
          <span>Delivery number</span>
          <span>Delivery note</span>
          <span>Box</span>
          <span>Client qty</span>
          <span>Factory qty</span>
          <span>Difference</span>
          {clientExtraFields.map((f) => (
            <span key={f.id}>Client: {f.label}</span>
          ))}
          {factoryExtraFields.map((f) => (
            <span key={f.id}>Factory: {f.label}</span>
          ))}
        </div>
        {results.length === 0 ? (
          <div className={`px-4 py-4 text-sm ${theme.muted}`}>No differences found — everything matches.</div>
        ) : (
          results.map((r, i) => (
            <div key={i} className={`grid gap-3 px-4 py-2.5 border-b ${theme.rowBorder} items-center text-sm last:border-b-0`} style={gridStyle}>
              <span className={`text-xs ${theme.subtle}`}>{r.type}</span>
              <span className={`font-mono ${theme.heading}`}>{r.delivery}</span>
              <span className={`font-mono ${theme.heading}`}>{r.clientBox || r.factoryBox || "—"}</span>
              <span className={`font-mono ${theme.heading}`}>{r.clientQty ?? "—"}</span>
              <span className={`font-mono ${theme.heading}`}>{r.factoryQty ?? "—"}</span>
              <span className={`text-xs font-medium ${theme.danger}`}>{describeDifference(r.status)}</span>
              {clientExtraFields.map((f) => (
                <span key={f.id} className={`text-sm ${theme.subtle}`}>
                  {r.clientExtra?.[f.label] || "—"}
                </span>
              ))}
              {factoryExtraFields.map((f) => (
                <span key={f.id} className={`text-sm ${theme.subtle}`}>
                  {r.factoryExtra?.[f.label] || "—"}
                </span>
              ))}
            </div>
          ))
        )}
      </div>
    </div>
  );
}

// ---------- Reconciliation tab ----------

function ReconciliationTab({ theme, profiles, selectedProfileId, setSelectedProfileId }) {
  const [clientIn, setClientIn] = useState([]);
  const [factoryIn, setFactoryIn] = useState([]);
  const [clientOut, setClientOut] = useState([]);
  const [factoryOut, setFactoryOut] = useState([]);

  const [files, setFiles] = useState({
    clientIn: { name: "", error: "" },
    factoryIn: { name: "", error: "" },
    clientOut: { name: "", error: "" },
    factoryOut: { name: "", error: "" },
  });

  const [results, setResults] = useState(null);

  const selectedProfile = profiles.find((p) => p.id === selectedProfileId) || null;

  const computeResults = (cIn, fIn, cOut, fOut, boxMappings) => {
    const combined = [];
    if (cIn.length > 0 || fIn.length > 0) combined.push(...matchByDeliveryAndBox(cIn, fIn, boxMappings).map((r) => ({ ...r, type: "In" })));
    if (cOut.length > 0 || fOut.length > 0) combined.push(...matchByDeliveryAndBox(cOut, fOut, boxMappings).map((r) => ({ ...r, type: "Out" })));
    return { ranAnySection: true, rows: combined.filter((r) => r.status !== "match") };
  };

  const handleFile = (e, setRows, key, side) => {
    const file = e.target.files?.[0];
    if (!file || !selectedProfile) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const extraFields = side === "client" ? selectedProfile.clientExtraFields : selectedProfile.factoryExtraFields;
        const rows = aggregateByDeliveryAndBox(parseSheet(json, selectedProfile.columnMappings, extraFields, side));

        if (rows.length === 0) {
          setFiles((f) => ({ ...f, [key]: { name: file.name, error: "Columns don't match this profile's mapping" } }));
          return;
        }

        setRows(rows);
        setFiles((f) => ({ ...f, [key]: { name: `${file.name} (${rows.length} rows)`, error: "" } }));
        setResults(null);
      } catch (err) {
        setFiles((f) => ({ ...f, [key]: { name: file.name, error: "Couldn't read this file" } }));
      }
    };
    reader.onerror = () => setFiles((f) => ({ ...f, [key]: { name: file.name, error: "Couldn't read this file" } }));
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const anyUploaded = clientIn.length > 0 || factoryIn.length > 0 || clientOut.length > 0 || factoryOut.length > 0;

  const runReconciliation = () => {
    setResults(computeResults(clientIn, factoryIn, clientOut, factoryOut, selectedProfile ? selectedProfile.boxMappings : []));
  };

  return (
    <div>
      <div className={`mb-5 border ${theme.panel} px-4 py-3 flex items-center gap-3`}>
        <span className={`text-xs font-medium uppercase tracking-wide shrink-0 ${theme.subtle}`}>Client profile</span>
        {profiles.length === 0 ? (
          <span className={`text-sm ${theme.muted}`}>No client profiles yet — create one in the Client Profile tab.</span>
        ) : (
          <select
            value={selectedProfileId || ""}
            onChange={(e) => setSelectedProfileId(e.target.value)}
            className={`border px-2 py-1.5 text-sm flex-1 max-w-xs ${theme.input}`}
          >
            {profiles.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
              </option>
            ))}
          </select>
        )}
        {selectedProfile && (
          <span className={`text-xs ${theme.muted}`}>
            {selectedProfile.boxMappings.length} box mapping{selectedProfile.boxMappings.length === 1 ? "" : "s"}
          </span>
        )}
      </div>

      <div className="mb-2">
        <h2 className={`text-sm font-semibold mb-3 uppercase tracking-wide ${theme.heading}`}>In</h2>
        <div className="flex flex-col md:flex-row gap-4">
          <UploadPanel
            theme={theme}
            title="Client — In"
            accent="#2B6CB0"
            onFile={(e) => handleFile(e, setClientIn, "clientIn", "client")}
            fileName={files.clientIn.name}
            fileError={files.clientIn.error}
            disabled={!selectedProfile}
            disabledHint="Select a client profile first"
          />
          <UploadPanel
            theme={theme}
            title="Factory — In"
            accent="#B05A1E"
            onFile={(e) => handleFile(e, setFactoryIn, "factoryIn", "factory")}
            fileName={files.factoryIn.name}
            fileError={files.factoryIn.error}
            disabled={!selectedProfile}
            disabledHint="Select a client profile first"
          />
        </div>
      </div>

      <div className="mb-6 mt-6">
        <h2 className={`text-sm font-semibold mb-3 uppercase tracking-wide ${theme.heading}`}>Out</h2>
        <div className="flex flex-col md:flex-row gap-4">
          <UploadPanel
            theme={theme}
            title="Client — Out"
            accent="#2B6CB0"
            onFile={(e) => handleFile(e, setClientOut, "clientOut", "client")}
            fileName={files.clientOut.name}
            fileError={files.clientOut.error}
            disabled={!selectedProfile}
            disabledHint="Select a client profile first"
          />
          <UploadPanel
            theme={theme}
            title="Factory — Out"
            accent="#B05A1E"
            onFile={(e) => handleFile(e, setFactoryOut, "factoryOut", "factory")}
            fileName={files.factoryOut.name}
            fileError={files.factoryOut.error}
            disabled={!selectedProfile}
            disabledHint="Select a client profile first"
          />
        </div>
      </div>

      <div className="mb-6">
        <button
          onClick={runReconciliation}
          disabled={!anyUploaded}
          className={`text-sm font-medium px-4 py-2 flex items-center gap-2 ${anyUploaded ? theme.buttonPrimary : `${theme.buttonDisabled} cursor-not-allowed`}`}
        >
          <Play size={14} /> Run reconciliation
        </button>
        {!anyUploaded && <p className={`text-xs mt-2 ${theme.muted}`}>Upload at least one file above to enable this.</p>}
      </div>

      {results && results.ranAnySection && (
        <ResultsTable theme={theme} results={results.rows} profile={selectedProfile} />
      )}
    </div>
  );
}

// ---------- Client Profile tab ----------

function ColumnMappingTable({ theme, profile, updateProfile }) {
  const [collapsed, setCollapsed] = useState({});
  const toggle = (side) => setCollapsed((prev) => ({ ...prev, [side]: !prev[side] }));
  const cols = "1fr 1fr 1fr 28px";
  const extraFieldsKey = (side) => (side === "client" ? "clientExtraFields" : "factoryExtraFields");

  const updateCoreRow = (fieldId, prop, value) =>
    updateProfile(profile.id, (p) => ({
      ...p,
      columnMappings: p.columnMappings.map((m) => (m.id === fieldId ? { ...m, [prop]: value } : m)),
    }));

  const updateExtraField = (side, fieldId, prop, value) =>
    updateProfile(profile.id, (p) => ({
      ...p,
      [extraFieldsKey(side)]: p[extraFieldsKey(side)].map((f) => (f.id === fieldId ? { ...f, [prop]: value } : f)),
    }));

  const removeExtraField = (side, fieldId) =>
    updateProfile(profile.id, (p) => ({
      ...p,
      [extraFieldsKey(side)]: p[extraFieldsKey(side)].filter((f) => f.id !== fieldId),
    }));

  const addExtraField = (side) =>
    updateProfile(profile.id, (p) => ({
      ...p,
      [extraFieldsKey(side)]: [...p[extraFieldsKey(side)], { id: Math.random().toString(36).slice(2), label: "", header: "", alt: "" }],
    }));

  const renderSide = (side, sideLabel) => {
    const isOpen = !collapsed[side];
    const altField = side === "client" ? "clientAlt" : "factoryAlt";
    const extraFields = profile[extraFieldsKey(side)];

    return (
      <div className={`border-b ${theme.rowBorder} last:border-b-0`}>
        <button onClick={() => toggle(side)} className="flex items-center gap-1.5 px-4 py-2.5 w-full text-left">
          <ChevronRight size={14} className={`transition-transform ${isOpen ? "rotate-90" : ""} ${theme.muted}`} />
          <span className={`text-sm font-semibold ${theme.heading}`}>{sideLabel}</span>
        </button>

        {isOpen && (
          <div className="px-4 pb-3 flex flex-col gap-1.5">
            <div className={`grid gap-2 text-[10px] uppercase tracking-wide font-medium ${theme.muted}`} style={{ gridTemplateColumns: cols }}>
              <span>Field</span>
              <span>Header from export</span>
              <span>Alternative</span>
              <span />
            </div>

            {profile.columnMappings.map((m) => (
              <div key={m.id} className="grid gap-2 items-center" style={{ gridTemplateColumns: cols }}>
                <span className={`text-sm px-2 py-1 ${theme.heading}`}>{m.label}</span>
                <input
                  value={m[side]}
                  onChange={(e) => updateCoreRow(m.id, side, e.target.value)}
                  placeholder="Export header"
                  className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
                />
                <input
                  value={m[altField]}
                  onChange={(e) => updateCoreRow(m.id, altField, e.target.value)}
                  placeholder="If blank, use…"
                  className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
                />
                <span />
              </div>
            ))}

            {extraFields.length > 0 && <div className={`h-px my-1 ${theme.rowBorder} border-t`} />}

            {extraFields.map((f, idx) => {
              const isLast = idx === extraFields.length - 1;
              return (
                <div key={f.id} className="grid gap-2 items-center" style={{ gridTemplateColumns: cols }}>
                  {isLast ? (
                    <div className="relative">
                      <input
                        value={f.label}
                        onChange={(e) => updateExtraField(side, f.id, "label", e.target.value)}
                        placeholder="e.g. Document date"
                        className={`w-full border px-2 py-1 text-sm ${theme.input}`}
                        style={{ paddingRight: 28 }}
                      />
                      <button
                        onClick={() => addExtraField(side)}
                        className={`absolute ${theme.iconMuted}`}
                        style={{ right: 6, top: "50%", transform: "translateY(-50%)" }}
                        aria-label="Add another column"
                      >
                        <Plus size={14} />
                      </button>
                    </div>
                  ) : (
                    <span className={`text-sm px-2 py-1 ${theme.heading}`}>{f.label || "—"}</span>
                  )}
                  <input
                    value={f.header}
                    onChange={(e) => updateExtraField(side, f.id, "header", e.target.value)}
                    placeholder="Export header"
                    className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
                  />
                  <input
                    value={f.alt}
                    onChange={(e) => updateExtraField(side, f.id, "alt", e.target.value)}
                    placeholder="If blank, use…"
                    className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
                  />
                  <button onClick={() => removeExtraField(side, f.id)} className={`flex items-center justify-center ${theme.iconMuted}`} aria-label="Remove field">
                    <Trash2 size={14} />
                  </button>
                </div>
              );
            })}

            {extraFields.length === 0 && (
              <button onClick={() => addExtraField(side)} className={`text-xs flex items-center gap-1 font-medium mt-1 ${theme.heading}`}>
                <Plus size={13} /> Add column
              </button>
            )}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className={`border ${theme.panel}`}>
      {renderSide("client", "Client")}
      {renderSide("factory", "Factory")}
    </div>
  );
}

function BoxMappingTable({ theme, mappings, setMappings }) {
  const update = (id, field, value) => setMappings(mappings.map((m) => (m.id === id ? { ...m, [field]: value } : m)));
  const remove = (id) => setMappings(mappings.filter((m) => m.id !== id));
  const add = () => setMappings([...mappings, { id: Math.random().toString(36).slice(2), clientBox: "", factoryBox: "" }]);

  return (
    <div className={`border ${theme.panel}`}>
      <div
        className={`grid gap-2 px-4 py-2 border-b ${theme.tabBorder} text-[11px] uppercase tracking-wide font-medium ${theme.muted}`}
        style={{ gridTemplateColumns: "1fr 1fr 28px" }}
      >
        <span>Client box number</span>
        <span>Factory box number</span>
        <span />
      </div>
      <div className="px-4 py-2 flex flex-col gap-1.5">
        {mappings.length === 0 && <p className={`text-sm py-1 ${theme.muted}`}>No box mappings yet.</p>}
        {mappings.map((m) => (
          <div key={m.id} className="grid gap-2" style={{ gridTemplateColumns: "1fr 1fr 28px" }}>
            <input
              value={m.clientBox}
              onChange={(e) => update(m.id, "clientBox", e.target.value)}
              placeholder="C-102"
              className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
            />
            <input
              value={m.factoryBox}
              onChange={(e) => update(m.id, "factoryBox", e.target.value)}
              placeholder="F-8891"
              className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
            />
            <button onClick={() => remove(m.id)} className={`flex items-center justify-center ${theme.iconMuted}`} aria-label="Remove mapping">
              <Trash2 size={14} />
            </button>
          </div>
        ))}
      </div>
      <div className={`px-4 py-2 border-t ${theme.rowBorder}`}>
        <button onClick={add} className={`text-xs flex items-center gap-1 font-medium ${theme.heading}`}>
          <Plus size={13} /> Add mapping
        </button>
      </div>
    </div>
  );
}

function ClientProfileTab({ theme, profiles, setProfiles, selectedProfileId, setSelectedProfileId }) {
  const [expandedId, setExpandedId] = useState(null);
  const [renamingId, setRenamingId] = useState(null);

  const updateProfile = (id, updater) => setProfiles(profiles.map((p) => (p.id === id ? updater(p) : p)));
  const toggleExpand = (id) => setExpandedId((current) => (current === id ? null : id));

  const addProfile = () => {
    const id = Math.random().toString(36).slice(2);
    setProfiles([
      ...profiles,
      {
        id,
        name: `New client ${profiles.length + 1}`,
        columnMappings: defaultColumnMappings(),
        clientExtraFields: [],
        factoryExtraFields: [],
        boxMappings: [],
      },
    ]);
    setExpandedId(id);
    setRenamingId(id);
    setSelectedProfileId(id);
  };

  const removeProfile = (id) => {
    setProfiles(profiles.filter((p) => p.id !== id));
    if (expandedId === id) setExpandedId(null);
    if (selectedProfileId === id) setSelectedProfileId(null);
  };

  return (
    <div className={`border ${theme.panel}`}>
      <div className={`px-4 py-2.5 border-b ${theme.tabBorder} text-xs font-medium uppercase tracking-wide ${theme.subtle}`}>
        Clients
      </div>

      {profiles.length === 0 && <p className={`px-4 py-4 text-sm ${theme.muted}`}>No client profiles yet.</p>}

      {profiles.map((p) => {
        const expanded = expandedId === p.id;
        return (
          <div key={p.id} className={`border-b ${theme.rowBorder} last:border-b-0`}>
            <div onClick={() => toggleExpand(p.id)} className={`flex items-center justify-between gap-2 px-4 py-2.5 cursor-pointer ${expanded ? theme.rowActive : theme.rowHover}`}>
              <div className="flex items-center gap-2 flex-1 min-w-0">
                <ChevronRight size={14} className={`transition-transform shrink-0 ${expanded ? `rotate-90 ${theme.heading}` : theme.muted}`} />
                {renamingId === p.id ? (
                  <input
                    autoFocus
                    value={p.name}
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => updateProfile(p.id, (pr) => ({ ...pr, name: e.target.value }))}
                    onBlur={() => setRenamingId(null)}
                    onKeyDown={(e) => e.key === "Enter" && setRenamingId(null)}
                    className={`border px-1.5 py-0.5 text-sm flex-1 min-w-0 ${theme.input}`}
                  />
                ) : (
                  <span className={`text-sm truncate ${expanded ? `font-medium ${theme.heading}` : theme.heading}`}>{p.name}</span>
                )}
              </div>
              <div className="flex items-center gap-0.5 shrink-0">
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setExpandedId(p.id);
                    setRenamingId(p.id);
                  }}
                  className={`p-1 ${theme.iconMuted}`}
                  aria-label="Rename"
                >
                  <Pencil size={13} />
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    removeProfile(p.id);
                  }}
                  className={`p-1 ${theme.iconMuted} ${theme.dangerHover}`}
                  aria-label="Delete"
                >
                  <Trash2 size={13} />
                </button>
              </div>
            </div>

            {expanded && (
              <div className={`px-4 pb-4 pt-1 flex flex-col gap-4 ${theme.panelAlt}`}>
                <div>
                  <h3 className={`text-sm font-semibold mb-1 ${theme.heading}`}>Export column mapping</h3>
                  <p className={`text-xs mb-3 ${theme.subtle}`}>
                    Enter the exact column header used in each side's export. Use "Alternative" for cases like
                    "Receipt" / "Dispatch" columns where only one is filled per row — if the main header is blank,
                    the alternate is checked instead. "Add column" under each side captures extra info (like a
                    document date) that appears in the exported Excel but never affects matching.
                  </p>
                  <ColumnMappingTable theme={theme} profile={p} updateProfile={updateProfile} />
                </div>

                <div>
                  <h3 className={`text-sm font-semibold mb-1 ${theme.heading}`}>Box number mapping</h3>
                  <p className={`text-xs mb-3 ${theme.subtle}`}>
                    This client's box numbers don't match the factory's. Map each one to its factory equivalent.
                  </p>
                  <BoxMappingTable theme={theme} mappings={p.boxMappings} setMappings={(next) => updateProfile(p.id, (pr) => ({ ...pr, boxMappings: next }))} />
                </div>
              </div>
            )}
          </div>
        );
      })}

      <div className="px-4 py-2 border-t border-slate-200/20 flex items-center justify-between">
        <button onClick={addProfile} className={`text-xs flex items-center gap-1 font-medium ${theme.heading}`}>
          <Plus size={13} /> Add client
        </button>
        {profiles.length > 0 && (
          <button
            onClick={() => {
              if (window.confirm("Clear all saved client profiles? This can't be undone.")) {
                setProfiles([]);
                setExpandedId(null);
              }
            }}
            className={`text-xs ${theme.muted} ${theme.dangerHover}`}
          >
            Clear saved profiles
          </button>
        )}
      </div>
    </div>
  );
}

// ---------- App ----------

export default function ReconciliationApp() {
  const [darkMode, setDarkMode] = useState(true);
  const [activeTab, setActiveTab] = useState("reconciliation");
  const [profiles, setProfiles] = useState([]);
  const [profilesLoaded, setProfilesLoaded] = useState(false);
  const [selectedProfileId, setSelectedProfileId] = useState(null);

  const theme = getTheme(darkMode);

  const STORAGE_KEY = "reconciliation-client-profiles";

  // Load saved client profiles once on mount, from this browser's localStorage.
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const loaded = JSON.parse(raw).map((p) => ({
          ...p,
          columnMappings: (p.columnMappings || []).map(({ removable, ...rest }) => rest),
          clientExtraFields: p.clientExtraFields || [],
          factoryExtraFields: p.factoryExtraFields || [],
          boxMappings: p.boxMappings || [],
        }));
        setProfiles(loaded);
      }
    } catch (err) {
      // No saved profiles yet, or storage unavailable — start empty.
    } finally {
      setProfilesLoaded(true);
    }
  }, []);

  // Save whenever profiles change, but not before the initial load finishes
  // (otherwise the empty starting state would overwrite saved data).
  useEffect(() => {
    if (!profilesLoaded) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(profiles));
    } catch (err) {
      console.error("Failed to save client profiles", err);
    }
  }, [profiles, profilesLoaded]);

  useEffect(() => {
    if (profiles.length === 0) {
      if (selectedProfileId !== null) setSelectedProfileId(null);
      return;
    }
    if (!profiles.some((p) => p.id === selectedProfileId)) {
      setSelectedProfileId(profiles[0].id);
    }
  }, [profiles, selectedProfileId]);

  return (
    <div className={`w-full min-h-screen px-6 py-6 ${theme.page}`} style={{ fontFamily: "Inter, ui-sans-serif, system-ui" }}>
      <div className="max-w-5xl mx-auto">
        <div className="mb-5 flex items-start justify-between gap-4">
          <div>
            <h1 className={`text-xl font-semibold ${theme.heading}`}>Packaging reconciliation</h1>
            <p className={`text-sm mt-0.5 ${theme.subtle}`}>
              Comparing inbound and outbound deliveries — client records against factory records, matched by delivery number.
            </p>
          </div>
          <DarkModeSwitch dark={darkMode} onToggle={() => setDarkMode((d) => !d)} />
        </div>

        <div className={`mb-6 border-b flex gap-6 ${theme.tabBorder}`}>
          {[
            { id: "reconciliation", label: "Reconciliation" },
            { id: "profiles", label: "Client Profile" },
          ].map((tab) => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`pb-2.5 text-sm font-medium border-b-2 -mb-px ${activeTab === tab.id ? theme.tabActive : `border-transparent ${theme.tabInactive}`}`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {activeTab === "reconciliation" ? (
          <ReconciliationTab
            theme={theme}
            profiles={profiles}
            selectedProfileId={selectedProfileId}
            setSelectedProfileId={setSelectedProfileId}
          />
        ) : (
          <ClientProfileTab
            theme={theme}
            profiles={profiles}
            setProfiles={setProfiles}
            selectedProfileId={selectedProfileId}
            setSelectedProfileId={setSelectedProfileId}
          />
        )}
      </div>
    </div>
  );
}
