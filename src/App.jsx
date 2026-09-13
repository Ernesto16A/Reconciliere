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
// profile's columnMappings. No auto-detection — headers must match exactly
// (with an optional alternate header as a fallback for blank cells).
function parseSheet(json, columnMappings, side) {
  const altField = side === "client" ? "clientAlt" : "factoryAlt";
  const getMapping = (key) => columnMappings.find((m) => m.key === key);

  const deliveryMap = getMapping("delivery");
  const boxMap = getMapping("box");
  const qtyMap = getMapping("qty");
  const customMappings = columnMappings.filter((m) => m.key === "custom" && m.label.trim());

  return json
    .map((row) => {
      const keys = Object.keys(row);

      const delivery = deliveryMap ? readField(row, keys, deliveryMap[side], deliveryMap[altField]) : "";
      const qty = qtyMap ? readField(row, keys, qtyMap[side], qtyMap[altField]) : "";
      const box = boxMap ? readField(row, keys, boxMap[side], boxMap[altField]) : "";

      if (!delivery && !qty) return null;

      const extra = {};
      customMappings.forEach((m) => {
        extra[m.label] = readField(row, keys, m[side], m[altField]);
      });

      return { id: Math.random().toString(36).slice(2), delivery, qty, box, extra };
    })
    .filter(Boolean);
}

// Sums quantities for rows sharing the same delivery number AND box number
// (so a repeated DN+Box combination adds up rather than duplicating), while
// keeping every distinct box under a delivery visible rather than picking
// just one.
function aggregateByDeliveryAndBox(rows) {
  const byKey = new Map();

  rows.forEach((r) => {
    const delivery = r.delivery.trim();
    if (!delivery) return;
    const box = (r.box || "").trim();
    const key = `${delivery}|||${box}`;

    if (!byKey.has(key)) {
      byKey.set(key, { delivery, box, qty: 0 });
    }
    byKey.get(key).qty += parseFloat(r.qty) || 0;
  });

  return [...byKey.values()];
}

// Rolls box-level rows up into one summary per delivery number: a total
// quantity (used for match/mismatch comparison) plus the full list of boxes
// that make up that total (used for display, so nothing is hidden).
function summarizeByDelivery(boxLevelRows) {
  const byDelivery = new Map();

  boxLevelRows.forEach((r) => {
    if (!byDelivery.has(r.delivery)) {
      byDelivery.set(r.delivery, { delivery: r.delivery, totalQty: 0, boxes: [] });
    }
    const entry = byDelivery.get(r.delivery);
    entry.totalQty += r.qty;
    entry.boxes.push({ box: r.box, qty: r.qty });
  });

  return byDelivery;
}

function matchByDelivery(clientRows, factoryRows) {
  const clientSummary = summarizeByDelivery(clientRows);
  const factorySummary = summarizeByDelivery(factoryRows);
  const deliveries = new Set([...clientSummary.keys(), ...factorySummary.keys()]);

  return [...deliveries].sort().map((delivery) => {
    const c = clientSummary.get(delivery);
    const f = factorySummary.get(delivery);

    let status;
    if (!f) status = "missing_factory";
    else if (!c) status = "missing_client";
    else if (c.totalQty !== f.totalQty) status = "mismatch";
    else status = "match";

    return { delivery, c, f, status };
  });
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

function downloadResultsAsExcel(results, boxMappings) {
  const data = [];

  results.forEach((r) => {
    const boxes = r.c?.boxes?.length ? r.c.boxes : [{ box: "", qty: "" }];
    boxes.forEach((b) => {
      data.push({
        Type: r.type,
        "Delivery note": r.delivery,
        "Client box number": b.box || "",
        "Client box qty": b.qty || "",
        "Factory box (mapped)": resolveFactoryBox(b.box, boxMappings) || "",
        "Client total qty": r.c?.totalQty ?? "",
        "Factory total qty": r.f?.totalQty ?? "",
        Difference: describeDifference(r.status),
      });
    });
  });

  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Reconciliation");

  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `reconciliation-results-${stamp}.xlsx`);
}

function defaultColumnMappings() {
  return [
    { id: "delivery", key: "delivery", label: "Delivery note", client: "", clientAlt: "", factory: "", factoryAlt: "", removable: false },
    { id: "box", key: "box", label: "Box number", client: "", clientAlt: "", factory: "", factoryAlt: "", removable: false },
    { id: "qty", key: "qty", label: "Quantity", client: "", clientAlt: "", factory: "", factoryAlt: "", removable: false },
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

function BoxList({ boxes, theme, mode, boxMappings }) {
  if (!boxes || boxes.length === 0) return <span className={theme.muted}>—</span>;
  return (
    <div className="flex flex-col gap-0.5">
      {boxes.map((b, i) => (
        <span key={i} className={`font-mono text-xs ${theme.heading}`}>
          {mode === "factory-mapped" ? resolveFactoryBox(b.box, boxMappings) || "—" : `${b.box || "—"} (${b.qty})`}
        </span>
      ))}
    </div>
  );
}

function ResultsTable({ theme, results, boxMappings }) {
  return (
    <div className="mb-6">
      <div className="flex items-center justify-between mb-2">
        <span className={`text-xs ${theme.muted}`}>
          {results.length} difference{results.length === 1 ? "" : "s"}
        </span>
        <button
          onClick={() => downloadResultsAsExcel(results, boxMappings)}
          disabled={results.length === 0}
          className={`text-xs font-medium px-2.5 py-1.5 border flex items-center gap-1.5 ${
            results.length === 0 ? `${theme.tabBorder} ${theme.muted} cursor-not-allowed` : `${theme.tabBorder} ${theme.heading} ${theme.rowHover}`
          }`}
        >
          <Download size={13} /> Download as Excel
        </button>
      </div>
      <div className={`border ${theme.panel}`}>
        <div className={`grid grid-cols-7 gap-3 px-4 py-2.5 border-b ${theme.tabBorder} text-[11px] uppercase tracking-wide ${theme.muted} font-medium`}>
          <span>Type</span>
          <span>Delivery note</span>
          <span>Client boxes (qty)</span>
          <span>Factory boxes (mapped)</span>
          <span>Client total qty</span>
          <span>Factory total qty</span>
          <span>Difference</span>
        </div>
        {results.length === 0 ? (
          <div className={`px-4 py-4 text-sm ${theme.muted}`}>No differences found — everything matches.</div>
        ) : (
          results.map((r) => (
            <div key={`${r.type}-${r.delivery}`} className={`grid grid-cols-7 gap-3 px-4 py-2.5 border-b ${theme.rowBorder} items-start text-sm last:border-b-0`}>
              <span className={`text-xs ${theme.subtle}`}>{r.type}</span>
              <span className={`font-mono ${theme.heading}`}>{r.delivery}</span>
              <BoxList boxes={r.c?.boxes} theme={theme} mode="client" />
              <BoxList boxes={r.c?.boxes} theme={theme} mode="factory-mapped" boxMappings={boxMappings} />
              <span className={`font-mono ${theme.heading}`}>{r.c ? r.c.totalQty : "—"}</span>
              <span className={`font-mono ${theme.heading}`}>{r.f ? r.f.totalQty : "—"}</span>
              <span className={`text-xs font-medium ${theme.danger}`}>{describeDifference(r.status)}</span>
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

  const computeResults = (cIn, fIn, cOut, fOut) => {
    const combined = [];
    if (cIn.length > 0 || fIn.length > 0) combined.push(...matchByDelivery(cIn, fIn).map((r) => ({ ...r, type: "In" })));
    if (cOut.length > 0 || fOut.length > 0) combined.push(...matchByDelivery(cOut, fOut).map((r) => ({ ...r, type: "Out" })));
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
        const rows = aggregateByDeliveryAndBox(parseSheet(json, selectedProfile.columnMappings, side));

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
    setResults(computeResults(clientIn, factoryIn, clientOut, factoryOut));
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
        <ResultsTable theme={theme} results={results.rows} boxMappings={selectedProfile ? selectedProfile.boxMappings : []} />
      )}
    </div>
  );
}

// ---------- Client Profile tab ----------

function ColumnMappingTable({ theme, mappings, setMappings }) {
  const [collapsed, setCollapsed] = useState({});

  const updateRow = (id, field, value) => setMappings(mappings.map((m) => (m.id === id ? { ...m, [field]: value } : m)));
  const removeRow = (id) => setMappings(mappings.filter((m) => m.id !== id));
  const addRow = () =>
    setMappings([
      ...mappings,
      { id: Math.random().toString(36).slice(2), key: "custom", label: "", client: "", clientAlt: "", factory: "", factoryAlt: "", removable: true },
    ]);

  const toggle = (side) => setCollapsed((prev) => ({ ...prev, [side]: !prev[side] }));

  const cols = "1fr 1fr 1fr 28px";

  const renderSide = (side, sideLabel) => {
    const isOpen = !collapsed[side];
    const altField = side === "client" ? "clientAlt" : "factoryAlt";

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
            {mappings.map((m) => (
              <div key={m.id} className="grid gap-2 items-center" style={{ gridTemplateColumns: cols }}>
                {m.removable ? (
                  <input
                    value={m.label}
                    onChange={(e) => updateRow(m.id, "label", e.target.value)}
                    placeholder="e.g. Document date"
                    className={`border px-2 py-1 text-sm ${theme.input}`}
                  />
                ) : (
                  <span className={`text-sm px-2 py-1 ${theme.heading}`}>{m.label}</span>
                )}
                <input
                  value={m[side]}
                  onChange={(e) => updateRow(m.id, side, e.target.value)}
                  placeholder="Export header"
                  className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
                />
                <input
                  value={m[altField]}
                  onChange={(e) => updateRow(m.id, altField, e.target.value)}
                  placeholder="If blank, use…"
                  className={`border px-2 py-1 text-sm font-mono ${theme.input}`}
                />
                {m.removable ? (
                  <button onClick={() => removeRow(m.id)} className={`flex items-center justify-center ${theme.iconMuted}`} aria-label="Remove field">
                    <Trash2 size={14} />
                  </button>
                ) : (
                  <span />
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className={`border ${theme.panel}`}>
      {renderSide("client", "Client")}
      {renderSide("factory", "Factory")}

      <div className={`px-4 py-2 border-t ${theme.rowBorder}`}>
        <button onClick={addRow} className={`text-xs flex items-center gap-1 font-medium ${theme.heading}`}>
          <Plus size={13} /> Add column
        </button>
      </div>
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
    setProfiles([...profiles, { id, name: `New client ${profiles.length + 1}`, columnMappings: defaultColumnMappings(), boxMappings: [] }]);
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
                    Enter the exact column header used in each side's export. Use the "(alt)" field for cases like
                    "Receipt" / "Dispatch" columns where only one is filled per row — if the main header is blank,
                    the alternate is checked instead. Add a column for anything extra this client requires, like
                    document date.
                  </p>
                  <ColumnMappingTable
                    theme={theme}
                    mappings={p.columnMappings}
                    setMappings={(next) => updateProfile(p.id, (pr) => ({ ...pr, columnMappings: next }))}
                  />
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
      if (raw) setProfiles(JSON.parse(raw));
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
