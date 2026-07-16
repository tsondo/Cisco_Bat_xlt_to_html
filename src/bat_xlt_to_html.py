#!/usr/bin/env python3
"""
bat_xlt_to_html.py — Convert a Cisco BAT (Bulk Administration Tool) bat.xlt
Excel template into a single self-contained HTML/JS application that runs
unprivileged in any browser (no macros, no ActiveX, no network access).

The generated HTML replicates the xlt's core function:
  * One tab per BAT transaction sheet, with the same columns and
    (Type[maxlen] MANDATORY/OPTIONAL) constraints parsed from row 1.
  * "Create File Format" builder for the dynamic tabs (Phones,
    Phones-Users, User Device Profiles, Remote Destination Profile,
    VG gateways) driven by the xlt's hidden "Database" sheet.
  * "Export to BAT Format" that produces the same comma-separated .txt
    (CRLF line endings, normalized uppercase header line) that the
    VBA macro's frmFileName routine writes.

Usage:
    python3 bat_xlt_to_html.py input.xlt [output.html]

Requires: xlrd  (pip install xlrd)
"""

import json
import re
import sys
import datetime
from pathlib import Path

try:
    import xlrd
except ImportError:
    sys.exit("This script requires the 'xlrd' package:  pip install xlrd")

# ---------------------------------------------------------------------------
# Sheets that are internal to the workbook, never shown to the user
INTERNAL_SHEETS = {"Database", "Xlt Help"}

# Dynamic tabs -> transaction key used to filter the Database sheet
# (mirrors batClass.ArrayListValue assignments in the VBA)
DYNAMIC_SHEETS = {
    "Phones": "Phone",
    "Phones-Users": "Phone",
    "User Device Profiles": "UDP",
    "Remote Destination Profile": "RDP",
    "VG224": "Gateway",
    "VG202-204": "Gateway",
    "VG310": "Gateway",
    "VG320": "Gateway",
    "VG350": "Gateway",
    "VG450": "Gateway",
    "ISR4461": "Gateway",
    "VG420": "Gateway",
}

# Tabs that offer the "Create Dummy MAC Address" option (per the VBA export)
DUMMY_MAC_SHEETS = ["Phones", "Phones-Users", "CTI Port", "CTI Port-Users"]

# Repeating-group templates hardcoded in the VBA (AddSingleSDSet,
# AddSingleBLFSDSet, AddSingleBLFDCPSet, AddSingleRDSet).  These are stable
# across bat.xlt versions because they are compiled into the macro, not the
# Database sheet.
REPEAT_GROUPS = {
    "speedDial": {
        "label": "Speed Dials",
        "sheets": ["Phones", "Phones-Users", "User Device Profiles"],
        "fields": [
            {"name": "Speed Dial Number", "type": "Integer", "len": 255, "req": False},
            {"name": "Speed Dial Label", "type": "String", "len": 30, "req": False},
            {"name": "Speed Dial Label Ascii", "type": "String", "len": 30, "req": False},
        ],
    },
    "blfSpeedDial": {
        "label": "BLF Speed Dials",
        "sheets": ["Phones", "Phones-Users", "User Device Profiles"],
        "fields": [
            {"name": "Busy Lamp Field Destination", "type": "String", "len": 255, "req": False},
            {"name": "Busy Lamp Field Directory Number", "type": "Integer", "len": 50, "req": False},
            {"name": "Busy Lamp Field Label", "type": "String", "len": 30, "req": False},
            {"name": "Busy Lamp Field Label ASCII", "type": "String", "len": 30, "req": False},
            {"name": "Busy Lamp Field Call Pickup", "type": "Boolean", "len": 1, "req": False},
        ],
    },
    "blfDcp": {
        "label": "BLF Directed Call Parks",
        "sheets": ["Phones", "Phones-Users", "User Device Profiles"],
        "fields": [
            {"name": "BLF Directed Call Park Directory Number", "type": "Integer", "len": 50, "req": False},
            {"name": "BLF Directed Call Park Label", "type": "String", "len": 30, "req": False},
            {"name": "BLF Directed Call Park Label ASCII", "type": "String", "len": 30, "req": False},
        ],
    },
    "remoteDest": {
        "label": "Remote Destinations",
        "sheets": ["Remote Destination Profile"],
        "fields": [
            {"name": "Remote Destination Name", "type": "String", "len": 50, "req": False},
            {"name": "Remote Destination Number", "type": "Integer", "len": 50, "req": False},
            {"name": "Time of Day Access", "type": "String", "len": 50, "req": False},
            {"name": "Time Zone", "type": "String", "len": 50, "req": False},
            {"name": "Enable Mobile Connect", "type": "Boolean", "len": 1, "req": False},
            {"name": "Answer Too Soon Timer", "type": "Integer", "len": 50, "req": False},
            {"name": "Answer Too Late Timer", "type": "Integer", "len": 50, "req": False},
            {"name": "Delay Before Ringing Timer", "type": "Integer", "len": 50, "req": False},
            {"name": "Is Mobile Phone", "type": "Boolean", "len": 1, "req": False},
        ],
    },
}

# Default line-field selection when the user hasn't customized the format
# (mirrors arrFieldName defaults in AddFlexiLineSet)
DEFAULT_LINE_FIELDS = [
    "Directory Number", "Display", "Line Text Label",
    "Forward Busy Destination", "Forward No Answer",
    "Call Pickup Group",
]

HDR_RE = re.compile(
    r"^\s*[\(\[]\s*"
    r"(?P<type>String|Integer|Number|Boolean)\s*"
    r"\[?\s*(?P<len>[^\]\)]*?)\s*\]?\s*"
    r"(?P<req>MANDATORY|OPTIONAL)\s*"
    r"[\)\]]",
    re.IGNORECASE,
)


def parse_header_cell(text):
    """Parse 'Field Name\\n(String[64] MANDATORY)' into a column dict."""
    text = str(text).replace("\r", "")
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    if not lines:
        return None
    name = lines[0]
    col = {"name": name, "type": "String", "len": None, "req": False}
    rest = " ".join(lines[1:])
    # annotation may be inline with the name (single-line headers)
    inline = re.search(r"[\(\[]\s*(?:String|Integer|Number|Boolean)\s*\[[^\]]*\]\s*(?:MANDATORY|OPTIONAL)\s*[\)\]]\s*$", name, re.I)
    if inline and not rest:
        rest = name[inline.start():]
        name = name[:inline.start()].strip()
        col["name"] = name
    m = HDR_RE.match(rest)
    if m:
        t = m.group("type").capitalize()
        col["type"] = "Number" if t == "Number" else t
        raw_len = m.group("len").strip()
        # lengths like "12/50" mean MAC(12) or device name(50) -> use max
        nums = re.findall(r"\d+", raw_len)
        col["len"] = max(int(n) for n in nums) if nums else None
        col["req"] = m.group("req").upper() == "MANDATORY"
    return col


def read_workbook(path):
    wb = xlrd.open_workbook(path)
    model = {
        "source": Path(path).name,
        "generated": datetime.date.today().isoformat(),
        "sheets": [],
        "database": [],
        "help": [],
        "dynamic": DYNAMIC_SHEETS,
        "dummyMacSheets": DUMMY_MAC_SHEETS,
        "repeatGroups": REPEAT_GROUPS,
        "defaultLineFields": DEFAULT_LINE_FIELDS,
    }

    for name in wb.sheet_names():
        if name in INTERNAL_SHEETS:
            continue
        sh = wb.sheet_by_name(name)
        cols = []
        for c in range(sh.ncols):
            cell = sh.cell_value(0, c)
            if str(cell).strip() == "":
                continue
            col = parse_header_cell(cell)
            if col:
                cols.append(col)
        if cols:
            model["sheets"].append({"name": name, "columns": cols})

    # Hidden Database sheet: ENUM, DISPLAYNAME, LENGTH, TYPE, NECESSITY,
    # LIST (Device/Line/Intercom), OPERATION (Phone:UDP:RDP:Gateway)
    if "Database" in wb.sheet_names():
        db = wb.sheet_by_name("Database")
        for r in range(1, db.nrows):
            nm = str(db.cell_value(r, 1)).strip()
            if not nm:
                continue
            try:
                ln = int(float(db.cell_value(r, 2)))
            except (ValueError, TypeError):
                ln = None
            model["database"].append({
                "name": nm,
                "len": ln,
                "type": str(db.cell_value(r, 3)).strip() or "String",
                "req": str(db.cell_value(r, 4)).strip().upper() == "MANDATORY",
                "cat": str(db.cell_value(r, 5)).strip(),        # Device/Line/Intercom
                "ops": [o for o in str(db.cell_value(r, 6)).strip().split(":") if o],
            })

    if "Xlt Help" in wb.sheet_names():
        hp = wb.sheet_by_name("Xlt Help")
        for r in range(hp.nrows):
            v = str(hp.cell_value(r, 0)).strip()
            if v:
                model["help"].append(v)

    return model


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>BAT File Builder — __SOURCE__</title>
<style>
:root{
  --bg:#eef1f4; --panel:#ffffff; --ink:#1a2330; --muted:#5b6b7d;
  --rail:#122033; --rail-ink:#c7d4e2; --rail-hi:#203954;
  --accent:#0d6efd; --accent-ink:#ffffff; --ok:#1a7f4b;
  --err:#c62828; --err-bg:#fdecea; --req:#8a5300; --line:#d5dce4;
  --mono:ui-monospace,'Cascadia Mono','Consolas',Menlo,monospace;
  --sans:system-ui,'Segoe UI',Roboto,Arial,sans-serif;
}
*{box-sizing:border-box}
html,body{margin:0;height:100%;font-family:var(--sans);color:var(--ink);background:var(--bg)}
body{display:flex;overflow:hidden}

/* ── left rail: transaction list ─────────────────────────── */
#rail{width:250px;min-width:250px;background:var(--rail);color:var(--rail-ink);
  display:flex;flex-direction:column;height:100vh}
#rail h1{font-size:14px;letter-spacing:.08em;text-transform:uppercase;margin:0;
  padding:14px 14px 4px;color:#fff}
#rail .sub{font-size:11px;color:#7f93a8;padding:0 14px 10px;font-family:var(--mono)}
#railSearch{margin:0 10px 8px;padding:6px 8px;border:1px solid #2c4258;border-radius:4px;
  background:#0c1826;color:#e8eef4;font-size:12px}
#railSearch:focus{outline:2px solid var(--accent)}
#tabs{overflow-y:auto;flex:1;padding-bottom:12px}
#tabs button{display:block;width:100%;text-align:left;background:none;border:0;
  color:var(--rail-ink);padding:7px 14px;font-size:12.5px;cursor:pointer;
  border-left:3px solid transparent}
#tabs button:hover{background:#1a2e44;color:#fff}
#tabs button.active{background:#203954;color:#fff;border-left-color:var(--accent);font-weight:600}
#tabs button .dyn{font-size:9px;background:#315170;border-radius:3px;padding:1px 4px;margin-left:6px;
  color:#aecdec;vertical-align:1px}

/* ── main column ─────────────────────────────────────────── */
#main{flex:1;display:flex;flex-direction:column;height:100vh;min-width:0}
#bar{background:var(--panel);border-bottom:1px solid var(--line);padding:10px 16px;
  display:flex;gap:8px;align-items:center;flex-wrap:wrap}
#bar h2{font-size:16px;margin:0 12px 0 0}
#bar button, #fmt button{font:600 12.5px var(--sans);padding:7px 12px;border-radius:4px;
  border:1px solid var(--line);background:#f7f9fb;color:var(--ink);cursor:pointer}
#bar button:hover,#fmt button:hover{border-color:var(--accent);color:var(--accent)}
#bar button.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
#bar button.primary:hover{filter:brightness(1.08);color:#fff}
#bar label.chk{font-size:12.5px;display:flex;gap:5px;align-items:center;color:var(--muted)}
#status{margin-left:auto;font-size:12px;color:var(--muted);font-family:var(--mono)}

/* ── format builder panel ─────────────────────────────────── */
#fmt{background:#f4f7fa;border-bottom:1px solid var(--line);padding:12px 16px;display:none}
#fmt.open{display:block}
#fmt .cols{display:flex;gap:18px;flex-wrap:wrap}
#fmt fieldset{border:1px solid var(--line);border-radius:6px;background:#fff;
  min-width:230px;max-width:300px;padding:8px 10px;margin:0}
#fmt legend{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--muted)}
#fmt .list{max-height:180px;overflow-y:auto;font-size:12px}
#fmt .list label{display:block;padding:2px 2px;cursor:pointer}
#fmt .list label:hover{background:#eef4fb}
#fmt .counts label{display:flex;justify-content:space-between;align-items:center;
  font-size:12.5px;padding:3px 0}
#fmt .counts input{width:60px;padding:3px 6px;border:1px solid var(--line);border-radius:4px;
  font-family:var(--mono)}
#fmt .apply{margin-top:10px}
#fmt .note{font-size:11.5px;color:var(--muted);margin-top:8px;max-width:640px}

/* ── data grid ────────────────────────────────────────────── */
#gridWrap{flex:1;overflow:auto;padding:0}
table{border-collapse:separate;border-spacing:0;font-size:12.5px;min-width:100%}
th{position:sticky;top:0;background:#e7edf3;border-bottom:2px solid var(--line);
  border-right:1px solid var(--line);padding:6px 8px;text-align:left;font-size:11.5px;
  white-space:nowrap;z-index:2;vertical-align:bottom}
th .meta{display:block;font-weight:400;color:var(--muted);font-family:var(--mono);font-size:10px}
th.req{color:var(--req)}
th.rownum, td.rownum{position:sticky;left:0;background:#e7edf3;border-right:2px solid var(--line);
  min-width:38px;text-align:right;color:var(--muted);font-family:var(--mono);z-index:3;padding:4px 6px}
td{border-bottom:1px solid var(--line);border-right:1px solid var(--line);padding:0;background:#fff}
td input{width:100%;min-width:130px;border:0;padding:6px 8px;font:12.5px var(--mono);
  background:transparent;color:var(--ink)}
td input:focus{outline:2px solid var(--accent);outline-offset:-2px;background:#f4f9ff}
td.invalid{background:var(--err-bg)}
td.invalid input{color:var(--err)}
td .del{border:0;background:none;color:#b0bcc9;cursor:pointer;font-size:14px;padding:2px 6px}
td .del:hover{color:var(--err)}
#empty{padding:40px;text-align:center;color:var(--muted);font-size:13px}

/* ── footer / errors ─────────────────────────────────────── */
#errs{background:var(--err-bg);color:var(--err);border-top:1px solid #f2b8b5;
  max-height:130px;overflow-y:auto;padding:8px 16px;font-size:12px;display:none;font-family:var(--mono)}
#errs.open{display:block}
#errs div{padding:1px 0}
#errs div.warn{color:#8a5300}
#bar select{font:12.5px var(--sans);padding:5px 6px;border:1px solid var(--line);border-radius:4px;background:#fff}

/* ── dialogs ─────────────────────────────────────────────── */
dialog{border:1px solid var(--line);border-radius:8px;box-shadow:0 12px 40px rgba(10,25,45,.25);
  padding:0;max-width:560px;width:92vw}
dialog::backdrop{background:rgba(10,20,35,.45)}
dialog header{padding:12px 16px;border-bottom:1px solid var(--line);font-weight:700;font-size:14px}
dialog .body{padding:14px 16px;font-size:13px}
dialog footer{padding:10px 16px;border-top:1px solid var(--line);display:flex;gap:8px;justify-content:flex-end}
dialog input[type=text],dialog textarea{width:100%;padding:7px 9px;border:1px solid var(--line);
  border-radius:4px;font-family:var(--mono);font-size:12.5px}
dialog textarea{height:160px;resize:vertical}
dialog button{font:600 12.5px var(--sans);padding:7px 14px;border-radius:4px;
  border:1px solid var(--line);background:#f7f9fb;cursor:pointer}
dialog button.primary{background:var(--accent);border-color:var(--accent);color:#fff}
#helpBody{max-height:60vh;overflow-y:auto;white-space:pre-wrap;font-size:12.5px;line-height:1.5}
kbd{font-family:var(--mono);background:#eef1f4;border:1px solid var(--line);border-radius:3px;
  padding:0 4px;font-size:11px}
@media (prefers-reduced-motion: reduce){*{transition:none!important}}
</style>
</head>
<body>

<nav id="rail" aria-label="BAT transactions">
  <h1>BAT File Builder</h1>
  <div class="sub" id="srcInfo"></div>
  <input id="railSearch" type="search" placeholder="Filter tabs…" aria-label="Filter tabs">
  <div id="tabs" role="tablist"></div>
</nav>

<div id="main">
  <div id="bar">
    <h2 id="sheetTitle"></h2>
    <button id="btnFmt" hidden>Create File Format</button>
    <button id="btnAddRow">Add Row</button>
    <button id="btnAddRows">Add 10 Rows</button>
    <button id="btnPaste">Paste from Excel</button>
    <button id="btnClear">Clear Tab</button>
    <label class="chk" id="devTypeWrap" hidden>
      Device Type:
      <select id="devType">
        <option value="phone">Phone</option>
        <option value="cti">CTI Port</option>
        <option value="h323">H.323 Client</option>
        <option value="vgcv">VGC Virtual Phone</option>
        <option value="vgcp">VGC Phone</option>
        <option value="cipc">Cisco IP Communicator</option>
      </select>
    </label>
    <label class="chk" id="dummyWrap" hidden>
      <input type="checkbox" id="dummyMac"> Create Dummy MAC Address
    </label>
    <button class="primary" id="btnExport">Export to BAT Format</button>
    <button id="btnHelp" title="Help">Help</button>
    <span id="status"></span>
  </div>

  <div id="fmt">
    <div class="cols">
      <fieldset id="fsDevice"><legend>Device Fields</legend><div class="list" id="listDevice"></div></fieldset>
      <fieldset id="fsLine"><legend>Line Fields (per line)</legend><div class="list" id="listLine"></div></fieldset>
      <fieldset id="fsIntercom"><legend>Intercom Fields (per intercom)</legend><div class="list" id="listIntercom"></div></fieldset>
      <fieldset><legend>Counts</legend><div class="counts" id="countInputs"></div>
        <button class="primary apply" id="btnApplyFmt">Apply Format</button>
      </fieldset>
    </div>
    <div class="note">Changing the format rebuilds this tab's columns. Data already entered is kept
      for columns whose names still exist. This mirrors the xlt's <b>Create File Format</b> dialog.</div>
  </div>

  <div id="gridWrap"><div id="empty">Click <b>Add Row</b> or <b>Paste from Excel</b> to begin.</div></div>
  <div id="errs"></div>
</div>

<dialog id="dlgExport">
  <header>Export to BAT Format</header>
  <div class="body">
    <p style="margin-top:0">File name (a <span style="font-family:var(--mono)">.txt</span> extension is added automatically):</p>
    <input type="text" id="exportName" spellcheck="false">
    <p id="exportSummary" style="color:var(--muted);font-size:12px"></p>
  </div>
  <footer>
    <button id="exportCancel">Cancel</button>
    <button class="primary" id="exportGo">Export</button>
  </footer>
</dialog>

<dialog id="dlgPaste">
  <header>Paste from Excel / CSV</header>
  <div class="body">
    <p style="margin-top:0">Copy rows in Excel and paste them here (<kbd>Ctrl</kbd>+<kbd>V</kbd>).
    Tab-separated (Excel) and comma-separated values are both accepted. Columns map left-to-right
    onto this tab's columns. Do <b>not</b> include the header row.</p>
    <textarea id="pasteArea" spellcheck="false" placeholder="SEP0011223344AA&#9;Bldg 2050 Rm 114&#9;…"></textarea>
  </div>
  <footer>
    <button id="pasteCancel">Cancel</button>
    <button class="primary" id="pasteGo">Add Rows</button>
  </footer>
</dialog>

<dialog id="dlgHelp">
  <header>Help</header>
  <div class="body"><div id="helpBody"></div></div>
  <footer><button class="primary" id="helpClose">Close</button></footer>
</dialog>

<script>
"use strict";
const MODEL = __MODEL__;

/* ── state ──────────────────────────────────────────────────────────── */
const state = {};           // sheetName -> {columns:[], rows:[[]], fmt:{...}, dummy:false}
let current = null;         // active sheet name

function sheetDef(name){ return MODEL.sheets.find(s => s.name === name); }
function isDynamic(name){ return Object.prototype.hasOwnProperty.call(MODEL.dynamic, name); }
function txnKey(name){ return MODEL.dynamic[name]; }

function dbFields(txn, cat){
  return MODEL.database.filter(f => f.cat === cat && f.ops.includes(txn));
}

function initSheet(name){
  if (state[name]) return state[name];
  const def = sheetDef(name);
  const st = {
    columns: def.columns.map(c => ({...c, base:true})),
    rows: [],
    dummy: false,
    devType: "phone",
    fmt: null,
  };
  if (isDynamic(name)){
    st.fmt = {device:[], line:[], intercom:[], counts:{lines:0,intercoms:0,speedDial:0,blfSpeedDial:0,blfDcp:0,remoteDest:0}};
  }
  state[name] = st;
  return st;
}

/* ── dynamic format ─────────────────────────────────────────────────── */
function rebuildDynamicColumns(name){
  const st = state[name];
  const def = sheetDef(name);
  const cols = def.columns.map(c => ({...c, base:true}));
  const f = st.fmt;
  // device fields
  for (const fd of f.device) cols.push({name: fd.name, type: fd.type, len: fd.len, req: fd.req});
  // per-line groups: "Field Name N" (Directory Number leads each set)
  const dnFirst = arr => {
    const i = arr.findIndex(x => /^(Intercom )?Directory Number$/i.test(x.name));
    return i > 0 ? [arr[i], ...arr.slice(0,i), ...arr.slice(i+1)] : arr;
  };
  for (let n = 1; n <= f.counts.lines; n++)
    for (const fd of dnFirst(f.line))
      cols.push({name: fd.name + " " + n, type: fd.type, len: fd.len, req: fd.req, grp: "L"+n});
  for (let n = 1; n <= f.counts.intercoms; n++)
    for (const fd of dnFirst(f.intercom))
      cols.push({name: fd.name + " " + n, type: fd.type, len: fd.len, req: fd.req, grp: "I"+n});
  // hardcoded repeat groups
  for (const [key, grp] of Object.entries(MODEL.repeatGroups)){
    if (!grp.sheets.includes(name)) continue;
    const cnt = f.counts[key] || 0;
    for (let n = 1; n <= cnt; n++)
      for (const fd of grp.fields)
        cols.push({name: fd.name + " " + n, type: fd.type, len: fd.len, req: fd.req});
  }
  // preserve entered data by column name
  const oldCols = st.columns, oldRows = st.rows;
  const map = cols.map(c => oldCols.findIndex(o => o.name === c.name));
  st.rows = oldRows.map(r => map.map(i => i >= 0 ? (r[i] ?? "") : ""));
  st.columns = cols;
}

/* ── grid rendering ─────────────────────────────────────────────────── */
const gridWrap = document.getElementById("gridWrap");

function renderGrid(){
  const st = state[current];
  if (!st.rows.length){
    gridWrap.innerHTML = '<div id="empty">Click <b>Add Row</b> or <b>Paste from Excel</b> to begin.</div>';
    updateStatus(); return;
  }
  const t = document.createElement("table");
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  hr.appendChild(Object.assign(document.createElement("th"), {className:"rownum", textContent:"#"}));
  for (const c of st.columns){
    const th = document.createElement("th");
    if (c.req) th.className = "req";
    th.textContent = c.name + (c.req ? " *" : "");
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = c.type + (c.len ? "[" + c.len + "]" : "") + (c.req ? " MANDATORY" : " OPTIONAL");
    th.appendChild(meta);
    hr.appendChild(th);
  }
  hr.appendChild(Object.assign(document.createElement("th"), {textContent:""}));
  thead.appendChild(hr); t.appendChild(thead);

  const tb = document.createElement("tbody");
  st.rows.forEach((row, ri) => {
    const tr = document.createElement("tr");
    const rn = document.createElement("td"); rn.className = "rownum"; rn.textContent = ri + 1;
    tr.appendChild(rn);
    st.columns.forEach((c, ci) => {
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.value = row[ci] ?? "";
      inp.dataset.r = ri; inp.dataset.c = ci;
      inp.setAttribute("aria-label", c.name + " row " + (ri+1));
      inp.addEventListener("input", onCell);
      inp.addEventListener("paste", onCellPaste);
      td.appendChild(inp);
      validateCell(td, inp.value, c);
      tr.appendChild(td);
    });
    const tdDel = document.createElement("td");
    const del = document.createElement("button");
    del.className = "del"; del.textContent = "✕"; del.title = "Delete row";
    del.addEventListener("click", () => { st.rows.splice(ri,1); renderGrid(); });
    tdDel.appendChild(del); tr.appendChild(tdDel);
    tb.appendChild(tr);
  });
  t.appendChild(tb);
  gridWrap.innerHTML = ""; gridWrap.appendChild(t);
  updateStatus();
}

function onCell(e){
  const st = state[current];
  const r = +e.target.dataset.r, c = +e.target.dataset.c;
  st.rows[r][c] = e.target.value;
  validateCell(e.target.parentElement, e.target.value, st.columns[c]);
  updateStatus();
}

/* Excel-style multi-cell paste directly into the grid */
function onCellPaste(e){
  const text = (e.clipboardData || window.clipboardData).getData("text");
  if (!text || (!text.includes("\t") && !text.includes("\n"))) return; // single value: default paste
  e.preventDefault();
  const st = state[current];
  const startR = +e.target.dataset.r, startC = +e.target.dataset.c;
  const rows = text.replace(/\r/g,"").split("\n").filter(l => l.length);
  rows.forEach((line, dr) => {
    const cells = line.split("\t");
    const rr = startR + dr;
    while (st.rows.length <= rr) st.rows.push(new Array(st.columns.length).fill(""));
    cells.forEach((v, dc) => {
      const cc = startC + dc;
      if (cc < st.columns.length) st.rows[rr][cc] = v.trim();
    });
  });
  renderGrid();
}

/* ══ VBA-ported validation engine ═════════════════════════════════════
   Masks and messages are ported 1:1 from the bat.xlt macros (Module1).
   allow() = fnValidateString (every char must be in mask);
   deny()  = fnValidateDisplay-style (no char may be in mask).           */
const ALNUM = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
const M = {
  dn:        "1234567890*#Xx[]^\\+?!-",
  dnTel:     "1234567890*#Xx[]^+?!-",
  dnWide:    "1234567890*#Xx[]^\\+?!-_" + "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ",
  devName:   ALNUM + "-_.",
  devNameSp: ALNUM + "-_ .",
  anmSpDot:  ALNUM + " -_.",
  partition: ALNUM + "-_ ",
  digits:    "1234567890",
  fwd:       ALNUM + "!#$%^&*()_+~`-={}|\\:?[];',./@",
  sdName:    ALNUM + "!#$%^&*()_+~`-={}|\\:?[];',./@\" ",
  sdNumber:  "01234567890*,#",
  blfsd:     "01234567890*#" + "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ",
  extPh:     "1234567890X*#+",
  e164:      ALNUM + " ",
  cpn:       "1234567890abcdABCD*#\\+",
  h323:      ALNUM + "-.",
  cipc:      "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ",
  hex:       "1234567890ABCDEF",
  loc:       "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_ ",
  udp:       "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-.,",
  mgcp:      ALNUM + "-_.",
  portDesc:  "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-_/@:{}[];,|+=()!#$^*? ",
  ipv4:      "1234567890.",
  ipv6:      "1234567890abcdefABCDEF.:",
  bssid:     "1234567890abcdefABCDEF:",
  parkFwd:   "1234567890*#Xx",
  // deny masks
  xLbl:      "&[]<>%\"",
  xDisplay:  "<>{}|[]\\%&\"",
  xAlert:    "<>{}|",
  xMgr:      " =+<>#;\\,\"",
  xSDLabel:  "~|\"",
  xBLFLabel: "~#|\"",
  xGrpName:  "&\\<>%\"",
  xCMC:      "[]<>%\"",
  xFAC:      "&[]<>%\"",
  xAssocPC:  "<>{}|[]\\%&\"",
  xMGCPDesc: "&<>%[]\"",
};
const FWDMSG = "can contain alphabets,numbers,spaces and characters '\" ','!','#','$','%','^','&','*','(',')','_','+','~','`','-','=','{','}','|','\\',':','?'[',']',';',''',',','.','/','@' only.";
const GRPMSG = "may not contain ampersand(&), double quotes(\"), less than(<), greater than(>),backslash (\\) nor the percent sign(%)";

function allow(v, mask){ for (const ch of v) if (!mask.includes(ch)) return false; return true; }
function deny(v, mask){ for (const ch of v) if (mask.includes(ch)) return true; return false; }
function lenMsg(name, n){ return "Length of " + name + " must be less than or equal to " + n + "."; }
function blankMsg(name){ return name + " cannot be blank."; }
// fnvalidateLenComma: fields exempt from the comma error in the macro
const COMMA_EXEMPT = /^(DESCRIPTION|SPEED DIAL NUMBER|MGCP DESCRIPTION|DISPLAY|LINE TEXT LABEL|SPEED DIAL LABEL|FIRSTNAME|LASTNAME|DEPARTMENT|PASSWORD|ROUTE PARTITION|USER DEVICE PROFILE)$/;
function lenComma(v, max, name){
  if (v.length > max) return lenMsg(name, max);
  if (!COMMA_EXEMPT.test(name.toUpperCase()) && !name.startsWith("Middle") && v.includes(","))
    return name + " cannot have comma.";
  return null;
}
// fnValidateDN incl. the '^ must be inside []' rule
function vDN(v, header){
  const h = (header||"").toUpperCase();
  let mask = M.dn, msg = "Directory Number may have numbers and characters '*','#','X','[',']','^','\\+','?','!','-' only.";
  if (h.includes("TELEPHONE")){ mask = M.dnTel; msg = "Telephone Number may have numbers and characters '*','#','X','[',']','^','+','?','!','-' only."; }
  else if (h.includes("IPCC") || h.includes("PRIMARY")) mask = M.dnWide;
  if (!allow(v, mask)) return msg;
  const p = v.indexOf("^");
  if (v !== "" && p >= 0){
    const caretMsg = "The '^' character should be enclosed between [] only.";
    if (p === 0 || p === v.length - 1) return caretMsg;
    if (v[p+1] !== "]" || v[p-1] !== "[") return caretMsg;
  }
  return null;
}
function tf1(v, lenLbl, valLbl){
  if (v.length > 1) return "Length of " + lenLbl + " must be equal to 1";
  if (v.length > 0 && !allow(v, "tf")) return valLbl;
  return null;
}
function digitsRange(v, digitMsg, lo, hi, rangeMsg){
  if (!allow(v, M.digits)) return digitMsg;
  const n = parseInt(v, 10);
  if (v !== "" && (n < lo || n > hi)) return rangeMsg;
  return null;
}
// fnCheckMacAddress: behavior depends on sheet + device-type radio + dummy MAC
function vMacDevice(v, ctx){
  const S = ctx.sheet.toUpperCase(), dt = ctx.devType || "phone";
  if (ctx.dummy) return null;                          // dummy MAC: skipped (confirm at export)
  const ctiMode = S === "CTI PORT" || S === "CTI PORT-USERS" || S === "USERS" ||
                  S === "UPDATE USERS" || S === "CUSTOM MANAGERS-ASSISTANTS" || dt === "cti";
  const h323Mode = S === "H.323 CLIENT" || S === "H.323 CLIENT-USERS" || S === "ADD LINES" ||
                   S === "ADD INTERCOM" || dt === "h323";
  if (dt === "vgcv" || dt === "vgcp"){                 // fnCheckVGCName
    if (v !== ""){
      const last2 = v.slice(-2);
      if (dt === "vgcp"){
        if (!allow(last2, M.digits) || +last2 < 1 || +last2 > 48)
          return "Last two digits of MAC Address should be between 01 and 48 only.";
      } else if (last2 !== "00") return "Last two digits of MAC Address should be 00.";
    }
    return null;
  }
  if (dt === "cipc"){                                  // fnCheckCIPCName
    if (v === "") return null;
    if (v.length < 1 || v.length > 15) return "Device Name for Cisco IP Communicator can contain only 1-15 hexadecimal characters";
    if (!allow(v, M.cipc)) return "Cisco IP Commuincator Phone name may contain only characters (A-Z 0-9). Alphabets need to be in upper case.";
    return null;
  }
  if (h323Mode && (S === "H.323 CLIENT" || S === "H.323 CLIENT-USERS" || dt === "h323")){ // fnCheckH323Name
    if (v === "") return blankMsg("Device Name");
    if (v.length > 50) return lenMsg("Device Name", 50);
    if (!allow(v, M.h323)) return "H.323 device name may contain only characters (A-Z a-z 0-9),dots and dashes only.";
    if (!allow(v[0], "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"))
      return "H.323 device name may contain only letters as the first character.";
    return null;
  }
  if (v === ""){
    if (ctx.mandatory === false) return null;
    if (ctiMode || h323Mode) return blankMsg("Device Name");
    if (S === "REMOTE DESTINATION PROFILE") return blankMsg("Desk Phone Name");
    return blankMsg("MAC Address");
  }
  if (ctiMode){
    if (v.length >= 16)
      return lenMsg(S === "USERS" || S === "UPDATE USERS" ? "MAC Address/Device Name" : "Device Name", 15);
    if (!allow(v, M.devName)) return "Device Name may contain alphabets, numbers and characters('-', '_' and '.') only.";
  } else if (h323Mode){
    if (v.length > 50) return lenMsg("Device Name", 50);
    if (!allow(v, M.devName)) return "Device Name may contain alphabets, numbers and characters('-', '_' and '.') only.";
  } else if (S === "USER DEVICE PROFILES"){
    if (v.length > 50) return lenMsg("Device Profile Name", 50);
    if (!allow(v, M.devNameSp)) return "Device Profile Name may contain alphabets, numbers, spaces and characters('-', '_' and '.') only.";
  } else if (S === "REMOTE DESTINATION PROFILE"){
    if (v.length > 50) return lenMsg("Device Name", 50);
    if (!allow(v, M.devName)) return "Desk Phone Name may contain alphabets, numbers and characters('-', '_' and '.') only.";
  } else {
    if (v.length > 50) return lenMsg("Device Name", 50);
    if (!allow(v, M.devNameSp)) return "Device Name may contain alphabets, numbers and characters('-', '_' and '.') only.";
  }
  return null;
}
/* Ordered binding table — most specific patterns first, mirroring the
   macro's If/ElseIf chains. fn receives (value, ctx) and returns the
   exact VBA error message or null.  ctx: {sheet, devType, dummy, header} */
const BINDINGS = [
 {p:/MOBILITY IDENTITY ANSWER TOO SOON TIMER/, f:v=>allow(v,M.digits)?null:"Mobility Identity Answer Too Soon Timer may have numbers only."},
 {p:/MOBILITY IDENTITY ANSWER TOO LATE TIMER/, f:v=>allow(v,M.digits)?null:"Mobility Identity Answer Too Late Timer may have numbers only."},
 {p:/MOBILITY IDENTITY DELAY BEFORE RINGING TIMER/, f:v=>allow(v,M.digits)?null:"Mobility Identity Delay Before Ringing Timer may have numbers only."},
 {p:/ANSWER TOO SOON TIMER/, f:v=>allow(v,M.digits)?null:"Answer Too Soon Timer may have numbers only."},
 {p:/ANSWER TOO LATE TIMER/, f:v=>allow(v,M.digits)?null:"Answer Too Late Timer may have numbers only."},
 {p:/DELAY BEFORE RINGING TIMER/, f:v=>allow(v,M.digits)?null:"Delay Before Ringing Timer may have numbers only."},
 {p:/REMOTE DESTINATION PROFILE NAME/, f:v=>{
    if (v==="") return "Remote Destination Profile Name cannot be blank.";
    return allow(v,M.anmSpDot)?null:"Remote Destination Profile Name can contain only letters,numbers,spaces,dashes,dots and underscore as valid characters.";}},
 {p:/REMOTE DESTINATION NAME/, f:v=>allow(v,M.anmSpDot)?null:"Remote Destination Name can contain only letters,numbers,spaces,dashes,dots and underscore as valid characters."},
 {p:/REMOTE DESTINATION NUMBER/, f:v=>allow(v,M.digits)?null:"Remote Destination Number may have numbers only."},
 {p:/REMOTE DESTINATION LIMIT/, f:v=>{
    if (v==="") return null;
    if (v.length>2 || !allow(v,M.digits) || +v<1 || +v>10) return "Range of Remote Destination Limit is 1-10";
    return null;}},
 {p:/MAXIMUM WAIT TIME FOR DESK PICKUP/, f:v=>{
    if (v==="") return null;
    if (v.length>5 || !allow(v,M.digits) || +v<0 || +v>30000) return "Range of Maximum Wait Time for Desk Pickup is 0-30000.";
    return null;}},
 {p:/INTERCOM CALLER NAME/, f:v=>allow(v,"tTfF.")?null:"Intercom Caller Name can take only t or f."},
 {p:/INTERCOM CALLER NUMBER/, f:v=>allow(v,"tTfF")?null:"Intercom Caller Number can take either t or f."},
 {p:/INTERCOM DIRECTORY NUMBER/, f:(v,c)=>vDN(v,c.header)},
 {p:/DEVICE PROFILE NAME/, f:(v,c)=>{
    if (v==="") return c.mandatory===false?null:"User Device Profile cannot be blank.";
    if (v.length>50) return lenMsg("Device Profile Name",50);
    return allow(v,M.udp)?null:"User Device Profile can contain alphabets, numbers, spaces and characters -, _ and . only.";}},
 {p:/DEFAULT PROFILE|USER DEVICE PROFILE/, f:v=>allow(v,M.udp)?null:"User Device Profile can contain alphabets, numbers, spaces and characters -, _ and . only."},
 {p:/MAC ADDRESS|OLD DEVICE NAME|^DEVICE NAME/, f:(v,c)=>vMacDevice(v,c)},
 {p:/DESK PHONE NAME/, f:(v,c)=>vMacDevice(v,Object.assign({},c,{sheet:"Remote Destination Profile"}))},
 {p:/BUSY LAMP FIELD LABEL ASCII|BLF DIRECTED CALL PARK LABEL ASCII/, f:v=>deny(v,M.xBLFLabel)?"Busy Lamp Field Label Ascii may not contain ~,#,| and double quotes.":null},
 {p:/BUSY LAMP FIELD LABEL|BLF DIRECTED CALL PARK LABEL/, f:v=>deny(v,M.xBLFLabel)?"Busy Lamp Field Label may not contain ~,#,| and double quotes.":null},
 {p:/BUSY LAMP FIELD DIRECTORY NUMBER|BLF DIRECTED CALL PARK DIRECTORY NUMBER/, f:v=>allow(v,M.blfsd)?null:"BLF Directory number may contain numbers and the characters(*,#)only."},
 {p:/BUSY LAMP FIELD DESTINATION/, f:()=>null},
 {p:/SPEED DIAL LABEL/, f:v=>{
    const e=lenComma(v,30,"Speed Dial Label"); if(e) return e;
    return deny(v,M.xSDLabel)?"Speed Dial Label may not contain ~,#,| and double quotes.":null;}},
 {p:/SPEED DIAL NUMBER|INTERCOM SPEED DIAL|^SPEED DIAL/, f:v=>
    (allow(v,M.sdNumber)||allow(v,M.sdName))?null:"Speed Dial "+FWDMSG},
 {p:/NO ANSWER RING DURATION.*MLPP|MLPP.*NO ANSWER RING DURATION/, f:v=>digitsRange(v,"No Answer Ring Duration  cannot contain characters other than digits.",4,60,"The No Answer Ring Duration (MLPP) value should be between 4 and 60.")},
 {p:/NO ANSWER RING DURATION/, f:v=>digitsRange(v,"No Answer Ring Duration  cannot contain characters other than digits.",1,300,"The No Answer Ring Duration (Cfwd) value should be between 1 and 300.")},
 {p:/BUSY TRIGGER/, f:v=>digitsRange(v,"Busy Trigger cannot contain characters other than digits.",1,200,"The Busy Trigger value should be between 1 and 200.")},
 {p:/MAXIMUM NUMBER OF CALLS/, f:v=>digitsRange(v,"'Maximum Number of Calls' cannot contain characters other than digits.",1,200,"The 'Maximum Number of Calls' value should be between 1 and 200.")},
 {p:/DND TIMEOUT/, f:v=>digitsRange(v,"DND Timeout may have numbers only.",0,120,"The Limit of DND Timeout is 0 - 120")},
 {p:/^DND$|DO NOT DISTURB/, f:v=>allow(v,"tTfF.")?null:"Do Not Disturb can take only t or f."},
 {p:/PARK MONITOR FORWARD NO RETRIEVE (INT|EXT) VOICE MAIL/, f:v=>allow(v,"ftFT")?null:"Park Monitor Forward No Retrieve Internal/External Voice Mail can contain only either T/F or t/f."},
 {p:/PARK MONITORING REVERSION TIMER/, f:v=>{
    if (v==="") return null;
    if (!allow(v,M.digits) || +v<0 || +v>1200) return "Park Monitoring Reversion Timer Value should lie between 0 to 1200.";
    return null;}},
 {p:/PARK MONITOR FORWARD NO RETRIEVE (INT|EXT) DESTINATION/, f:v=>allow(v,M.parkFwd)?null:"Park Monitor Forward may contain numbers and characters (*#X) only."},
 {p:/HOTLINE DEVICE/, f:v=>allow(v,"ftFT")?null:"Hotline Device can contain only either T/F or t/f."},
 {p:/IS MOBILE PHONE/, f:v=>allow(v,"ftFT")?null:"Is Mobile Phone can contain only either T/F or t/f."},
 {p:/REJECT ANONYMOUS CALLS/, f:v=>tf1(v,"Reject Anonymous Calls","Reject Anonymous Calls can contain t or f only.")},
 {p:/URGENT PRIORITY|E164 IS URGENT/, f:v=>tf1(v,"Urgent Priority","UrgentPriority can contain t or f only.")},
 {p:/URI \d+ IS PRIMARY/, f:v=>tf1(v,"URI Is Primary On Directory Number","Allow URI Is Primary On Directory Number can contain t or f only.")},
 {p:/URI \d+ ROUTE PARTITION/, f:v=>allow(v,M.partition)?null:"Route Partition can contain alphabets,numbers,spaces and characters '_','-' only."},
 {p:/URI \d+ ON DIRECTORY NUMBER/, f:v=>allow(v,M.fwd)?null:"URI "+FWDMSG},
 {p:/DIRECTORY URI/, f:v=>allow(v,M.fwd)?null:"URI "+FWDMSG},
 {p:/AAR CSS|AAR CALLING SEARCH SPACE/, f:v=>allow(v,M.anmSpDot)?null:"AAR Calling Search Space cannot contain characters other than alphabets,numbers,spaces and charaters '_','-' and '.' ."},
 {p:/AAR GROUP/, f:v=>allow(v,M.anmSpDot)?null:"AAR Group cannot contain characters other than alphabets,numbers,spaces and charaters '_','-' and '.' ."},
 {p:/AAR DESTINATION MASK|EXTERNAL PHONE NUMBER MASK/, f:v=>allow(v,M.extPh)?null:"External Phone Number Mask cannot contain characters other than numbers,X and charaters '*','#' and '+'."},
 {p:/SUBSCRIBE CALLING SEARCH SPACE|CALLING SEARCH SPACE|LINE CSS|INTERCOM CSS|MONITORING CSS|^CSS| CSS/, f:v=>allow(v,M.anmSpDot)?null:"Calling Search Space cannot contain characters other than alphabets,numbers,spaces and characters '_','-' and '.' ."},
 {p:/MEDIA RESOURCE GROUP LIST/, f:v=>allow(v,M.anmSpDot)?null:"Media Resource Group List cannot contain characters other than alphabets,numbers,spaces and characters '_','-' and '.' ."},
 {p:/USER HOLD MOH AUDIO SOURCE/, f:v=>allow(v,M.digits)?null:"User Hold MOH Audio Source can take only integer values."},
 {p:/NETWORK HOLD MOH AUDIO SOURCE/, f:v=>allow(v,M.digits)?null:"Network hold MOH audio source can take only integer values."},
 {p:/SOFTKEY TEMPLATE/, f:v=>allow(v,M.anmSpDot)?null:"SoftKey Template cannot contain characters other than alphabets,numbers,spaces and charaters '_','-' and '.' ."},
 {p:/VOICE MAIL PROFILE/, f:v=>allow(v,M.anmSpDot)?null:"Voice Mail Profile cannot contain characters other than alphabets,numbers,spaces and charaters '_','-' and '.' ."},
 {p:/DEVICE POOL/, f:v=>allow(v,M.anmSpDot)?null:"Device Pool cannot contain characters other than digits, alphabets and -,., _"},
 {p:/PHONE LOAD NAME/, f:v=>allow(v,M.devName)?null:"Phone Load Name cannot contain characters other than digits, alphabets and -,., _"},
 {p:/E164|E\.164/, f:v=>allow(v,M.e164)?null:"E164 field cannot contain characters other than digits."},
 {p:/^LOCATION/, f:v=>allow(v,M.loc)?null:"Location may contain alphabets, numbers, spaces and characters ('-' , '_' , '.') only."},
 {p:/LINE TEXT LABEL/, f:v=>deny(v,M.xLbl)?"Line Text Label cannot contain &,[,],<,>,% and double quotes.":null},
 {p:/ASCII ALERTING NAME|ALERTING NAME/, f:v=>deny(v,M.xAlert)?"Alerting Name cannot contain characters <,>,{,},|.":null},
 {p:/ASCII DISPLAY|^DISPLAY|INTERCOM DISPLAY/, f:v=>deny(v,M.xDisplay)?"Display cannot contain characters <,>,{,},|,[,],\\,%,& and double quotes.":null},
 {p:/FORWARD ALL/, f:v=>allow(v,M.fwd)?null:"Forward All Destination "+FWDMSG},
 {p:/FORWARD BUSY/, f:v=>allow(v,M.fwd)?null:"Forward Busy Destination "+FWDMSG},
 {p:/FORWARD NO ANSWER/, f:v=>allow(v,M.fwd)?null:"Forward No Answer "+FWDMSG},
 {p:/FORWARD NO COVERAGE/, f:v=>allow(v,M.fwd)?null:"Forward No Coverage "+FWDMSG},
 {p:/PICKUP GROUP NAME/, f:(v,c)=>{
    if (c.sheet.toUpperCase()==="CALL PICKUP GROUP" && v==="") return "Pickup Group Name cannot be blank.";
    return allow(v,M.anmSpDot)?null:"Pickup Group Name can contain only letters,numbers,spaces,dashes,dots and underscore as valid characters.";}},
 {p:/PICKUP GROUP NUMBER/, f:v=>{
    if (v==="") return "Pickup Group Number cannot be blank.";
    if (v.length>24 || !allow(v,M.cpn)) return "Pickup Group Number can accept up to 24 digits and allow the following characters: numeric (0-9), letter A-D, plus (+), pound (#) and asterisk (*). Note that slash (\\) must be used in front of plus (+).";
    return null;}},
 {p:/CALL PICKUP GROUP/, f:v=>allow(v,M.anmSpDot)?null:"Pickup Group Name can contain only letters,numbers,spaces,dashes,dots and underscore as valid characters."},
 {p:/LINE INDEX/, f:v=>{
    if (v==="") return null;                       // blank handled by Add Lines cross-field
    if (!allow(v,M.digits)) return "Line Index may have numbers only.";
    if (+v<1 || +v>34) return "Line Index may lie between 1 and 34 only.";
    return null;}},
 {p:/NUMBER OF LINES/, f:v=>{
    if (v==="") return null;                       // mandatory handled per-sheet
    if (!allow(v,M.digits)) return "Number of Lines field cannot contain characters other than numbers.";
    if (+v<0 || +v>34) return "The value for Number of Lines should be between 0 and 34.";
    return null;}},
 {p:/AUTHORIZATIONLEVEL|AUTHORIZATION LEVEL/, f:v=>{
    if (v==="") return null;
    if (!allow(v,M.digits)) return "Authorization level may have numbers only.";
    if (+v<0 || +v>255) return "Authorization level may lie between 0 and 255 only.";
    return null;}},
 {p:/AUTHORIZATION CODE NAME/, f:v=>deny(v,M.xFAC)?"Authorization Code Name may not contain ampersand(&), double quotes(\"), brackets([]), less than(<), greater than(>), nor the percent sign(%)":null},
 {p:/^MANAGER$|MANAGER USER ID|^MANAGER ID|ASSISTANT ID/, f:(v,c)=>{
    const e=lenComma(v,30,"Manager"); if(e && /^MANAGER$/.test(c.header.toUpperCase())) return e;
    return deny(v,M.xMgr)?"Manager field cannot contain characters =, +, <, >, #, ;, \\, \", comma and space.":null;}},
 {p:/USERID|USER ID|OWNER ID|END USER ID/, f:(v,c)=>{
    const S=c.sheet.toUpperCase();
    if (v==="" && (S==="PHONES-USERS"||S==="TIME OF DAY ACCESS"||S==="USERS"||S==="UPDATE USERS"))
      return "User ID cannot be blank.";
    return null;}},
 {p:/^PIN$/, f:v=>allow(v,M.digits)?null:"PIN can contain numbers only."},
 {p:/ALLOW CONTROL OF DEVICE FROM CTI/, f:v=>tf1(v,"Allow Control of Device from CTI","Allow Control of Device from CTI can contain t or f only.")},
 {p:/ENABLE MOBILE VOICE ACCESS/, f:v=>tf1(v,"Enable Mobile Voice Access","Allow Enable Mobile Voice Access can contain t or f only.")},
 {p:/ENABLE MOBILITY/, f:v=>tf1(v,"Enable Mobility","Allow Enable Mobility can contain t or f only.")},
 {p:/ASSOCIATED PC/, f:v=>deny(v,M.xAssocPC)?"Associated PC cannot contain characters <,>,{,},|,[,],\\,%,& and double quotes.":null},
 {p:/TELEPHONE NUMBER|MOBILE NUMBER|HOME NUMBER|PAGER NUMBER|PRIMARY EXTENSION|IPCC EXTENSION/, f:(v,c)=>vDN(v,c.header)},
 {p:/FIRST NAME|MIDDLE NAME|LAST NAME/, f:(v,c)=>{
    const which=/FIRST/.test(c.header.toUpperCase())?"First":(/MIDDLE/.test(c.header.toUpperCase())?"Middle":"Last");
    return lenComma(v,64,which+"Name");}},
 {p:/MGCP DOMAIN NAME|^DOMAIN NAME/, f:v=>allow(v,M.mgcp)?null:"MGCP Domain Name may have only alphabets, numbers and characters ('-', '.' and '_')"},
 {p:/PORT IDENTIFIER/, f:v=>{
    if (v==="") return null;
    if (v.includes(",")) return "Comma is not Allowed in Port Number.";
    if (v.length!==3) return "Length of Port Number should be equal to 3.";
    if (!allow(v,M.digits)) return "Port Number may have only numbers.";
    const n=+v;
    if (!((n>=1&&n<=24)||(n>=101&&n<=124))) return "Port Number value must lie between 001 and 024 OR 101 and 124.";
    return null;}},
 {p:/PORT( \d+)? DESCRIPTION/, f:v=>allow(v,M.portDesc)?null:" Port Description may contain alphabets, numbers, spaces and the characters ('-' , '_', '/', '@',':','{', '}', '[', ']', ';', ',', '|', '+', '=', '(', ')', '!', '#', '$', '^', '*', '?' and '.') only."},
 {p:/PORT NUMBER/, f:v=>{
    if (v==="") return null;
    if (!allow(v,M.digits)) return "The Port Number should contains numbers only.";
    if (+v<1 || +v>24) return "Port Number value may lie between 1 and 24 only.";
    return null;}},
 {p:/TIME OF DAY ACCESS NAME/, f:v=>deny(v,M.xGrpName)?"Time of Day Access Name "+GRPMSG:null},
 {p:/ROUTE PARTITION|INTERCOM ROUTE PARTITION|PORT( \d+)? PARTITION|^PARTITION/, f:v=>allow(v,M.partition)?null:"Route Partition can contain alphabets,numbers,spaces and characters '_','-' only."},
 {p:/IPV4 ADDRESS/, f:v=>{
    if (v==="") return null;
    if (v.length<7||v.length>15||!allow(v,M.ipv4)) return "IPv4 address can contain from 7 to 15 characters. It must be in dotted decimal format (digits and dots only).";
    return null;}},
 {p:/IPV6 ADDRESS/, f:v=>{
    if (v==="") return null;
    if (v.length<1||v.length>50||!allow(v,M.ipv6)) return "IPv6 address can contain from 1 to 50 characters. It can be formatted as needed but may only contain Hexadecimal digits (0-9, A-F), colons and dots";
    return null;}},
 {p:/^BSSID/, f:v=>{
    if (v==="") return null;
    if (v.length<1||v.length>20||!allow(v,M.bssid)) return "BSSIDwithMask can contain from 1 to 20 characters. It can be formatted as needed but may only contain Hexadecimal digits (0-9, A-F), colons.";
    return null;}},
 {p:/DIRECTORY NUMBER/, f:(v,c)=>{
    if (v.length>50) return lenMsg("Directory Number",50);
    return vDN(v,c.header);}},
];
// sheet -> which group-name mask/message applies (per fnValidate*Name family)
const NAME_SHEET_RULES = {
  "Trust Group":       {blank:"Name Field is Mandatory.", msg:"Trust Group Name "+GRPMSG},
  "Trust Element":     {blank:"Name Field is Mandatory.", msg:"Trust Element Name "+GRPMSG},
  "Enrolled Group":    {blank:"Group Name Field is Mandatory.", msg:"Enrolled Group Name "+GRPMSG},
  "Exclusion Group":   {blank:"Name Field is Mandatory.", msg:"Exclusion Group Name "+GRPMSG},
  "FallBack Profile":  {blank:"Name Field is Mandatory.", msg:"FallBack Profile Name "+GRPMSG},
};
function validateCellVBA(v, col, ctx){
  v = (v||"").trim();
  const H = col.name.toUpperCase();
  // sheet-specific "Name" columns
  const nr = NAME_SHEET_RULES[ctx.sheet];
  if (nr && /^NAME$/.test(H)){
    if (v==="") return nr.blank;
    return deny(v, M.xGrpName) ? nr.msg : null;
  }
  if (ctx.sheet==="Insert CMC" && H==="DESCRIPTION")
    return deny(v, M.xCMC) ? "Description may not contain double quotes(\"), brackets([]), less than(<), greater than(>), nor the percent sign(%)" : null;
  if (/VG200|VG224|VG202|VG310|VG320|VG350|VG450|VG420|ISR4461|NM-HDA/.test(ctx.sheet.toUpperCase()) && H==="DESCRIPTION")
    return deny(v, M.xMGCPDesc) ? "Description cannot contain characters &,<, >, %,[,] and \"." : null;
  if (ctx.sheet==="Insert Infrastructure Device" && /WAPLOCATION|^DESCRIPTION/.test(H)){
    if (v.length>63 || deny(v,"\\\"")) return "WAPLocation can contain up to 63 characters. All characters except double quotes, backslash and non-printable characters";
    return null;
  }
  for (const b of BINDINGS){
    if (b.p.test(H)){
      const msg = b.f(v, {sheet:ctx.sheet, devType:ctx.devType, dummy:ctx.dummy, header:col.name, mandatory:col.req});
      if (msg) return msg;
      // binding handled this column; still enforce header length if binding didn't
      if (col.len && v.length > col.len && !/LENGTH OF/i.test(msg||"")) return lenMsg(col.name, col.len);
      return null;
    }
  }
  // generic layer (fnBlank / fnLen formats) for unbound columns
  if (col.len && v.length > col.len) return lenMsg(col.name, col.len);
  if (col.type.toUpperCase()==="BOOLEAN" && v!=="" && !allow(v,"tTfF"))
    return col.name + " can contain t or f only.";
  return null;
}
function commaWarning(v, col){
  v = (v||"").trim();
  if (!v.includes(",")) return null;
  return col.name + " contains a comma. BAT files are comma-separated, so this value will split into two fields on import (the original xlt allowed this).";
}
function validateCell(td, v, col){
  const ctx = current ? {sheet:current, devType:state[current].devType, dummy:state[current].dummy} : {sheet:"", devType:"phone", dummy:false};
  const err = validateCellVBA(v, col, ctx);
  td.classList.toggle("invalid", !!err);
  td.title = err || commaWarning(v, col) || "";
}
/* ── cross-field rules, ported per sheet ────────────────────────────── */
function groupCols(columns, re){
  // returns {N: [colIdx,...]} for columns whose name ends in a line number
  const groups = {};
  columns.forEach((c, ci) => {
    const m = c.name.match(re);
    if (m){ const n = m[1]; (groups[n] = groups[n]||[]).push(ci); }
  });
  return groups;
}
function crossFieldVBA(sheet, columns, row, ctx, rowLabel, errs){
  const S = sheet.toUpperCase();
  const val = ci => (row[ci]||"").trim();
  const findCol = re => columns.findIndex(c => re.test(c.name.toUpperCase()));

  if (S==="PHONES"||S==="PHONES-USERS"||S==="USER DEVICE PROFILES"||S==="REMOTE DESTINATION PROFILE"){
    // populated line group needs its Directory Number (fnCheckOtherLineFields)
    const lineGroups = groupCols(columns, /(\d+)$/);
    for (const [n, idxs] of Object.entries(lineGroups)){
      const dnIdx = idxs.find(ci => /^(INTERCOM )?DIRECTORY NUMBER \d+$/.test(columns[ci].name.toUpperCase()));
      if (dnIdx===undefined) continue;
      const others = idxs.filter(ci => ci!==dnIdx &&
        /DIRECTORY NUMBER|DISPLAY|LINE TEXT LABEL|FORWARD|CALL PICKUP|ROUTE PARTITION|ALERTING|AUTO ANSWER|URI /.test(columns[ci].name.toUpperCase()));
      if (val(dnIdx)==="" && others.some(ci => val(ci)!==""))
        errs.push(rowLabel + "The Directory Number is mandatory field.");
    }
    // Number of Lines vs populated DNs (non-Phones flexi sheets)
    const nolIdx = findCol(/^NUMBER OF LINES/);
    if (nolIdx>=0 && S!=="PHONES"){
      if (val(nolIdx)==="") errs.push(rowLabel + "Number of Lines field is mandatory");
      else {
        const nol = parseInt(val(nolIdx),10)||0;
        if (nol > ctx.fmtLines)
          errs.push(rowLabel + "The value entered in the Number of Lines field should be less than or equal to that entered in the Phone Lines Text box.");
        const dnIdxs = columns.map((c,ci)=>({c,ci})).filter(x=>/^DIRECTORY NUMBER \d+$/.test(x.c.name.toUpperCase()));
        const blanksInRange = dnIdxs.slice(0, nol).filter(x=>val(x.ci)==="").length;
        if (blanksInRange>0 && dnIdxs.length>0)
          errs.push(rowLabel + "The given Number of lines in the field and Directory Number do not match.");
      }
    }
    // BLF sets: Destination and Directory Number are mutually exclusive
    const blfGroups = groupCols(columns, /^BUSY LAMP FIELD .+?(\d+)$/i);
    for (const idxs of Object.values(blfGroups)){
      const dest = idxs.find(ci => /BUSY LAMP FIELD DESTINATION/.test(columns[ci].name.toUpperCase()));
      const dn   = idxs.find(ci => /BUSY LAMP FIELD DIRECTORY NUMBER/.test(columns[ci].name.toUpperCase()));
      if (dest!==undefined && dn!==undefined && val(dest)!=="" && val(dn)!=="")
        errs.push(rowLabel + " Give Values for either BUSY LAMP FIELD DIRECTORY NUMBER or BUSY LAMP FIELD DESTINATION");
    }
  }

  if (S==="ADD LINES"||S==="ADD INTERCOM"){
    const liGroups = groupCols(columns, /LINE INDEX (\d+)$/i);
    const nums = Object.keys(liGroups).map(Number).sort((a,b)=>a-b);
    let prevIdxBlank = false;
    for (const n of nums){
      const liIdx = columns.findIndex(c => new RegExp("LINE INDEX "+n+"$","i").test(c.name));
      const setIdxs = columns.map((c,ci)=>ci).filter(ci =>
        ci!==liIdx && new RegExp("\\b"+n+"$").test(columns[ci].name) &&
        !/MAC ADDRESS/i.test(columns[ci].name));
      const setFilled = setIdxs.some(ci=>val(ci)!=="");
      if (val(liIdx)==="" && setFilled){
        errs.push(rowLabel + "Enter the Line Index value first.");
      } else if (val(liIdx)!=="" && prevIdxBlank){
        errs.push(rowLabel + "Enter the values for the previous Line Index first.");
      }
      if (val(liIdx)!==""){
        const dnIdx = setIdxs.find(ci => /DIRECTORY NUMBER/i.test(columns[ci].name));
        if (dnIdx!==undefined && val(dnIdx)==="")
          errs.push(rowLabel + (S==="ADD INTERCOM" ? "Intercom Directory Number cannot be blank." : "Directory Number cannot be blank."));
      }
      prevIdxBlank = val(liIdx)==="";
    }
  }

  if (S==="INSERT FAC"){
    const nameIdx = findCol(/AUTHORIZATION CODE NAME/), lvlIdx = findCol(/AUTHORIZATIONLEVEL|AUTHORIZATION LEVEL/);
    if (nameIdx>=0 && lvlIdx>=0 && val(nameIdx)==="" && val(lvlIdx)==="")
      errs.push(rowLabel + "At least one value among Authorization Code Name and Authorization Level is mandatory.");
  }

  if (S==="PHONE MIGRATION"){
    const oldIdx = findCol(/OLD DEVICE NAME/), newIdx = findCol(/NEW DEVICE MAC ADDRESS/);
    if ((oldIdx>=0 && val(oldIdx)==="") || (newIdx>=0 && val(newIdx)===""))
      errs.push(rowLabel + "Both the MAC Address Field Are Mandatory.");
  }

  if (S==="CUSTOM MANAGERS-ASSISTANTS"){
    const mgrDNs = columns.map((c,ci)=>ci).filter(ci=>/MANAGER LINE DN/i.test(columns[ci].name));
    const pxyDNs = columns.map((c,ci)=>ci).filter(ci=>/PROXY LINE DN/i.test(columns[ci].name));
    if (mgrDNs.length && !mgrDNs.some(ci=>val(ci)!=="")) errs.push(rowLabel + "Atleast one Manager line DN is mandatory.");
    if (pxyDNs.length && !pxyDNs.some(ci=>val(ci)!=="")) errs.push(rowLabel + "Atleast one proxy line DN is mandatory.");
  }

  if (S==="TRUST GROUP"){
    const t = findCol(/^TRUSTED$|^TRUST$/);
    if (t>=0 && val(t)==="") errs.push(rowLabel + "Trust is a Mandatory field.");
  }
  if (S==="TRUST ELEMENT"){
    const et = findCol(/ELEMENT TYPE/), tg = findCol(/TRUST GROUP/);
    if (et>=0 && val(et)==="") errs.push(rowLabel + "Element Type is Mandatory.");
    if (tg>=0 && val(tg)==="") errs.push(rowLabel + "Trust Group is Mandatory.");
    if (tg>=0 && deny(val(tg), M.xGrpName)) errs.push(rowLabel + "Trust Group Name "+GRPMSG);
  }
  if (S==="FALLBACK PROFILE"){
    for (const [re,msg] of [[/QOS SENSITIVITY/,"FallBack Qos Sensitivity level is Mandatory."],
                            [/FALLBACK CALL CSS/,"FallBack Call CSS is Mandatory."],
                            [/CALL ANSWER TIMER/,"FallBack Call Answer Timer is Mandatory."],
                            [/NUMBER OF DIGITS FOR CALLER ID/,"Number of Digits for Caller Id Partial Match is Mandatory."]]){
      const ci = findCol(re);
      if (ci>=0 && val(ci)==="") errs.push(rowLabel + msg);
    }
  }
  if (S==="END USER CAPF PROFILE"){
    for (const [re,msg] of [[/END USER ID/,"End User ID is Mandatory."],
                            [/INSTANCE ID/,"Instance ID is Mandatory."],
                            [/CERTIFICATE OPERATION/,"Certificate Operation is Mandatory."]]){
      const ci = findCol(re);
      if (ci>=0 && val(ci)==="") errs.push(rowLabel + msg);
    }
  }
  if (S==="MOBILITY PROFILE"){
    const n = findCol(/MOBILITY PROFILE NAME/), o = findCol(/MOBILE CLIENT CALLING OPTION/);
    if (n>=0 && val(n)==="") errs.push(rowLabel + "Name is Mandatory.");
    if (o>=0 && val(o)==="") errs.push(rowLabel + "Mobile Client Calling Option is Mandatory.");
  }
  if (S==="INSERT INFRASTRUCTURE DEVICE"){
    const n = findCol(/ACCESSPOINT OR SWITCH NAME/);
    if (n>=0 && val(n)==="") errs.push(rowLabel + "AccessPoint or Switch Name cannot be blank.");
    const ids = [findCol(/IPV4 ADDRESS/), findCol(/IPV6 ADDRESS/), findCol(/^BSSID/)].filter(i=>i>=0);
    if (ids.length && !ids.some(ci=>val(ci)!==""))
      errs.push(rowLabel + "No identifying information (IPv4,IPv6,BSSID) provided for Infrastructure Device");
  }
  if (S==="CALL PICKUP GROUP"){
    const others = columns.map((c,ci)=>ci).filter(ci=>/OTHER PICKUP GROUP NAME/i.test(columns[ci].name));
    let prevBlank = false;
    for (const ci of others){
      if (val(ci)!=="" && prevBlank)
        errs.push(rowLabel + "Enter the preceding Other Pickup Group Name first before adding the next Other Pickup Group Name.");
      prevBlank = val(ci)==="";
    }
  }
}

/* ── status / errors ────────────────────────────────────────────────── */
function updateStatus(){
  const st = state[current];
  const n = st.rows.filter(r => r.some(v => (v||"").trim() !== "")).length;
  document.getElementById("status").textContent =
    st.columns.length + " cols · " + n + " record" + (n===1?"":"s");
}
function showErrors(list, warns){
  const el = document.getElementById("errs");
  warns = warns || [];
  if (!list.length && !warns.length){ el.classList.remove("open"); el.innerHTML=""; return; }
  const esc = s => s.replace(/</g,"&lt;");
  el.innerHTML =
    list.slice(0,200).map(e => "<div>"+esc(e)+"</div>").join("") +
    warns.slice(0,50).map(w => "<div class='warn'>Warning — "+esc(w)+"</div>").join("");
  el.classList.add("open");
}

/* ── export: replicate frmFileName header normalization ─────────────── */
function normalizeHeader(rawName, sheet){
  let v = rawName.trim().toUpperCase();
  const U = sheet.toUpperCase();
  if (v.includes("DESK PHONE NAME")) return "DESK PHONE NAME";
  if (v.includes("MAC ADDRESS")){
    if (U === "PHONES" || U === "CATALYST 6000 (FXS) PORTS" || U === "PHONES-USERS") return "MAC ADDRESS";
    if (U === "ADD INTERCOM" || U === "ADD LINES") return "MAC ADDRESS";       // Phones target (UDP target uses Device Profile Name column instead)
    if (/^VG(224|310|320|350|450|420)$|^VG202-204$|^ISR4461$/.test(U)) return "DOMAIN NAME";
    if (U === "PHONE MIGRATION") return v;
    return "DEVICE NAME";
  }
  if (v.includes("USER DEVICE PROFILE")) return "USER DEVICE PROFILE";
  if (v.startsWith("USERID")) return "USER ID";
  if (v.includes("MANAGER") && U !== "DEFAULT MANAGERS-ASSISTANTS" && U !== "CUSTOM MANAGERS-ASSISTANTS")
    return "MANAGER USER ID";
  if (v.includes("REMOTE DESTINATION PROFILE")) return "REMOTE DESTINATION PROFILE NAME";
  return v;
}
function e164Renumber(headers){
  const counters = {};
  const map = [
    ["E164 NUMBER MASK", "E.164 NUMBER MASK"],
    ["E164 IS URGENT", "E.164 IS URGENT"],
    ["E164 ADD TO LOCAL ROUTE PARTITION", "E.164 ADD TO LOCAL ROUTE PARTITION"],
    ["E164 ADVERTISE VIA GLOBALLY", "E.164 ADVERTISE VIA GLOBALLY"],
    ["E164 ROUTE PARTITION", "E.164 ROUTE PARTITION"],
  ];
  return headers.map(h => {
    for (const [pat, out] of map){
      if (h.includes(pat)){
        counters[pat] = (counters[pat] || 0) + 1;
        return out + " " + counters[pat];
      }
    }
    return h;
  });
}

function buildExport(){
  const st = state[current];
  const sheet = current;
  const errs = [], warns = [];
  const macCol = st.columns.findIndex(c => /MAC ADDRESS|OLD DEVICE NAME/i.test(c.name));
  const dataRows = st.rows.filter(r => r.some(v => (v||"").trim() !== ""));
  if (!dataRows.length){ errs.push("There is no data to export."); return {errs, warns}; }
  const ctx = {sheet, devType: st.devType || "phone", dummy: st.dummy,
               fmtLines: st.fmt ? (st.fmt.counts.lines||0) : 0};
  let macFilledWithDummy = false;

  dataRows.forEach((row, i) => {
    const rowLabel = "Row " + (i+1) + ": ";
    st.columns.forEach((c, ci) => {
      const v = (row[ci] || "").trim();
      if (st.dummy && ci === macCol){ if (v!=="") macFilledWithDummy = true; return; }
      const msg = validateCellVBA(v, c, ctx);
      if (msg){ errs.push(rowLabel + msg); return; }
      // generic mandatory (fnBlank format) for columns the bindings left blank-tolerant
      if (c.req && v === "" && !(st.dummy && ci === macCol)){
        // cross-field/sheet rules own some blanks with their exact messages
        const owned = /NUMBER OF LINES|AUTHORIZATION CODE NAME|AUTHORIZATIONLEVEL|MANAGER LINE DN|PROXY LINE DN/i.test(c.name)
          || NAME_SHEET_RULES[sheet] || /MAC ADDRESS|DEVICE PROFILE NAME|PICKUP GROUP|REMOTE DESTINATION PROFILE NAME|USERID|USER ID|OLD DEVICE NAME|NEW DEVICE MAC ADDRESS/i.test(c.name)
          || ["Trust Element","End User CAPF Profile","Mobility Profile","Insert Infrastructure Device","FallBack Profile"].includes(sheet);
        if (!owned) errs.push(rowLabel + blankMsg(c.name));
      }
      const w = commaWarning(v, c);
      if (w) warns.push(rowLabel + w);
    });
    crossFieldVBA(sheet, st.columns, row, ctx, rowLabel, errs);
  });
  if (errs.length) return {errs, warns};

  // the macro's dummy-MAC confirmation, verbatim
  if (st.dummy && macFilledWithDummy){
    const isCti = (st.devType === "cti") || /CTI PORT/i.test(sheet);
    const msg = isCti
      ? "You have opted for Dummy Device Name. Dummy Device\n names will not be exported. Continue?"
      : "You have opted for Dummy MAC Address. MAC Address\n values will not be exported. Continue?";
    if (!confirm(msg)) return {errs:["Export cancelled."], warns};
  }

  let headers = st.columns.map(c => normalizeHeader(c.name, sheet));
  headers = e164Renumber(headers);
  const lines = [headers.join(",")];
  for (const row of dataRows){
    const vals = st.columns.map((c, ci) => {
      let v = (row[ci] || "").trim();
      if (st.dummy && ci === macCol) v = "";     // macro omits MAC when dummy checked
      return v;
    });
    lines.push(vals.join(","));
  }
  return {errs: [], warns, text: lines.join("\r\n") + "\r\n", count: dataRows.length};
}

function download(name, text){
  const blob = new Blob([text], {type:"text/plain"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name.endsWith(".txt") ? name : name + ".txt";
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}

/* ── format builder UI ──────────────────────────────────────────────── */
function renderFmt(){
  const fmtEl = document.getElementById("fmt");
  const st = state[current];
  if (!isDynamic(current)){ fmtEl.classList.remove("open"); return; }
  const txn = txnKey(current);
  const fill = (id, cat, chosen) => {
    const box = document.getElementById(id);
    const fields = dbFields(txn, cat);
    box.parentElement.style.display = fields.length ? "" : "none";
    box.innerHTML = "";
    for (const f of fields){
      const lab = document.createElement("label");
      const cb = document.createElement("input");
      cb.type = "checkbox"; cb.value = f.name;
      cb.checked = chosen.some(x => x.name === f.name);
      lab.appendChild(cb);
      lab.append(" " + f.name + (f.req ? " *" : ""));
      box.appendChild(lab);
    }
  };
  fill("listDevice", "Device", st.fmt.device);
  fill("listLine", "Line", st.fmt.line);
  fill("listIntercom", "Intercom", st.fmt.intercom);

  const counts = document.getElementById("countInputs");
  counts.innerHTML = "";
  const addCount = (key, label, max) => {
    const lab = document.createElement("label");
    lab.append(label + " ");
    const inp = document.createElement("input");
    inp.type = "number"; inp.min = 0; inp.max = max; inp.value = st.fmt.counts[key] || 0;
    inp.dataset.key = key;
    lab.appendChild(inp); counts.appendChild(lab);
  };
  const hasLine = dbFields(txn, "Line").length, hasInt = dbFields(txn, "Intercom").length;
  if (hasLine) addCount("lines", "Number of Lines", 50);
  if (hasInt)  addCount("intercoms", "Number of Intercoms", 10);
  for (const [key, grp] of Object.entries(MODEL.repeatGroups))
    if (grp.sheets.includes(current)) addCount(key, grp.label, 100);
}

document.getElementById("btnApplyFmt").addEventListener("click", () => {
  const st = state[current];
  const txn = txnKey(current);
  const collect = (id, cat) => {
    const names = [...document.querySelectorAll("#"+id+" input:checked")].map(i => i.value);
    return dbFields(txn, cat).filter(f => names.includes(f.name));
  };
  st.fmt.device = collect("listDevice", "Device");
  st.fmt.line = collect("listLine", "Line");
  st.fmt.intercom = collect("listIntercom", "Intercom");
  document.querySelectorAll("#countInputs input").forEach(i => {
    st.fmt.counts[i.dataset.key] = Math.max(0, parseInt(i.value || "0", 10) || 0);
  });
  if (st.fmt.counts.lines > 0 && !st.fmt.line.length)
    st.fmt.line = dbFields(txn, "Line").filter(f => MODEL.defaultLineFields.includes(f.name));
  if (st.fmt.counts.intercoms > 0 && !st.fmt.intercom.length)
    st.fmt.intercom = dbFields(txn, "Intercom");
  rebuildDynamicColumns(current);
  renderGrid();
  showErrors([]);
});

/* ── tab switching ──────────────────────────────────────────────────── */
function selectSheet(name){
  current = name;
  initSheet(name);
  document.querySelectorAll("#tabs button").forEach(b =>
    b.classList.toggle("active", b.dataset.name === name));
  document.getElementById("sheetTitle").textContent = name;
  const dyn = isDynamic(name);
  document.getElementById("btnFmt").hidden = !dyn;
  document.getElementById("fmt").classList.remove("open");
  if (dyn) renderFmt();
  const dm = MODEL.dummyMacSheets.includes(name);
  document.getElementById("dummyWrap").hidden = !dm;
  document.getElementById("dummyMac").checked = state[name].dummy;
  const dtShown = name === "Phones" || name === "Phones-Users";
  document.getElementById("devTypeWrap").hidden = !dtShown;
  document.getElementById("devType").value = state[name].devType || "phone";
  showErrors([]);
  renderGrid();
}
document.getElementById("devType").addEventListener("change", e => {
  state[current].devType = e.target.value;
  renderGrid();   // revalidate cells under the new device type
});

function buildRail(){
  const tabs = document.getElementById("tabs");
  tabs.innerHTML = "";
  const q = document.getElementById("railSearch").value.trim().toLowerCase();
  for (const s of MODEL.sheets){
    if (q && !s.name.toLowerCase().includes(q)) continue;
    const b = document.createElement("button");
    b.dataset.name = s.name; b.setAttribute("role","tab");
    b.textContent = s.name;
    if (isDynamic(s.name)){
      const tag = document.createElement("span");
      tag.className = "dyn"; tag.textContent = "FMT";
      b.appendChild(tag);
    }
    if (s.name === current) b.classList.add("active");
    b.addEventListener("click", () => selectSheet(s.name));
    tabs.appendChild(b);
  }
}
document.getElementById("railSearch").addEventListener("input", buildRail);

/* ── toolbar ────────────────────────────────────────────────────────── */
function addRows(n){
  const st = state[current];
  for (let i=0;i<n;i++) st.rows.push(new Array(st.columns.length).fill(""));
  renderGrid();
}
document.getElementById("btnAddRow").addEventListener("click", () => addRows(1));
document.getElementById("btnAddRows").addEventListener("click", () => addRows(10));
document.getElementById("btnClear").addEventListener("click", () => {
  if (confirm("Clear all data on '" + current + "'?")){
    state[current].rows = []; renderGrid(); showErrors([]);
  }
});
document.getElementById("btnFmt").addEventListener("click", () =>
  document.getElementById("fmt").classList.toggle("open"));
document.getElementById("dummyMac").addEventListener("change", e =>
  state[current].dummy = e.target.checked);

/* paste dialog */
const dlgPaste = document.getElementById("dlgPaste");
document.getElementById("btnPaste").addEventListener("click", () => {
  document.getElementById("pasteArea").value = ""; dlgPaste.showModal();
  document.getElementById("pasteArea").focus();
});
document.getElementById("pasteCancel").addEventListener("click", () => dlgPaste.close());
document.getElementById("pasteGo").addEventListener("click", () => {
  const st = state[current];
  const raw = document.getElementById("pasteArea").value.replace(/\r/g,"");
  const lines = raw.split("\n").filter(l => l.trim().length);
  for (const line of lines){
    const cells = line.includes("\t") ? line.split("\t") : line.split(",");
    const row = new Array(st.columns.length).fill("");
    cells.forEach((v,i) => { if (i < row.length) row[i] = v.trim(); });
    st.rows.push(row);
  }
  dlgPaste.close(); renderGrid();
});

/* export dialog */
const dlgExport = document.getElementById("dlgExport");
document.getElementById("btnExport").addEventListener("click", () => {
  const res = buildExport();
  showErrors(res.errs, res.warns);
  if (res.errs.length) return;
  document.getElementById("exportName").value =
    current.replace(/[^A-Za-z0-9_-]+/g,"") || "BATexport";
  document.getElementById("exportSummary").textContent =
    res.count + " record(s), " + state[current].columns.length +
    " column(s). Output is a comma-separated .txt with a BAT header line (CRLF line endings)." +
    (res.warns.length ? " " + res.warns.length + " warning(s) — see the panel below the grid." : "");
  dlgExport.showModal();
});
document.getElementById("exportCancel").addEventListener("click", () => dlgExport.close());
document.getElementById("exportGo").addEventListener("click", () => {
  const nm = document.getElementById("exportName").value.trim();
  if (!nm || /[\\\/]/.test(nm)){ alert("Enter a valid filename."); return; }
  const res = buildExport();
  if (res.errs.length){ showErrors(res.errs, res.warns); dlgExport.close(); return; }
  download(nm, res.text);
  dlgExport.close();
});

/* help */
const dlgHelp = document.getElementById("dlgHelp");
document.getElementById("btnHelp").addEventListener("click", () => {
  const info = "Generated from " + MODEL.source + " on " + MODEL.generated +
    ".\nThis tool replicates the Cisco BAT xlt macros in plain HTML/JS — no macros or ActiveX required.\n" +
    "Validation rules, character masks, and error messages are ported 1:1 from the xlt's VBA code, " +
    "including cross-field rules (Number of Lines matching, Line Index sequencing, BLF either/or, " +
    "FAC either/or, dummy-MAC handling, and per-device-type name rules).\n" +
    "Notes: (1) A few VBA validators failed silently with no message; this tool supplies message text " +
    "for those (IPv4/IPv6/BSSID formats). (2) The original xlt allows commas in some fields " +
    "(Description, Display, labels, names) even though commas corrupt the comma-separated output; " +
    "this tool permits them identically but shows a non-blocking warning.\n\n";
  document.getElementById("helpBody").textContent = info + MODEL.help.join("\n\n");
  dlgHelp.showModal();
});
document.getElementById("helpClose").addEventListener("click", () => dlgHelp.close());

/* ── boot ───────────────────────────────────────────────────────────── */
document.getElementById("srcInfo").textContent = MODEL.source;
buildRail();
selectSheet(MODEL.sheets[0].name);
</script>
</body>
</html>
"""


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_suffix(".html")
    model = read_workbook(str(src))
    html = HTML_TEMPLATE.replace("__SOURCE__", model["source"])
    html = html.replace("__MODEL__", json.dumps(model, separators=(",", ":")))
    out.write_text(html, encoding="utf-8")
    print(f"Wrote {out}  ({out.stat().st_size:,} bytes, "
          f"{len(model['sheets'])} tabs, {len(model['database'])} database fields)")


if __name__ == "__main__":
    main()
