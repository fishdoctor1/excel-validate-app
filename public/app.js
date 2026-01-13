let rows = [];        // source-of-truth rows (includes excel_row + table_selector + C..AI)
let columns = [];     // C..AI column list
let displayRows = []; // display table (dates already toISOString in server)

const $ = (id) => document.getElementById(id);

function setStatus(text, cls="muted") {
  const el = $("status");
  el.className = cls;
  el.textContent = text;
}

function setSqlStatus(text, cls="muted") {
  const el = $("sqlStatus");
  el.className = cls;
  el.textContent = text;
}

function buildTable() {
  const tbl = $("tbl");
  tbl.innerHTML = "";

  // Header
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");

  const thRow = document.createElement("th");
  thRow.textContent = "excel_row";
  trh.appendChild(thRow);

  for (const c of columns) {
    const th = document.createElement("th");
    th.textContent = c;
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  tbl.appendChild(thead);

  // Body
  const tbody = document.createElement("tbody");
  for (let i = 0; i < displayRows.length; i++) {
    const rDisp = displayRows[i];
    const rSrc = rows[i]; // keep aligned (same ordering from server)

    const tr = document.createElement("tr");

    const tdRow = document.createElement("td");
    tdRow.textContent = String(rDisp.excel_row);
    tr.appendChild(tdRow);

    for (const c of columns) {
      const td = document.createElement("td");
      td.setAttribute("contenteditable", "true");
      td.dataset.rowIndex = String(i);
      td.dataset.col = c;

      const v = rDisp[c];
      td.textContent = (v === null || v === undefined) ? "" : String(v);

      td.addEventListener("input", (e) => {
        const rowIndex = Number(e.target.dataset.rowIndex);
        const col = e.target.dataset.col;
        const text = e.target.textContent;

        // IMPORTANT:
        // - empty UI cell => null
        // - do NOT trim/upper/lower
        // - keep user override exactly as typed
        if (text === "") {
          rows[rowIndex][col] = null;
          displayRows[rowIndex][col] = null;
        } else {
          rows[rowIndex][col] = text;
          displayRows[rowIndex][col] = text;
        }
      });

      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  tbl.appendChild(tbody);
}

$("btnExtract").addEventListener("click", async () => {
  const f = $("file").files?.[0];
  if (!f) return setStatus("กรุณาเลือกไฟล์ก่อน", "error");

  setStatus("กำลัง extract...", "muted");
  setSqlStatus("");
  $("sql").value = "";

  const fd = new FormData();
  fd.append("file", f);

  const resp = await fetch("/api/extract", { method: "POST", body: fd });
  const data = await resp.json();

  if (!data.ok) {
    setStatus(data.error || "extract failed", "error");
    $("extractWrap").style.display = "none";
    return;
  }

  rows = data.rows;
  columns = data.columns;
  displayRows = data.displayRows;

  setStatus(`Extract สำเร็จ: ${rows.length} rows`, "ok");
  $("extractWrap").style.display = "block";
  buildTable();
});

$("btnSql").addEventListener("click", async () => {
  const confirmText = $("confirmText").value;

  setSqlStatus("กำลัง generate SQL...", "muted");
  $("sql").value = "";

  const resp = await fetch("/api/generate-sql", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rows, confirmText })
  });

  const data = await resp.json();
  if (!data.ok) {
    setSqlStatus(data.error || "generate failed", "error");
    return;
  }

  setSqlStatus("Generate SQL สำเร็จ", "ok");
  $("sql").value = data.sql;
});
