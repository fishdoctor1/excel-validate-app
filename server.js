const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const { parse } = require("csv-parse/sync");
const ExcelJS = require("exceljs");

const app = express();
app.use(express.json({ limit: "20mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ storage: multer.memoryStorage() });

/**
 * Column mapping (Template A->AI)
 * C..AI = main extract columns (C->AI) strictly ordered
 */
const C_TO_AI = [
  "event_module", // C
  "event_module_action", // D
  "document_no", // E
  "prefix_document_no", // F
  "suffix_document_no", // G
  "description", // H
  "account_event", // I
  "account_code", // J
  "account_type", // K
  "status", // L
  "company", // M
  "channel", // N
  "product", // O
  "bau", // P
  "payment_type", // Q
  "interco", // R
  "fund", // S
  "amount", // T
  "policy_no", // U
  "policy_group_no", // V
  "policy_account_no", // W
  "policy_year", // X
  "account_effective_date", // Y
  "policy_effective_date", // Z
  "pn_no", // AA
  "sn_no", // AB
  "failure_type", // AC
  "failure_reason", // AD
  "plan_type", // AE
  "reinsurance_type", // AF
  "period", // AG
  "issue_date", // AH
  "payment_date", // AI
];

const ALL_HEADERS_TEMPLATE = [
  ...C_TO_AI, // C..AI
];

const TEMPLATE_DIR = path.join(__dirname, "templates");
const TEMPLATE_XLSX = path.join(
  TEMPLATE_DIR,
  "accounting_validation_template.xlsx"
);
const TEMPLATE_CSV = path.join(
  TEMPLATE_DIR,
  "accounting_validation_template.csv"
);

function SQL_VALIDATION_SUFFIX_SUCCESS(inputdate) {
    return `SELECT
    e.document_no,
    e.account_code,
    e.plan_type,
    e.reinsurance_type,
    CASE
      WHEN db.document_no IS NULL THEN ARRAY[
        'event_module','event_module_action','document_no','prefix_document_no','suffix_document_no','description',
        'account_event','account_code','account_type','status','company','channel','product','bau','payment_type',
        'interco','fund','amount','policy_no','policy_group_no','policy_account_no','policy_year',
        'account_effective_date','policy_effective_date','pn_no','sn_no','failure_type','failure_reason',
        'plan_type','reinsurance_type','period','issue_date','payment_date'
      ]::text[]
      ELSE array_remove(ARRAY[
        CASE WHEN NULLIF(db.event_module::text,'') IS DISTINCT FROM NULLIF(e.event_module::text,'') THEN 'event_module' END,
        CASE WHEN NULLIF(db.event_module_action::text,'') IS DISTINCT FROM NULLIF(e.event_module_action::text,'') THEN 'event_module_action' END,
        CASE WHEN NULLIF(db.document_no::text,'') IS DISTINCT FROM NULLIF(e.document_no::text,'') THEN 'document_no' END,
        CASE WHEN NULLIF(db.prefix_document_no::text,'') IS DISTINCT FROM NULLIF(e.prefix_document_no::text,'') THEN 'prefix_document_no' END,
        CASE WHEN NULLIF(db.suffix_document_no::text,'') IS DISTINCT FROM NULLIF(e.suffix_document_no::text,'') THEN 'suffix_document_no' END,
        CASE WHEN NULLIF(db.description::text,'') IS DISTINCT FROM NULLIF(e.description::text,'') THEN 'description' END,
        CASE WHEN NULLIF(db.account_event::text,'') IS DISTINCT FROM NULLIF(e.account_event::text,'') THEN 'account_event' END,
        CASE WHEN NULLIF(db.account_code::text,'') IS DISTINCT FROM NULLIF(e.account_code::text,'') THEN 'account_code' END,
        CASE WHEN NULLIF(db.account_type::text,'') IS DISTINCT FROM NULLIF(e.account_type::text,'') THEN 'account_type' END,
        CASE WHEN NULLIF(db.status::text,'') IS DISTINCT FROM NULLIF(e.status::text,'') THEN 'status' END,
        CASE WHEN NULLIF(db.company::text,'') IS DISTINCT FROM NULLIF(e.company::text,'') THEN 'company' END,
        CASE WHEN NULLIF(db.channel::text,'') IS DISTINCT FROM NULLIF(e.channel::text,'') THEN 'channel' END,
        CASE WHEN NULLIF(db.product::text,'') IS DISTINCT FROM NULLIF(e.product::text,'') THEN 'product' END,
        CASE WHEN NULLIF(db.bau::text,'') IS DISTINCT FROM NULLIF(e.bau::text,'') THEN 'bau' END,
        CASE WHEN NULLIF(db.payment_type::text,'') IS DISTINCT FROM NULLIF(e.payment_type::text,'') THEN 'payment_type' END,
        CASE WHEN NULLIF(db.interco::text,'') IS DISTINCT FROM NULLIF(e.interco::text,'') THEN 'interco' END,
        CASE WHEN NULLIF(db.fund::text,'') IS DISTINCT FROM NULLIF(e.fund::text,'') THEN 'fund' END,
        CASE WHEN trunc(db.amount::numeric, 2) IS DISTINCT FROM e.amount THEN 'amount' END,
        CASE WHEN NULLIF(db.policy_no::text,'') IS DISTINCT FROM NULLIF(e.policy_no::text,'') THEN 'policy_no' END,
        CASE WHEN NULLIF(db.policy_group_no::text,'') IS DISTINCT FROM NULLIF(e.policy_group_no::text,'') THEN 'policy_group_no' END,
        CASE WHEN NULLIF(db.policy_account_no::text,'') IS DISTINCT FROM NULLIF(e.policy_account_no::text,'') THEN 'policy_account_no' END,
        CASE WHEN NULLIF(db.policy_year::text,'')::int IS DISTINCT FROM e.policy_year THEN 'policy_year' END,
        CASE WHEN db.account_effective_date::timestamp IS DISTINCT FROM e.account_effective_date THEN 'account_effective_date' END,
        CASE WHEN db.policy_effective_date::timestamp IS DISTINCT FROM e.policy_effective_date THEN 'policy_effective_date' END,
        CASE WHEN NULLIF(db.pn_no::text,'') IS DISTINCT FROM NULLIF(e.pn_no::text,'') THEN 'pn_no' END,
        CASE WHEN NULLIF(db.sn_no::text,'') IS DISTINCT FROM NULLIF(e.sn_no::text,'') THEN 'sn_no' END,
        CASE WHEN NULLIF(db.plan_type::text,'') IS DISTINCT FROM NULLIF(e.plan_type::text,'') THEN 'plan_type' END,
        CASE WHEN NULLIF(db.reinsurance_type::text,'') IS DISTINCT FROM NULLIF(e.reinsurance_type::text,'') THEN 'reinsurance_type' END,
        CASE WHEN NULLIF(db.period::text,'')::int IS DISTINCT FROM e.period THEN 'period' END,
        CASE WHEN db.issue_date::timestamp IS DISTINCT FROM e.issue_date THEN 'issue_date' END,
        CASE WHEN db.payment_date::timestamp IS DISTINCT FROM e.payment_date THEN 'payment_date' END
      ]::text[], NULL)
    END AS invalid_columns,
    (CASE
      WHEN db.document_no IS NULL THEN FALSE
      ELSE CARDINALITY(
        array_remove(ARRAY[
          CASE WHEN NULLIF(db.event_module::text,'') IS DISTINCT FROM NULLIF(e.event_module::text,'') THEN 'event_module' END,
          CASE WHEN NULLIF(db.event_module_action::text,'') IS DISTINCT FROM NULLIF(e.event_module_action::text,'') THEN 'event_module_action' END,
          CASE WHEN NULLIF(db.document_no::text,'') IS DISTINCT FROM NULLIF(e.document_no::text,'') THEN 'document_no' END,
          CASE WHEN NULLIF(db.prefix_document_no::text,'') IS DISTINCT FROM NULLIF(e.prefix_document_no::text,'') THEN 'prefix_document_no' END,
          CASE WHEN NULLIF(db.suffix_document_no::text,'') IS DISTINCT FROM NULLIF(e.suffix_document_no::text,'') THEN 'suffix_document_no' END,
          CASE WHEN NULLIF(db.description::text,'') IS DISTINCT FROM NULLIF(e.description::text,'') THEN 'description' END,
          CASE WHEN NULLIF(db.account_event::text,'') IS DISTINCT FROM NULLIF(e.account_event::text,'') THEN 'account_event' END,
          CASE WHEN NULLIF(db.account_code::text,'') IS DISTINCT FROM NULLIF(e.account_code::text,'') THEN 'account_code' END,
          CASE WHEN NULLIF(db.account_type::text,'') IS DISTINCT FROM NULLIF(e.account_type::text,'') THEN 'account_type' END,
          CASE WHEN NULLIF(db.status::text,'') IS DISTINCT FROM NULLIF(e.status::text,'') THEN 'status' END,
          CASE WHEN NULLIF(db.company::text,'') IS DISTINCT FROM NULLIF(e.company::text,'') THEN 'company' END,
          CASE WHEN NULLIF(db.channel::text,'') IS DISTINCT FROM NULLIF(e.channel::text,'') THEN 'channel' END,
          CASE WHEN NULLIF(db.product::text,'') IS DISTINCT FROM NULLIF(e.product::text,'') THEN 'product' END,
          CASE WHEN NULLIF(db.bau::text,'') IS DISTINCT FROM NULLIF(e.bau::text,'') THEN 'bau' END,
          CASE WHEN NULLIF(db.payment_type::text,'') IS DISTINCT FROM NULLIF(e.payment_type::text,'') THEN 'payment_type' END,
          CASE WHEN NULLIF(db.interco::text,'') IS DISTINCT FROM NULLIF(e.interco::text,'') THEN 'interco' END,
          CASE WHEN NULLIF(db.fund::text,'') IS DISTINCT FROM NULLIF(e.fund::text,'') THEN 'fund' END,
          CASE WHEN trunc(db.amount::numeric, 2) IS DISTINCT FROM e.amount THEN 'amount' END,
          CASE WHEN NULLIF(db.policy_no::text,'') IS DISTINCT FROM NULLIF(e.policy_no::text,'') THEN 'policy_no' END,
          CASE WHEN NULLIF(db.policy_group_no::text,'') IS DISTINCT FROM NULLIF(e.policy_group_no::text,'') THEN 'policy_group_no' END,
          CASE WHEN NULLIF(db.policy_account_no::text,'') IS DISTINCT FROM NULLIF(e.policy_account_no::text,'') THEN 'policy_account_no' END,
          CASE WHEN NULLIF(db.policy_year::text,'')::int IS DISTINCT FROM e.policy_year THEN 'policy_year' END,
          CASE WHEN db.account_effective_date::timestamp IS DISTINCT FROM e.account_effective_date THEN 'account_effective_date' END,
          CASE WHEN db.policy_effective_date::timestamp IS DISTINCT FROM e.policy_effective_date THEN 'policy_effective_date' END,
          CASE WHEN NULLIF(db.pn_no::text,'') IS DISTINCT FROM NULLIF(e.pn_no::text,'') THEN 'pn_no' END,
          CASE WHEN NULLIF(db.sn_no::text,'') IS DISTINCT FROM NULLIF(e.sn_no::text,'') THEN 'sn_no' END,
          CASE WHEN NULLIF(db.plan_type::text,'') IS DISTINCT FROM NULLIF(e.plan_type::text,'') THEN 'plan_type' END,
          CASE WHEN NULLIF(db.reinsurance_type::text,'') IS DISTINCT FROM NULLIF(e.reinsurance_type::text,'') THEN 'reinsurance_type' END,
          CASE WHEN NULLIF(db.period::text,'')::int IS DISTINCT FROM e.period THEN 'period' END,
          CASE WHEN db.issue_date::timestamp IS DISTINCT FROM e.issue_date THEN 'issue_date' END,
          CASE WHEN db.payment_date::timestamp IS DISTINCT FROM e.payment_date THEN 'payment_date' END
        ]::text[], NULL)
      ) = 0
    END) AS is_valid
  FROM excel e
  LEFT JOIN accounting.accounting_${inputdate}_success_events db
    ON NULLIF(db.document_no::text,'') IS NOT DISTINCT FROM NULLIF(e.document_no::text,'')
  AND NULLIF(db.account_code::text,'') IS NOT DISTINCT FROM NULLIF(e.account_code::text,'')
  AND NULLIF(db.plan_type::text,'') IS NOT DISTINCT FROM NULLIF(e.plan_type::text,'')
  AND NULLIF(db.reinsurance_type::text,'') IS NOT DISTINCT FROM NULLIF(e.reinsurance_type::text,'')
  `;
}

function SQL_VALIDATION_SUFFIX_FAIL(inputdate) {
    return `SELECT
    e.document_no,
    e.account_code,
    e.plan_type,
    e.reinsurance_type,
    CASE
      WHEN db.document_no IS NULL THEN ARRAY[
        'event_module','event_module_action','document_no','prefix_document_no','suffix_document_no','description',
        'account_event','account_code','account_type','status','company','channel','product','bau','payment_type',
        'interco','fund','amount','policy_no','policy_group_no','policy_account_no','policy_year',
        'account_effective_date','policy_effective_date','pn_no','sn_no','failure_type','failure_reason',
        'plan_type','reinsurance_type','period','issue_date','payment_date'
      ]::text[]
      ELSE array_remove(ARRAY[
        CASE WHEN NULLIF(db.event_module::text,'') IS DISTINCT FROM NULLIF(e.event_module::text,'') THEN 'event_module' END,
        CASE WHEN NULLIF(db.event_module_action::text,'') IS DISTINCT FROM NULLIF(e.event_module_action::text,'') THEN 'event_module_action' END,
        CASE WHEN NULLIF(db.document_no::text,'') IS DISTINCT FROM NULLIF(e.document_no::text,'') THEN 'document_no' END,
        CASE WHEN NULLIF(db.prefix_document_no::text,'') IS DISTINCT FROM NULLIF(e.prefix_document_no::text,'') THEN 'prefix_document_no' END,
        CASE WHEN NULLIF(db.suffix_document_no::text,'') IS DISTINCT FROM NULLIF(e.suffix_document_no::text,'') THEN 'suffix_document_no' END,
        CASE WHEN NULLIF(db.description::text,'') IS DISTINCT FROM NULLIF(e.description::text,'') THEN 'description' END,
        CASE WHEN NULLIF(db.account_event::text,'') IS DISTINCT FROM NULLIF(e.account_event::text,'') THEN 'account_event' END,
        CASE WHEN NULLIF(db.account_code::text,'') IS DISTINCT FROM NULLIF(e.account_code::text,'') THEN 'account_code' END,
        CASE WHEN NULLIF(db.account_type::text,'') IS DISTINCT FROM NULLIF(e.account_type::text,'') THEN 'account_type' END,
        CASE WHEN NULLIF(db.status::text,'') IS DISTINCT FROM NULLIF(e.status::text,'') THEN 'status' END,
        CASE WHEN NULLIF(db.company::text,'') IS DISTINCT FROM NULLIF(e.company::text,'') THEN 'company' END,
        CASE WHEN NULLIF(db.channel::text,'') IS DISTINCT FROM NULLIF(e.channel::text,'') THEN 'channel' END,
        CASE WHEN NULLIF(db.product::text,'') IS DISTINCT FROM NULLIF(e.product::text,'') THEN 'product' END,
        CASE WHEN NULLIF(db.bau::text,'') IS DISTINCT FROM NULLIF(e.bau::text,'') THEN 'bau' END,
        CASE WHEN NULLIF(db.payment_type::text,'') IS DISTINCT FROM NULLIF(e.payment_type::text,'') THEN 'payment_type' END,
        CASE WHEN NULLIF(db.interco::text,'') IS DISTINCT FROM NULLIF(e.interco::text,'') THEN 'interco' END,
        CASE WHEN NULLIF(db.fund::text,'') IS DISTINCT FROM NULLIF(e.fund::text,'') THEN 'fund' END,
        CASE WHEN trunc(db.amount::numeric, 2) IS DISTINCT FROM e.amount THEN 'amount' END,
        CASE WHEN NULLIF(db.policy_no::text,'') IS DISTINCT FROM NULLIF(e.policy_no::text,'') THEN 'policy_no' END,
        CASE WHEN NULLIF(db.policy_group_no::text,'') IS DISTINCT FROM NULLIF(e.policy_group_no::text,'') THEN 'policy_group_no' END,
        CASE WHEN NULLIF(db.policy_account_no::text,'') IS DISTINCT FROM NULLIF(e.policy_account_no::text,'') THEN 'policy_account_no' END,
        CASE WHEN NULLIF(db.policy_year::text,'')::int IS DISTINCT FROM e.policy_year THEN 'policy_year' END,
        CASE WHEN db.account_effective_date::timestamp IS DISTINCT FROM e.account_effective_date THEN 'account_effective_date' END,
        CASE WHEN db.policy_effective_date::timestamp IS DISTINCT FROM e.policy_effective_date THEN 'policy_effective_date' END,
        CASE WHEN NULLIF(db.failure_type::text,'') IS DISTINCT FROM NULLIF(e.failure_type::text,'') THEN 'failure_type' END,
        CASE WHEN NULLIF(db.failure_reason::text,'') IS DISTINCT FROM NULLIF(e.failure_reason::text,'') THEN 'failure_reason' END,
        CASE WHEN NULLIF(db.pn_no::text,'') IS DISTINCT FROM NULLIF(e.pn_no::text,'') THEN 'pn_no' END,
        CASE WHEN NULLIF(db.sn_no::text,'') IS DISTINCT FROM NULLIF(e.sn_no::text,'') THEN 'sn_no' END,
        CASE WHEN NULLIF(db.plan_type::text,'') IS DISTINCT FROM NULLIF(e.plan_type::text,'') THEN 'plan_type' END,
        CASE WHEN NULLIF(db.reinsurance_type::text,'') IS DISTINCT FROM NULLIF(e.reinsurance_type::text,'') THEN 'reinsurance_type' END,
        CASE WHEN NULLIF(db.period::text,'')::int IS DISTINCT FROM e.period THEN 'period' END,
        CASE WHEN db.issue_date::timestamp IS DISTINCT FROM e.issue_date THEN 'issue_date' END,
        CASE WHEN db.payment_date::timestamp IS DISTINCT FROM e.payment_date THEN 'payment_date' END
      ]::text[], NULL)
    END AS invalid_columns,
    (CASE
      WHEN db.document_no IS NULL THEN FALSE
      ELSE CARDINALITY(
        array_remove(ARRAY[
          CASE WHEN NULLIF(db.event_module::text,'') IS DISTINCT FROM NULLIF(e.event_module::text,'') THEN 'event_module' END,
          CASE WHEN NULLIF(db.event_module_action::text,'') IS DISTINCT FROM NULLIF(e.event_module_action::text,'') THEN 'event_module_action' END,
          CASE WHEN NULLIF(db.document_no::text,'') IS DISTINCT FROM NULLIF(e.document_no::text,'') THEN 'document_no' END,
          CASE WHEN NULLIF(db.prefix_document_no::text,'') IS DISTINCT FROM NULLIF(e.prefix_document_no::text,'') THEN 'prefix_document_no' END,
          CASE WHEN NULLIF(db.suffix_document_no::text,'') IS DISTINCT FROM NULLIF(e.suffix_document_no::text,'') THEN 'suffix_document_no' END,
          CASE WHEN NULLIF(db.description::text,'') IS DISTINCT FROM NULLIF(e.description::text,'') THEN 'description' END,
          CASE WHEN NULLIF(db.account_event::text,'') IS DISTINCT FROM NULLIF(e.account_event::text,'') THEN 'account_event' END,
          CASE WHEN NULLIF(db.account_code::text,'') IS DISTINCT FROM NULLIF(e.account_code::text,'') THEN 'account_code' END,
          CASE WHEN NULLIF(db.account_type::text,'') IS DISTINCT FROM NULLIF(e.account_type::text,'') THEN 'account_type' END,
          CASE WHEN NULLIF(db.status::text,'') IS DISTINCT FROM NULLIF(e.status::text,'') THEN 'status' END,
          CASE WHEN NULLIF(db.company::text,'') IS DISTINCT FROM NULLIF(e.company::text,'') THEN 'company' END,
          CASE WHEN NULLIF(db.channel::text,'') IS DISTINCT FROM NULLIF(e.channel::text,'') THEN 'channel' END,
          CASE WHEN NULLIF(db.product::text,'') IS DISTINCT FROM NULLIF(e.product::text,'') THEN 'product' END,
          CASE WHEN NULLIF(db.bau::text,'') IS DISTINCT FROM NULLIF(e.bau::text,'') THEN 'bau' END,
          CASE WHEN NULLIF(db.payment_type::text,'') IS DISTINCT FROM NULLIF(e.payment_type::text,'') THEN 'payment_type' END,
          CASE WHEN NULLIF(db.interco::text,'') IS DISTINCT FROM NULLIF(e.interco::text,'') THEN 'interco' END,
          CASE WHEN NULLIF(db.fund::text,'') IS DISTINCT FROM NULLIF(e.fund::text,'') THEN 'fund' END,
          CASE WHEN trunc(db.amount::numeric, 2) IS DISTINCT FROM e.amount THEN 'amount' END,
          CASE WHEN NULLIF(db.policy_no::text,'') IS DISTINCT FROM NULLIF(e.policy_no::text,'') THEN 'policy_no' END,
          CASE WHEN NULLIF(db.policy_group_no::text,'') IS DISTINCT FROM NULLIF(e.policy_group_no::text,'') THEN 'policy_group_no' END,
          CASE WHEN NULLIF(db.policy_account_no::text,'') IS DISTINCT FROM NULLIF(e.policy_account_no::text,'') THEN 'policy_account_no' END,
          CASE WHEN NULLIF(db.policy_year::text,'')::int IS DISTINCT FROM e.policy_year THEN 'policy_year' END,
          CASE WHEN db.account_effective_date::timestamp IS DISTINCT FROM e.account_effective_date THEN 'account_effective_date' END,
          CASE WHEN db.policy_effective_date::timestamp IS DISTINCT FROM e.policy_effective_date THEN 'policy_effective_date' END,
          CASE WHEN NULLIF(db.pn_no::text,'') IS DISTINCT FROM NULLIF(e.pn_no::text,'') THEN 'pn_no' END,
          CASE WHEN NULLIF(db.sn_no::text,'') IS DISTINCT FROM NULLIF(e.sn_no::text,'') THEN 'sn_no' END,
          CASE WHEN NULLIF(db.plan_type::text,'') IS DISTINCT FROM NULLIF(e.plan_type::text,'') THEN 'plan_type' END,
          CASE WHEN NULLIF(db.failure_type::text,'') IS DISTINCT FROM NULLIF(e.failure_type::text,'') THEN 'failure_type' END,
          CASE WHEN NULLIF(db.failure_reason::text,'') IS DISTINCT FROM NULLIF(e.failure_reason::text,'') THEN 'failure_reason' END,
          CASE WHEN NULLIF(db.reinsurance_type::text,'') IS DISTINCT FROM NULLIF(e.reinsurance_type::text,'') THEN 'reinsurance_type' END,
          CASE WHEN NULLIF(db.period::text,'')::int IS DISTINCT FROM e.period THEN 'period' END,
          CASE WHEN db.issue_date::timestamp IS DISTINCT FROM e.issue_date THEN 'issue_date' END,
          CASE WHEN db.payment_date::timestamp IS DISTINCT FROM e.payment_date THEN 'payment_date' END
        ]::text[], NULL)
      ) = 0
    END) AS is_valid
  FROM excel e
  LEFT JOIN accounting.accounting_${inputdate}_fail_events db
    ON NULLIF(db.document_no::text,'') IS NOT DISTINCT FROM NULLIF(e.document_no::text,'')
  AND NULLIF(db.account_code::text,'') IS NOT DISTINCT FROM NULLIF(e.account_code::text,'')
  AND NULLIF(db.plan_type::text,'') IS NOT DISTINCT FROM NULLIF(e.plan_type::text,'')
  AND NULLIF(db.reinsurance_type::text,'') IS NOT DISTINCT FROM NULLIF(e.reinsurance_type::text,'')
  `;
}

function sqlEscapeText(s) {
  return String(s).replace(/'/g, "''");
}

function isNullish(v) {
  return (
    v === null || v === undefined || (typeof v === "number" && Number.isNaN(v))
  );
}

/**
 * Build keys/patterns from ALL rows (after override):
 * - key:    `${sn_no}-${pn_no}`
 * - like:   `${sn_no}-${pn_no}%`
 * Rules:
 * - only include rows that have BOTH sn_no and pn_no (no guessing)
 * - de-duplicate
 */
function buildSnPnLists(rows) {
  const keysSet = new Set();
  const likesSet = new Set();

  for (const r of rows) {
    const sn = r?.sn_no;
    const pn = r?.pn_no;

    if (isNullish(sn) || isNullish(pn)) continue; // ไม่เดา ไม่เติม

    const key = `${sn}-${pn}`;
    keysSet.add(key);
    likesSet.add(`${key}%`);
  }

  const keys = Array.from(keysSet);
  const likes = Array.from(likesSet);

  return { keys, likes };
}

function ensureTemplateFiles() {
  if (!fs.existsSync(TEMPLATE_DIR))
    fs.mkdirSync(TEMPLATE_DIR, { recursive: true });

  // CSV header only
  fs.writeFileSync(TEMPLATE_CSV, ALL_HEADERS_TEMPLATE.join(",") + "\n", "utf8");

  // XLSX header only
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("template");
  ws.addRow(ALL_HEADERS_TEMPLATE);

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).alignment = {
    vertical: "middle",
    horizontal: "center",
    wrapText: true,
  };
  ws.views = [{ state: "frozen", ySplit: 1 }];

  // set column widths (simple)
  ws.columns = ALL_HEADERS_TEMPLATE.map((h) => ({
    header: h,
    key: h,
    width: Math.min(35, Math.max(12, h.length + 4)),
  }));

  return wb.xlsx.writeFile(TEMPLATE_XLSX);
}

// call once on startup
ensureTemplateFiles().catch(console.error);

/** Utility: SQL escaping for text */
function sqlEscapeText(s) {
  return String(s).replace(/'/g, "''");
}

/** Utility: detect "empty cell" for extraction */
function isEmptyCell(v) {
  return (
    v === undefined || v === null || (typeof v === "number" && Number.isNaN(v))
  );
}

/** Convert Excel Date or other to a display value (Phase 1) */
function toDisplayValue(v) {
  if (isEmptyCell(v)) return null;
  if (v instanceof Date) return v.toISOString();
  return v;
}

/**
 * CSV parsing: by default, blank fields become '' in parsers.
 * According to your note: CSV blank usually treated as NULL.
 * => convert '' -> null (no guessing of '' unless explicitly supported by file/rule)
 */
function csvCellToValue(v) {
  if (v === "") return null;
  return v;
}

/** Read XLSX buffer to rows (array of arrays) */
function readXlsxToMatrix(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true, raw: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  // header: 1 => arrays; defval: undefined keeps missing cells undefined
  const matrix = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    raw: true,
    defval: undefined,
  });
  return matrix;
}

/** Read CSV buffer to matrix */
function readCsvToMatrix(buffer) {
  const text = buffer.toString("utf8");
  const records = parse(text, {
    relax_quotes: true,
    relax_column_count: true,
    skip_empty_lines: false,
  });
  return records;
}

/**
 * Phase 1: Extract rows based on mapping C->AI
 * - excel_row starts at 2 for first data row
 * - filters rows where all C->AI empty
 * - DOES NOT generate SQL
 * - outputs extract table columns C->AI only (but keeps table_selector internally for phase 2)
 */
function extractFromMatrix(matrix, sourceKind) {
  // matrix[0] is header row (row 1)
  const out = [];
  for (let r = 1; r < matrix.length; r++) {
    const row = matrix[r] || [];

    // indexes: A=0, B=1, C=2 ... AI=34
    const table_selector_raw = row[1]; // B
    const cToAiRaw = row.slice(0, 33); // C..AI inclusive (2..34)

    // normalize per source for "empty"
    const cToAi = cToAiRaw.map((v) => {
      if (sourceKind === "csv") return csvCellToValue(v);
      return isEmptyCell(v) ? null : v;
    });

    // filter: if all C->AI empty => skip
    const allEmpty = cToAi.every(
      (v) =>
        v === null ||
        v === undefined ||
        (typeof v === "number" && Number.isNaN(v))
    );
    if (allEmpty) continue;

    const excel_row = r + 1; // because r=1 means excel row 2

    const obj = {
      excel_row,
      table_selector:
        sourceKind === "csv"
          ? csvCellToValue(table_selector_raw)
          : isEmptyCell(table_selector_raw)
          ? null
          : table_selector_raw,
    };

    // map C->AI into object
    for (let i = 0; i < C_TO_AI.length; i++) {
      obj[C_TO_AI[i]] = cToAi[i] === undefined ? null : cToAi[i];
    }

    out.push(obj);
  }

  // For UI: also provide display version (dates -> ISO string)
  const displayRows = out.map((r) => {
    const d = { excel_row: r.excel_row };
    for (const k of C_TO_AI) d[k] = toDisplayValue(r[k]);
    return d;
  });

  return { rows: out, displayRows };
}

/**
 * Phase 2: Generate SQL only if confirmText matches exactly.
 * SQL format:
 * WITH excel AS (
 *   SELECT v.excel_row, v.table_selector, v.event_module, ..., v.payment_date
 *   FROM (VALUES (...), (...)) AS v(excel_row, table_selector, event_module, ..., payment_date)
 * )
 */

function generateSql(rows, confirmText, inputdate, mode) {
  // if (confirmText !== "Extract ตรง Excel 100%") {
  //   return {
  //     ok: false,
  //     error:
  //       'ต้องยืนยันคำว่า "Extract ตรง Excel 100%" ก่อน จึงจะ generate SQL ได้',
  //   };
  // }

  // inputdate must be exactly 8 digits like 20260615 (no guessing)
  if (!/^\d{8}$/.test(String(inputdate || ""))) {
    return {
      ok: false,
      error: "inputdate ต้องเป็นตัวเลข 8 หลัก เช่น 20260615",
    };
  }

  if (mode !== "success" && mode !== "fail") {
    return { ok: false, error: 'mode ต้องเป็น "success" หรือ "fail"' };
  }

  const selectedSuffix =
    mode === "success" ? SQL_VALIDATION_SUFFIX_SUCCESS(inputdate) : SQL_VALIDATION_SUFFIX_FAIL(inputdate);

  // datatype casts (ตาม mapping เดิมของแอป)
  const typeMap = {
    event_module: "::text",
    event_module_action: "::text",
    document_no: "::text",
    prefix_document_no: "::text",
    suffix_document_no: "::text",
    description: "::text",
    account_event: "::text",
    account_code: "::text",
    account_type: "::text",
    status: "::text",
    company: "::text",
    channel: "::text",
    product: "::text",
    bau: "::text",
    payment_type: "::text",
    interco: "::text",
    fund: "::text",
    amount: "::numeric(18,2)",
    policy_no: "::text",
    policy_group_no: "::text",
    policy_account_no: "::text",
    policy_year: "::int",
    account_effective_date: "::timestamp",
    policy_effective_date: "::timestamp",
    pn_no: "::text",
    sn_no: "::text",
    failure_type: "::text",
    failure_reason: "::text",
    plan_type: "::text",
    reinsurance_type: "::text",
    period: "::int",
    issue_date: "::timestamp",
    payment_date: "::timestamp",
  };

  const colList = [...C_TO_AI];

  function literalWithCast(colName, value) {
    const cast = typeMap[colName] || "::text";

    if (isNullish(value)) return `NULL${cast}`;

    // ห้าม normalize/trim/upper/lower
    if (cast === "::int") {
      if (typeof value === "number" && Number.isFinite(value))
        return `${Math.trunc(value)}${cast}`;
      return `'${sqlEscapeText(value)}'${cast}`;
    }

    if (cast === "::numeric(18,2)") {
      // ไม่ normalize input ให้เป็น 2 decimals ที่ JS; ให้ DB cast เอง
      if (typeof value === "number" && Number.isFinite(value))
        return `'${String(value)}'${cast}`;
      return `'${sqlEscapeText(value)}'${cast}`;
    }

    if (cast === "::timestamp") {
      if (value instanceof Date) return `'${value.toISOString()}'${cast}`;
      return `'${sqlEscapeText(value)}'${cast}`;
    }

    return `'${sqlEscapeText(value)}'${cast}`;
  }

  const valuesLines = rows.map((r) => {
    const parts = [];
    for (const c of C_TO_AI) {
      parts.push(literalWithCast(c, r[c]));
    }
    return `    (${parts.join(", ")})`;
  });

  const sqlCte = `WITH excel AS (
  SELECT
    v.${C_TO_AI.join(",\n    v.")}
  FROM (VALUES
${valuesLines.join(",\n")}
  ) AS v(${colList.join(", ")})
)
`;

  // SQL หลัก (เดิม): CTE + validation suffix
  const sql_main = sqlCte + " " + selectedSuffix;

  // ---- เพิ่ม 3 SQL แยก textarea ----
  const { keys, likes } = buildSnPnLists(rows);

  // เตรียม literal list (ไม่เดา ไม่เติม; ถ้าไม่มี key ก็จะเป็น empty array)
  const keysList = keys.map((k) => `'${sqlEscapeText(k)}'`).join(", ");
  const likesList = likes.map((p) => `'${sqlEscapeText(p)}'`).join(", ");

  const sql_running_no = `DELETE
FROM accounting.gisx_accounting_running_no x
WHERE x.prefix IN (${keysList || "NULL"});`;

  const sql_success_events = `DELETE
FROM accounting.accounting_${inputdate}_success_events x
WHERE x.document_no LIKE ANY (ARRAY[${likesList || "NULL"}]);`;

  const sql_fail_events = `DELETE
FROM accounting.accounting_${inputdate}_fail_events x
WHERE x.document_no LIKE ANY (ARRAY[${likesList || "NULL"}]);`;

  return {
    ok: true,
    sql_main,
    sql_running_no,
    sql_success_events,
    sql_fail_events,
  };
}

/** Template download endpoints (MUST DO EVERY TIME) */
app.get("/api/template/xlsx", (req, res) => res.download(TEMPLATE_XLSX));
app.get("/api/template/csv", (req, res) => res.download(TEMPLATE_CSV));

/** Phase 1 endpoint: upload -> extract */
app.post("/api/extract", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: "no file" });

  const ext = path.extname(req.file.originalname).toLowerCase();
  try {
    let matrix, kind;
    if (ext === ".xlsx" || ext === ".xls") {
      kind = "xlsx";
      matrix = readXlsxToMatrix(req.file.buffer);
    } else if (ext === ".csv") {
      kind = "csv";
      matrix = readCsvToMatrix(req.file.buffer);
    } else {
      return res
        .status(400)
        .json({ ok: false, error: "รองรับเฉพาะ .xlsx/.xls/.csv" });
    }

    const { rows, displayRows } = extractFromMatrix(matrix, kind);
    return res.json({
      ok: true,
      kind,
      // rows: internal (may include Date objects) - send displayRows for UI; keep rows in client for SQL generation
      rows, // client will send back for phase 2 (source of truth = uploaded file + user overrides in UI)
      displayRows, // safe for table rendering
      columns: C_TO_AI,
    });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e?.message || e) });
  }
});

/** Phase 2 endpoint: generate SQL only after confirmation */
app.post("/api/generate-sql", (req, res) => {
  const { rows, confirmText, inputdate, mode } = req.body || {};
  if (!Array.isArray(rows)) return res.status(400).json({ ok: false, error: "rows ต้องเป็น array" });

  const result = generateSql(rows, confirmText, inputdate, mode);
  return res.json(result);
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running: http://localhost:${port}`);
  console.log(`Template XLSX: http://localhost:${port}/api/template/xlsx`);
  console.log(`Template CSV : http://localhost:${port}/api/template/csv`);
});
