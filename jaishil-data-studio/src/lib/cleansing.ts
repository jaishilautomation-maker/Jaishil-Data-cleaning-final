import * as XLSX from "xlsx";

export type Profile = "stock" | "debtors" | "sales" | "purchase" | "jv" | "ho_exp" | "creditors";

export interface ChangeLogEntry {
  rowNumber: number;
  reason: string;
  data: Record<string, unknown>;
}

export interface CleanResult {
  cleanedRows: Record<string, unknown>[];
  removed: ChangeLogEntry[];
  headers: string[];
  originalCount: number;
  blob?: Blob;
}

const EXCLUDED_BRANCHES = ["HEAD OFFICE", "HO", "TEST", "DUMMY", "INTERNAL"];

const STOCK_QTY_KEYS = [
  "closing stock",
  "closing qty",
  "closing quantity",
  "stock",
  "qty",
  "quantity",
  "balance",
  "closing balance",
];

const BRANCH_KEYS = ["branch", "branch name", "location", "godown"];

function norm(s: string) {
  return s.toString().trim().toLowerCase();
}

function findKey(headers: string[], candidates: string[]): string | null {
  const map = new Map(headers.map((h) => [norm(h), h]));
  for (const c of candidates) {
    if (map.has(c)) return map.get(c)!;
  }
  // partial match
  for (const h of headers) {
    const n = norm(h);
    if (candidates.some((c) => n.includes(c))) return h;
  }
  return null;
}

function toNumber(v: unknown): number | null {
  if (v === null || v === undefined || v === "") return null;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}

export async function cleanFile(file: File, profile: Profile): Promise<CleanResult> {
  const formData = new FormData();
  formData.append("file", file);
  formData.append("profile", profile);

  // Call the Flask API
  const apiUrl = import.meta.env.VITE_API_URL || "http://localhost:5000";
  const response = await fetch(`${apiUrl}/clean`, {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "Unknown error");
    throw new Error(`Server Error: ${errorText}`);
  }

  const blob = await response.blob();
  const buf = await blob.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  
  // Get active sheet (which has the cleaned data)
  const dataSheetName = wb.SheetNames[0];
  const dataSheet = wb.Sheets[dataSheetName];
  const cleanedRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(dataSheet, { defval: "" });
  const headers = cleanedRows.length ? Object.keys(cleanedRows[0]) : [];

  // Get change log if it exists
  const removed: ChangeLogEntry[] = [];
  const logSheetName = wb.SheetNames.find(
    (name) => name === "Change Log" || name === "Changes_Log"
  );
  if (logSheetName) {
    const logSheet = wb.Sheets[logSheetName];
    const logRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(logSheet, { defval: "" });
    logRows.forEach((row, i) => {
      const rowNum = Number(
        row["Original Row Number"] ?? 
        row["Original Row #"] ?? 
        row["Row Number"] ?? 
        row["Row"] ?? 
        (i + 1)
      );
      const reason = String(
        row["Reason for Removal"] ?? 
        row["Reason"] ?? 
        "Removed"
      );
      
      const sheetName = row["Sheet Name"] ? `[${row["Sheet Name"]}] ` : "";
      const reasonWithSheet = sheetName + reason;

      // Store the rest of the columns as data
      const data: Record<string, unknown> = {};
      Object.keys(row).forEach((key) => {
        if (
          key !== "Original Row Number" && 
          key !== "Original Row #" && 
          key !== "Row Number" && 
          key !== "Row" && 
          key !== "Reason for Removal" && 
          key !== "Reason" &&
          key !== "Sheet Name"
        ) {
          data[key] = row[key];
        }
      });

      removed.push({
        rowNumber: rowNum,
        reason: reasonWithSheet,
        data: data,
      });
    });
  }

  // Count total cleaned rows across all worksheets (except Change Log / Changes_Log)
  let totalCleanedCount = 0;
  wb.SheetNames.forEach((sheetName) => {
    if (sheetName !== "Change Log" && sheetName !== "Changes_Log") {
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet);
      totalCleanedCount += rows.length;
    }
  });

  const originalCount = totalCleanedCount + removed.length;

  return {
    cleanedRows,
    removed,
    headers,
    originalCount,
    blob,
  };
}

export function downloadCleanedFile(result: CleanResult, originalName: string) {
  if (result.blob) {
    const url = URL.createObjectURL(result.blob);
    const a = document.createElement("a");
    a.href = url;
    const base = originalName.replace(/\.(xlsx|xls)$/i, "");
    a.download = `${base}_Cleaned.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    return;
  }

  // Fallback client-side generation
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(result.cleanedRows, { header: result.headers });
  XLSX.utils.book_append_sheet(wb, ws, "Cleaned Data");

  if (result.removed.length) {
    const logRows = result.removed.map((r) => ({
      "Original Row Number": r.rowNumber,
      "Reason for Removal": r.reason,
      ...r.data,
    }));
    const logWs = XLSX.utils.json_to_sheet(logRows);
    XLSX.utils.book_append_sheet(wb, logWs, "Change Log");
  }

  const base = originalName.replace(/\.(xlsx|xls)$/i, "");
  XLSX.writeFile(wb, `${base}_Cleaned.xlsx`);
}

export const PROFILE_LABELS: Record<Profile, { title: string; desc: string }> = {
  stock: { title: "Stock Summary", desc: "Stock inventory cleaning — removes header and subtotal rows dynamically" },
  debtors: { title: "Debtors Report", desc: "Branch exclusion cleaning — removes internal/HO branches" },
  sales: { title: "Sales Register", desc: "Sales cleaning — removes non-products and corrects typos" },
  purchase: { title: "Purchase Register", desc: "Purchase cleaning — converts ton to kg and corrects suppliers" },
  jv: { title: "Journal Vouchers", desc: "JV cleaning — filters by Jaishil and excludes standard ledger types" },
  ho_exp: { title: "HO Expenses", desc: "Head Office expenses — checks for grand total mismatches" },
  creditors: { title: "Sundry Creditors", desc: "Sundry creditors — removes company name from columns" },
};
