import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import toast, { Toaster } from "react-hot-toast";
import "./App.css";

import { auth, db } from "./firebase";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  signOut,
  type User,
} from "firebase/auth";
import {
  doc,
  getDoc,
  serverTimestamp,
  setDoc,
  deleteDoc,
} from "firebase/firestore";

type UiRow = {
  coil: string;
  unit: "MT";
  totalAvailableStockMt: number;
  blockStockMt: number;
  freeStockMt: number;
  tentativeShipmentDate: string;
};

type TempTableDoc = {
  asAtDate: string; // YYYY-MM-DD
  sourceFileName: string;
  rows: UiRow[];
  updatedAt?: unknown;
};

const EXACT_ALLOWED = new Set(["colorbond", "lgalvanized", "zincalume"]);

function norm(s: unknown) {
  return String(s ?? "")
    .trim()
    .toLowerCase();
}
function num(v: unknown): number {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string") {
    const n = Number(v.replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}
function round3(n: number) {
  return Math.round(n * 1000) / 1000;
}

function recomputeRow(row: UiRow): UiRow {
  const free = Math.max(row.totalAvailableStockMt - row.blockStockMt, 0);
  return {
    ...row,
    freeStockMt: round3(free),
    tentativeShipmentDate: row.tentativeShipmentDate ?? "",
  };
}

// Fixed doc id = "current"
const TABLE_DOC_REF = doc(db, "temp_tables", "current");

export default function App() {
  const [loading, setLoading] = useState(false);
  const [loadingText, setLoadingText] = useState("");
  type BusyAction = "load" | "login" | "save" | "delete" | "export" | null;
  const [busy, setBusy] = useState<BusyAction>(null);

  const [user, setUser] = useState<User | null>(null);

  // simple login inputs
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  const [asAtDate, setAsAtDate] = useState(() => {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  });

  const [rows, setRows] = useState<UiRow[]>([]);
  const [sourceFileName, setSourceFileName] = useState("");
  const [status, setStatus] = useState<string>("");

  // Auth state
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return () => unsub();
  }, []);

  // Load existing temp table after login
  useEffect(() => {
    if (!user) return;

    run("load", "Loading saved temp table...", async () => {
      const snap = await getDoc(TABLE_DOC_REF);
      if (!snap.exists()) {
        setStatus("No saved temp table.");
        toast("No saved temp table.", { icon: "ℹ️" });
        return;
      }
      const data = snap.data() as TempTableDoc;

      setAsAtDate(data.asAtDate || asAtDate);
      setSourceFileName(data.sourceFileName || "");
      setRows(Array.isArray(data.rows) ? data.rows : []);
      setStatus("Loaded saved temp table.");
      toast.success("Loaded saved temp table.");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    }).catch((e: any) => {
      console.error(e);
      setStatus("Failed to load temp table.");
      toast.error(e?.message || "Failed to load");
    });

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [user]);

  async function run(
    action: BusyAction,
    text: string,
    fn: () => Promise<void>
  ) {
    try {
      setBusy(action);
      setLoading(true);
      setLoadingText(text);
      await fn();
    } finally {
      setLoading(false);
      setLoadingText("");
      setBusy(null);
    }
  }

  const totals = useMemo(() => {
    const totalAvailable = rows.reduce(
      (s, r) => s + r.totalAvailableStockMt,
      0
    );
    const totalFree = rows.reduce((s, r) => s + r.freeStockMt, 0);
    const totalBlocked = rows.reduce((s, r) => s + r.blockStockMt, 0);
    return {
      totalAvailable: round3(totalAvailable),
      totalFree: round3(totalFree),
      totalBlocked: round3(totalBlocked),
      count: rows.length,
    };
  }, [rows]);

  async function login() {
    await run("login", "Signing in...", async () => {
      await signInWithEmailAndPassword(auth, email.trim(), password);
      setStatus("Signed in.");
      toast.success("Signed in.");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    }).catch((e: any) => {
      console.error(e);
      setStatus("Sign in failed.");
      toast.error(e?.message || "Sign in failed");
    });
  }

  async function logout() {
    await signOut(auth);
    setRows([]);
    setSourceFileName("");
    setStatus("Signed out.");
  }

  async function exportToExcel() {
    if (rows.length === 0) {
      setStatus("Nothing to export.");
      toast.error("Nothing to export.");
      return;
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Coils Stock", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    ws.addRow([`Coils stock as at: ${asAtDate}`]);
    ws.mergeCells("A1:F1");
    ws.getCell("A1").font = { bold: true, size: 14 };
    ws.getCell("A1").alignment = { vertical: "middle", horizontal: "left" };
    ws.getRow(1).height = 22;

    ws.addRow([]);

    const headerRowIndex = 3;
    ws.addRow([
      "Coil",
      "Unit",
      "Total available stock (MT)",
      "Block stock (MT)",
      "Free stock (MT)",
      "Tentative shipment date",
    ]);

    const headerRow = ws.getRow(headerRowIndex);
    headerRow.height = 20;
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF1F4E79" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFBFBFBF" } },
        left: { style: "thin", color: { argb: "FFBFBFBF" } },
        bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
        right: { style: "thin", color: { argb: "FFBFBFBF" } },
      };
    });

    rows.forEach((r) => {
      ws.addRow([
        r.coil,
        r.unit,
        r.totalAvailableStockMt,
        r.blockStockMt,
        r.freeStockMt,
        r.tentativeShipmentDate || "",
      ]);
    });

    ws.columns = [
      { key: "coil", width: 45 },
      { key: "unit", width: 10 },
      { key: "total", width: 24 },
      { key: "block", width: 18 },
      { key: "free", width: 18 },
      { key: "ship", width: 26 },
    ];

    const firstDataRow = headerRowIndex + 1;
    const lastDataRow = ws.rowCount;

    for (let r = firstDataRow; r <= lastDataRow; r++) {
      const row = ws.getRow(r);
      row.height = 18;

      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE0E0E0" } },
          left: { style: "thin", color: { argb: "FFE0E0E0" } },
          bottom: { style: "thin", color: { argb: "FFE0E0E0" } },
          right: { style: "thin", color: { argb: "FFE0E0E0" } },
        };
        cell.alignment = { vertical: "middle", wrapText: true };
      });

      ws.getCell(`C${r}`).numFmt = "0.000";
      ws.getCell(`D${r}`).numFmt = "0.000";
      ws.getCell(`E${r}`).numFmt = "0.000";

      ws.getCell(`A${r}`).alignment = {
        vertical: "middle",
        horizontal: "left",
        wrapText: true,
      };
      ws.getCell(`B${r}`).alignment = {
        vertical: "middle",
        horizontal: "center",
      };
      ws.getCell(`C${r}`).alignment = {
        vertical: "middle",
        horizontal: "right",
      };
      ws.getCell(`D${r}`).alignment = {
        vertical: "middle",
        horizontal: "right",
      };
      ws.getCell(`E${r}`).alignment = {
        vertical: "middle",
        horizontal: "right",
      };
      ws.getCell(`F${r}`).alignment = {
        vertical: "middle",
        horizontal: "left",
      };
    }

    ws.autoFilter = {
      from: { row: headerRowIndex, column: 1 },
      to: { row: headerRowIndex, column: 6 },
    };

    ws.pageSetup.orientation = "landscape";
    ws.pageSetup.fitToPage = true;
    ws.pageSetup.fitToWidth = 1;
    ws.pageSetup.fitToHeight = 0;

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const safeDate = asAtDate.replaceAll("-", "");
    saveAs(blob, `Coils_Stock_${safeDate}.xlsx`);
    setStatus("Exported Excel file.");
    toast.success("Exported Excel file.");
  }

  async function onFileChange(file?: File | null) {
    if (!file) return;
    setSourceFileName(file.name);
    setStatus("Parsing Excel...");

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];

    const data = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
      defval: "",
    });

    const filtered = data
      .filter((r) => {
        const cat = norm(r["Item Category Code"]);
        return EXACT_ALLOWED.has(cat);
      })
      .map((r) => {
        const coil = String(r["Item Description"] ?? "").trim();
        const totalKg = num(r["Total Quantity"]);
        const totalMt = totalKg / 1000;

        const totalAvailableStockMt = round3(totalMt);
        const blockStockMt = 0;
        const freeStockMt = round3(
          Math.max(totalAvailableStockMt - blockStockMt, 0)
        );

        return {
          coil,
          unit: "MT" as const,
          totalAvailableStockMt,
          blockStockMt,
          freeStockMt,
          tentativeShipmentDate: "",
        };
      })
      .filter((r) => r.coil.length > 0);

    setRows(filtered);
    setStatus(`Parsed. Matched rows: ${filtered.length}`);
    toast.success(`Parsed. Matched rows: ${filtered.length}`);
  }

  function updateRowLocal(index: number, patch: Partial<UiRow>) {
    setRows((prev) => {
      const next = [...prev];
      const old = next[index];
      const merged: UiRow = recomputeRow({ ...old, ...patch } as UiRow);
      next[index] = merged;
      return next;
    });
  }

  async function saveToFirestore() {
    if (!user) return toast.error("Please sign in first.");
    if (rows.length === 0) return toast.error("No rows to save.");

    await run("save", "Saving to Firestore...", async () => {
      const cleanedRows = rows.map((r) => recomputeRow(r));

      const payload: TempTableDoc = {
        asAtDate,
        sourceFileName: sourceFileName ?? "",
        rows: cleanedRows,
        updatedAt: serverTimestamp(),
      };

      await setDoc(TABLE_DOC_REF, payload, { merge: false });

      setRows(cleanedRows);
      setStatus("Saved.");
      toast.success("Saved.");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    }).catch((e: any) => {
      console.error(e);
      setStatus("Save failed.");
      toast.error(e?.message || "Save failed");
    });
  }

  async function deleteTemp() {
    if (!user) return toast.error("Please sign in first.");

    await run("delete", "Deleting temp table...", async () => {
      await deleteDoc(TABLE_DOC_REF);
      setRows([]);
      setSourceFileName("");
      setStatus("Deleted.");
      toast.success("Deleted.");
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    }).catch((e: any) => {
      console.error(e);
      setStatus("Delete failed.");
      toast.error(e?.message || "Delete failed");
    });
  }

  // --- UI ---
  if (!user) {
    return (
      <div className="container" style={{ maxWidth: 520 }}>
        <Toaster />
        <div className="card" style={{ marginTop: 32 }}>
          <div className="cardHeader">
            <h2>Sign in</h2>
          </div>
          <div className="cardBody">
            <div className="controls" style={{ gridTemplateColumns: "1fr" }}>
              <div className="field">
                <div className="label">Email</div>
                <input
                  className="input"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="you@example.com"
                />
              </div>
              <div className="field">
                <div className="label">Password</div>
                <input
                  className="input"
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                />
              </div>
              <button className="btn btnPrimary" onClick={login}>
                Sign in
              </button>
              <div className="status">{status}</div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="container">
      <Toaster />

      {loading && (
        <div className="loadingOverlay" role="status" aria-live="polite">
          <div className="loadingCard">
            <div className="spinner" />
            <div className="loadingText">{loadingText || "Working..."}</div>
          </div>
        </div>
      )}

      <div className="topbar">
        <div className="brand">
          <h1>Coils Stock</h1>
          <p>Filter → Convert (kg → MT) → Edit → Save (Firestore) → Export</p>
        </div>

        <div className="pills">
          <div className="pill">
            <b>Rows:</b> {totals.count}
          </div>
          <div className="pill">
            <b>Total Available (MT):</b> {totals.totalAvailable}
          </div>
          <div className="pill">
            <b>Total Blocked (MT):</b> {totals.totalBlocked}
          </div>
          <div className="pill">
            <b>Total Free (MT):</b> {totals.totalFree}
          </div>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 14 }}>
        <div className="cardHeader">
          <h2>Controls</h2>
          <div className="actions">
            <button
              className="btn"
              onClick={exportToExcel}
              disabled={rows.length === 0}
            >
              Export to Excel
            </button>
            <button
              className="btn btnPrimary"
              onClick={saveToFirestore}
              disabled={loading || rows.length === 0}
            >
              {busy === "save" ? "Saving..." : "Save to Firestore (temp)"}
            </button>
            <button
              className="btn btnDanger"
              onClick={deleteTemp}
              disabled={loading}
            >
              {busy === "delete" ? "Deleting..." : "Delete temp"}
            </button>
            <button className="btn" onClick={logout} disabled={loading}>
              Sign out
            </button>
          </div>
        </div>

        <div className="cardBody">
          <div className="controls">
            <div className="field" style={{ gridColumn: "span 3" }}>
              <div className="label">Coils stock as at</div>
              <input
                className="input"
                type="date"
                value={asAtDate}
                onChange={(e) => setAsAtDate(e.target.value)}
              />
            </div>

            <div className="field" style={{ gridColumn: "span 5" }}>
              <div className="label">Upload Excel</div>
              <input
                className="fileInput"
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => onFileChange(e.target.files?.[0])}
              />
            </div>

            <div className="field" style={{ gridColumn: "span 4" }}>
              <div className="label">Source file</div>
              <input className="input" value={sourceFileName || "—"} readOnly />
            </div>
          </div>

          <div className="status">{status}</div>
        </div>
      </div>

      <div className="card">
        <div className="cardHeader">
          <h2>Table</h2>
          <div style={{ fontSize: 12, color: "var(--muted)" }}>
            Unit is fixed as <b>MT</b>. Free stock is auto-calculated.
          </div>
        </div>

        <div className="tableWrap">
          <table className="table">
            <thead>
              <tr>
                <th>Coil</th>
                <th className="tCenter">Unit</th>
                <th className="tRight">Total available stock (MT)</th>
                <th className="tRight">Block stock (MT)</th>
                <th className="tRight">Free stock (MT)</th>
                <th>Tentative shipment date</th>
              </tr>
            </thead>

            <tbody>
              {rows.map((r, i) => {
                const invalid = r.blockStockMt > r.totalAvailableStockMt;
                return (
                  <tr key={i} className={invalid ? "warnRow" : undefined}>
                    <td>{r.coil}</td>
                    <td className="tCenter">{r.unit}</td>
                    <td className="tRight">
                      {r.totalAvailableStockMt.toFixed(3)}
                    </td>
                    <td className="tRight">
                      <input
                        className="inlineInput"
                        type="number"
                        step="0.001"
                        value={r.blockStockMt}
                        onChange={(e) =>
                          updateRowLocal(i, {
                            blockStockMt: num(e.target.value),
                          })
                        }
                      />
                    </td>
                    <td className="tRight">{r.freeStockMt.toFixed(3)}</td>
                    <td>
                      <input
                        className="inlineInput inlineInputWide"
                        type="text"
                        value={r.tentativeShipmentDate}
                        onChange={(e) =>
                          updateRowLocal(i, {
                            tentativeShipmentDate: e.target.value,
                          })
                        }
                        placeholder="e.g., 2026-01-20"
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {rows.length > 0 && (
          <div className="cardBody">
            <div className="note">
              Free stock is calculated as <b>Total available - Block stock</b>{" "}
              (clamped at 0). Rows where block stock exceeds total are outlined.
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
