import { useState, useEffect, useMemo, useCallback, Fragment } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Upload, PiggyBank, Receipt, Settings2, Trash2, Plus, X, Check,
  AlertCircle, ChevronRight, ChevronLeft, Search, Download, FileSpreadsheet,
  Pencil, ArrowRight, Users, Stamp, LogOut, Lock, BarChart3, Tags
} from "lucide-react";
import {
  PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
} from "recharts";
import { auth } from "./firebase";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { subscribeDoc, saveDoc } from "./storage";

/* ---------------------------------- helpers ---------------------------------- */

const DEFAULT_CATEGORIES = [
  { name: "Income", color: "#3FA7A0", group: "income" },
  { name: "Housing", color: "#D4A24E", group: "expense" },
  { name: "Groceries", color: "#4F9D69", group: "expense" },
  { name: "Dining Out", color: "#E0654F", group: "expense" },
  { name: "Transportation", color: "#6FA8DC", group: "expense" },
  { name: "Utilities", color: "#B07CD0", group: "expense" },
  { name: "Subscriptions", color: "#F2B134", group: "expense" },
  { name: "Health", color: "#5DD39E", group: "expense" },
  { name: "Shopping", color: "#EF8354", group: "expense" },
  { name: "Entertainment", color: "#9D8DF1", group: "expense" },
  { name: "Pets", color: "#7FB3D5", group: "expense" },
  { name: "Travel", color: "#E8A0BF", group: "expense" },
  { name: "Uncategorized", color: "#7A8B99", group: "expense" },
];

const KEYWORD_MAP = {
  Groceries: ["publix", "kroger", "walmart", "aldi", "whole foods", "grocery", "winn-dixie", "trader joe", "sprouts"],
  "Dining Out": ["starbucks", "restaurant", "chipotle", "mcdonald", "doordash", "uber eats", "grubhub", "cafe", "coffee", "pizza", "taco"],
  Transportation: ["shell", "exxon", "chevron", "uber", "lyft", " gas ", "parking", "toll", "wawa"],
  Utilities: ["electric", "duke energy", "comcast", "at&t", "verizon", "utility", "fiber", "internet", "water co", "spectrum"],
  Subscriptions: ["netflix", "spotify", "hulu", "disney+", "amazon prime", "subscription", "apple.com/bill", "icloud"],
  Health: ["pharmacy", "cvs", "walgreens", "clinic", "medical", "dental", "doctor"],
  Shopping: ["amazon", "target", "best buy", "ebay", "etsy"],
  Entertainment: ["movie", "theatre", "theater", "cinema", "ticketmaster"],
  Pets: ["petco", "petsmart", " vet ", "chewy"],
  Travel: ["airline", "hotel", "airbnb", "marriott", "delta", "southwest", "expedia"],
  Housing: ["mortgage", "rent", "hoa", "property mgmt"],
};

const PALETTE = ["#D4A24E", "#4F9D69", "#E0654F", "#6FA8DC", "#B07CD0", "#F2B134", "#5DD39E", "#EF8354", "#9D8DF1", "#7FB3D5", "#E8A0BF", "#7A8B99"];

function guessCategory(desc) {
  const d = ` ${(desc || "").toLowerCase()} `;
  for (const [cat, kws] of Object.entries(KEYWORD_MAP)) {
    if (kws.some((k) => d.includes(k))) return cat;
  }
  return "Uncategorized";
}

function parseAmount(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  let s = String(raw).trim();
  if (s === "") return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.slice(1, -1);
  }
  s = s.replace(/[$,\s]/g, "");
  if (s.startsWith("-")) {
    neg = true;
    s = s.slice(1);
  } else if (s.startsWith("+")) {
    s = s.slice(1);
  }
  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return neg ? -n : n;
}

function parseDateString(raw) {
  if (raw === null || raw === undefined) return null;
  const s = String(raw).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    let [, mo, da, yr] = m;
    if (yr.length === 2) yr = (parseInt(yr, 10) > 50 ? "19" : "20") + yr;
    return `${yr}-${mo.padStart(2, "0")}-${da.padStart(2, "0")}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return null;
}

function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("en-US", { style: "currency", currency: "USD" });
}

function catGroup(cat) {
  return cat && cat.group === "income" ? "income" : "expense";
}

function normCatName(s) {
  return (s || "").trim().toLowerCase();
}

function findCategory(categories, name) {
  const n = normCatName(name);
  return categories.find((c) => normCatName(c.name) === n);
}

function groupOf(categories, name) {
  return catGroup(findCategory(categories, name));
}

function thisMonth() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function fmtMonthLabel(ym) {
  const [y, m] = ym.split("-").map(Number);
  const d = new Date(y, m - 1, 1);
  return d.toLocaleDateString("en-US", { month: "long", year: "numeric" });
}

function makeId() {
  if (window.crypto && window.crypto.randomUUID) return window.crypto.randomUUID();
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function dedupeKey(t) {
  return `${t.date}|${t.amount.toFixed(2)}|${t.description.toLowerCase().trim()}|${(t.source || "").toLowerCase()}`;
}

function guessColumn(headers, candidates) {
  const lower = headers.map((h) => String(h || "").toLowerCase().trim());
  for (const c of candidates) {
    const idx = lower.findIndex((h) => h === c);
    if (idx >= 0) return idx;
  }
  for (const c of candidates) {
    const idx = lower.findIndex((h) => h.includes(c));
    if (idx >= 0) return idx;
  }
  return -1;
}

const DATE_CANDS = ["date", "transaction date", "posted date", "posting date"];
const DESC_CANDS = ["description", "memo", "payee", "name", "transaction", "details"];
const AMOUNT_CANDS = ["amount", "amt"];
const DEBIT_CANDS = ["debit", "withdrawal", "withdrawals"];
const CREDIT_CANDS = ["credit", "deposit", "deposits"];
const CATEGORY_CANDS = ["category", "type"];

/* ---------------------------------- styles ---------------------------------- */


const Styles = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

    .hbl-root {
      --bg: #131D27;
      --surface: #1C2A37;
      --surface-raised: #243646;
      --line: rgba(212,162,78,0.16);
      --border: rgba(255,255,255,0.08);
      --gold: #D4A24E;
      --gold-bright: #E9C57E;
      --ink: #EFEBE1;
      --ink-dim: #9FB0BD;
      --green: #5FB683;
      --coral: #E0735F;
      --display-font: 'Fraunces', Georgia, 'Times New Roman', serif;
      --body-font: 'Inter', system-ui, -apple-system, sans-serif;
      --mono-font: 'IBM Plex Mono', 'SF Mono', Consolas, monospace;

      background: var(--bg);
      color: var(--ink);
      font-family: var(--body-font);
      min-height: 100%;
      width: 100%;
      border-radius: 14px;
      overflow: hidden;
      background-image:
        repeating-linear-gradient(to bottom, transparent 0, transparent 39px, var(--line) 40px);
    }
    .hbl-shell { max-width: 980px; margin: 0 auto; padding: 28px 24px 48px; }
    .hbl-root * { box-sizing: border-box; }

    .hbl-top { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; margin-bottom: 22px; flex-wrap: wrap; }
    .hbl-title { font-family: var(--display-font); font-weight: 600; font-size: 30px; letter-spacing: 0.2px; margin: 0; color: var(--ink); }
    .hbl-sub { color: var(--ink-dim); font-size: 13px; margin-top: 4px; display:flex; align-items:center; gap:6px; }
    .hbl-who { display: flex; align-items: center; gap: 8px; }
    .hbl-who input {
      background: var(--surface-raised); border: 1px solid var(--border); color: var(--ink);
      border-radius: 8px; padding: 7px 10px; font-size: 13px; font-family: var(--body-font); width: 120px;
    }
    .hbl-who-label { font-size: 11px; color: var(--ink-dim); text-transform: uppercase; letter-spacing: 0.08em; }

    .hbl-nav { display: flex; gap: 6px; margin-bottom: 24px; border-bottom: 1px solid var(--border); padding-bottom: 0; }
    .hbl-nav-tab {
      display: flex; align-items: center; gap: 7px; background: transparent; border: none; color: var(--ink-dim);
      font-family: var(--body-font); font-size: 13.5px; font-weight: 500; padding: 10px 14px; cursor: pointer;
      border-bottom: 2px solid transparent; transition: color 0.15s, border-color 0.15s;
    }
    .hbl-nav-tab:hover { color: var(--ink); }
    .hbl-nav-tab.active { color: var(--gold-bright); border-bottom-color: var(--gold); }

    .hbl-card {
      background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px;
    }
    .hbl-card + .hbl-card { margin-top: 16px; }

    .hbl-section-title { font-family: var(--display-font); font-size: 17px; font-weight: 600; margin: 0 0 14px; color: var(--ink); }
    .hbl-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; }

    .hbl-btn {
      display: inline-flex; align-items: center; gap: 6px; font-family: var(--body-font); font-weight: 500; font-size: 13px;
      padding: 8px 14px; border-radius: 8px; border: 1px solid var(--border); background: var(--surface-raised);
      color: var(--ink); cursor: pointer; transition: background 0.15s, border-color 0.15s;
    }
    .hbl-btn:hover { border-color: var(--gold); }
    .hbl-btn-primary { background: var(--gold); color: #1A1306; border-color: var(--gold); font-weight: 600; }
    .hbl-btn-primary:hover { background: var(--gold-bright); }
    .hbl-btn-danger { color: var(--coral); }
    .hbl-btn-danger:hover { border-color: var(--coral); }
    .hbl-btn:disabled { opacity: 0.45; cursor: not-allowed; }
    .hbl-btn-sm { padding: 5px 9px; font-size: 12px; }

    .hbl-input, .hbl-select {
      background: var(--surface-raised); border: 1px solid var(--border); color: var(--ink); border-radius: 8px;
      padding: 8px 10px; font-size: 13px; font-family: var(--body-font); width: 100%;
    }
    .hbl-input:focus, .hbl-select:focus { outline: none; border-color: var(--gold); }
    label.hbl-label { font-size: 11.5px; color: var(--ink-dim); text-transform: uppercase; letter-spacing: 0.06em; display: block; margin-bottom: 5px; }

    .hbl-mono { font-family: var(--mono-font); }

    /* the seal / stamp signature element */
    .hbl-seal {
      width: 58px; height: 58px; border-radius: 50%; border: 2px dashed var(--ink-dim); flex-shrink: 0;
      display: flex; align-items: center; justify-content: center; transform: rotate(-6deg); position: relative;
      font-family: var(--mono-font); font-size: 13px; font-weight: 600; color: var(--ink-dim);
    }
    .hbl-seal::after {
      content: ""; position: absolute; inset: 6px; border: 1px solid currentColor; border-radius: 50%; opacity: 0.45;
    }
    .hbl-seal.over { color: var(--coral); border-color: var(--coral); }
    .hbl-seal.under { color: var(--green); border-color: var(--green); }
    .hbl-seal.empty { color: var(--ink-dim); border-color: var(--ink-dim); }

    .hbl-cat-row { display: flex; align-items: center; gap: 16px; padding: 14px 0; border-bottom: 1px dashed var(--border); }
    .hbl-cat-row:last-child { border-bottom: none; }
    .hbl-cat-main { flex: 1; min-width: 0; }
    .hbl-cat-head { display: flex; align-items: center; gap: 8px; margin-bottom: 7px; }
    .hbl-dot { width: 9px; height: 9px; border-radius: 50%; flex-shrink: 0; }
    .hbl-cat-name { font-weight: 600; font-size: 14px; }
    .hbl-cat-figs { font-family: var(--mono-font); font-size: 12.5px; color: var(--ink-dim); margin-left: auto; white-space: nowrap; }
    .hbl-track { height: 7px; border-radius: 5px; background: var(--surface-raised); overflow: hidden; }
    .hbl-fill { height: 100%; border-radius: 5px; transition: width 0.4s ease; }

    .hbl-totalbar { display: flex; align-items: center; gap: 18px; padding: 4px 0 18px; }
    .hbl-total-num { font-family: var(--display-font); font-size: 34px; font-weight: 600; }
    .hbl-total-label { font-size: 12px; color: var(--ink-dim); text-transform: uppercase; letter-spacing: 0.07em; margin-top: 2px;}

    table.hbl-table { width: 100%; border-collapse: collapse; font-size: 13px; }
    table.hbl-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-dim); font-weight: 500; padding: 8px 10px; border-bottom: 1px solid var(--border); }
    table.hbl-table th.hbl-th-sort { cursor: pointer; user-select: none; white-space: nowrap; }
    table.hbl-table th.hbl-th-sort:hover { color: var(--gold-bright); }
    table.hbl-table td { padding: 9px 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }
    table.hbl-table tr:hover td { background: rgba(255,255,255,0.02); }
    .hbl-amt-out { color: var(--coral); font-family: var(--mono-font); }
    .hbl-amt-in { color: var(--green); font-family: var(--mono-font); }

    .hbl-pill { display: inline-flex; align-items: center; gap: 5px; padding: 3px 9px; border-radius: 999px; font-size: 11.5px; background: var(--surface-raised); border: 1px solid var(--border); color: var(--ink-dim); }

    .hbl-empty { text-align: center; padding: 48px 20px; color: var(--ink-dim); }
    .hbl-empty svg { opacity: 0.5; margin-bottom: 10px; }

    .hbl-dropzone {
      border: 2px dashed var(--border); border-radius: 12px; padding: 36px 20px; text-align: center; cursor: pointer;
      transition: border-color 0.15s, background 0.15s;
    }
    .hbl-dropzone:hover, .hbl-dropzone.drag { border-color: var(--gold); background: rgba(212,162,78,0.05); }

    .hbl-mapgrid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin: 16px 0; }
    .hbl-step-badge {
      display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 50%;
      background: var(--gold); color: #1A1306; font-size: 12px; font-weight: 700; font-family: var(--mono-font); margin-right: 8px;
    }

    .hbl-toast {
      display: flex; align-items: center; gap: 10px; background: var(--surface-raised); border: 1px solid var(--green);
      color: var(--ink); padding: 10px 14px; border-radius: 9px; font-size: 13px; margin-bottom: 14px;
    }
    .hbl-toast.warn { border-color: var(--coral); }

    .hbl-grid2 { display: grid; grid-template-columns: 1.3fr 1fr; gap: 18px; align-items: start; }
    @media (max-width: 720px) {
      .hbl-grid2 { grid-template-columns: 1fr; }
      .hbl-mapgrid { grid-template-columns: 1fr; }
      .hbl-nav { overflow-x: auto; }
      table.hbl-table { font-size: 12px; }
    }

    .hbl-swatch { width: 20px; height: 20px; border-radius: 50%; border: 2px solid transparent; cursor: pointer; flex-shrink: 0; }
    .hbl-swatch.sel { border-color: var(--ink); }

    .hbl-segment { display: inline-flex; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; }
    .hbl-segment-btn {
      background: var(--surface-raised); border: none; border-left: 1px solid var(--border); color: var(--ink-dim);
      font-family: var(--body-font); font-size: 12.5px; font-weight: 500; padding: 7px 13px; cursor: pointer;
    }
    .hbl-segment-btn:first-child { border-left: none; }
    .hbl-segment-btn.active { background: var(--gold); color: #1A1306; font-weight: 600; }

    .hbl-subtoggle {
      background: none; border: none; color: var(--ink-dim); cursor: pointer; display: flex; align-items: center;
      padding: 2px; margin-left: -2px;
    }
    .hbl-subtoggle:hover { color: var(--gold-bright); }
    .hbl-subbreak { margin-top: 8px; padding-left: 14px; border-left: 1px dashed var(--border); display: flex; flex-direction: column; gap: 5px; }
    .hbl-subbreak-row { display: flex; justify-content: space-between; font-size: 12px; color: var(--ink-dim); }
    .hbl-subbreak-row span:first-child { color: var(--ink); }

    .hbl-chip {
      display: inline-flex; align-items: center; gap: 5px; padding: 3px 6px 3px 10px; border-radius: 999px;
      background: var(--surface-raised); border: 1px dashed var(--border); font-size: 11.5px; color: var(--ink);
    }
    .hbl-chip button { background: none; border: none; color: var(--ink-dim); cursor: pointer; display: flex; align-items: center; padding: 2px; }
    .hbl-chip button:hover { color: var(--coral); }

    .hbl-modal-overlay {
      position: fixed; inset: 0; background: rgba(8,12,16,0.6); display: flex; align-items: center;
      justify-content: center; z-index: 1000; padding: 20px;
    }
    .hbl-modal-box {
      background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 22px;
      width: 100%; max-width: 320px; box-shadow: 0 16px 50px rgba(0,0,0,0.45);
    }
  `}</style>
);

/* ---------------------------------- small bits ---------------------------------- */

function Seal({ pct, size }) {
  let cls = "empty";
  if (pct !== null) cls = pct > 100 ? "over" : pct >= 0 ? "under" : "empty";
  return (
    <div className={`hbl-seal ${cls}`} style={size ? { width: size, height: size } : undefined}>
      {pct === null ? "—" : `${Math.round(pct)}%`}
    </div>
  );
}

function CategoryRow({ cat, spent, budget, txForCategory }) {
  const [expanded, setExpanded] = useState(false);
  const pct = budget > 0 ? (spent / budget) * 100 : spent > 0 ? 999 : null;
  const fillColor = pct === null ? "var(--ink-dim)" : pct > 100 ? "var(--coral)" : "var(--green)";
  const fillWidth = pct === null ? 0 : Math.min(100, pct);

  const bySub = {};
  for (const t of txForCategory || []) {
    if (!t.subcategory || !t.subcategory.trim()) continue;
    bySub[t.subcategory] = (bySub[t.subcategory] || 0) + t.amount;
  }
  const subEntries = Object.entries(bySub).sort((a, b) => b[1] - a[1]);
  const other = spent - subEntries.reduce((s, [, v]) => s + v, 0);
  const hasSubs = subEntries.length > 0;

  return (
    <div className="hbl-cat-row">
      <Seal pct={pct === null ? null : pct} />
      <div className="hbl-cat-main">
        <div className="hbl-cat-head">
          <span className="hbl-dot" style={{ background: cat.color }} />
          <span className="hbl-cat-name">{cat.name}</span>
          {hasSubs && (
            <button className="hbl-subtoggle" onClick={() => setExpanded(!expanded)}>
              <ChevronRight size={12} style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
            </button>
          )}
          <span className="hbl-cat-figs hbl-mono">
            {fmtMoney(spent)} {budget > 0 ? `of ${fmtMoney(budget)}` : "(no budget set)"}
          </span>
        </div>
        <div className="hbl-track">
          <div className="hbl-fill" style={{ width: `${fillWidth}%`, background: fillColor }} />
        </div>
        {hasSubs && expanded && (
          <div className="hbl-subbreak">
            {subEntries.map(([name, amt]) => (
              <div key={name} className="hbl-subbreak-row">
                <span>{name}</span>
                <span className="hbl-mono">{fmtMoney(amt)}</span>
              </div>
            ))}
            {other > 0.004 && (
              <div className="hbl-subbreak-row" style={{ color: "var(--ink-dim)" }}>
                <span>Other {cat.name}</span>
                <span className="hbl-mono">{fmtMoney(other)}</span>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------- main ---------------------------------- */

function LedgerApp({ user }) {
  const [loading, setLoading] = useState(true);
  const [transactions, setTransactions] = useState([]);
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES);
  const [budgets, setBudgets] = useState({});
  const [userName, setUserName] = useState(user && user.email ? user.email.split("@")[0] : "");
  const [tab, setTab] = useState("dashboard");
  const [month, setMonth] = useState(thisMonth());
  const [viewMode, setViewMode] = useState("month");
  const [ytdYear, setYtdYear] = useState(String(new Date().getFullYear()));
  const [toast, setToast] = useState(null);
  const [categoryModal, setCategoryModal] = useState(null);

  useEffect(() => {
    let gotTx = false;
    let gotCfg = false;
    const checkDone = () => { if (gotTx && gotCfg) setLoading(false); };

    const unsubTx = subscribeDoc("transactions", (data) => {
      setTransactions(Array.isArray(data) ? data : []);
      gotTx = true;
      checkDone();
    });
    const unsubCfg = subscribeDoc("budget-config", (data) => {
      const cfg = data || { categories: DEFAULT_CATEGORIES, budgets: {} };
      setCategories(cfg.categories && cfg.categories.length ? cfg.categories : DEFAULT_CATEGORIES);
      setBudgets(cfg.budgets || {});
      gotCfg = true;
      checkDone();
    });
    return () => {
      unsubTx();
      unsubCfg();
    };
  }, []);

  const persistTransactions = useCallback(async (next) => {
    setTransactions(next);
    const ok = await saveDoc("transactions", next);
    if (!ok) setToast({ type: "warn", text: "Could not save — check your connection and try again." });
  }, []);

  const persistConfig = useCallback(async (nextCats, nextBudgets) => {
    setCategories(nextCats);
    setBudgets(nextBudgets);
    const ok = await saveDoc("budget-config", { categories: nextCats, budgets: nextBudgets });
    if (!ok) setToast({ type: "warn", text: "Could not save — check your connection and try again." });
  }, []);

  function requestAddCategory(callback) {
    setCategoryModal({ mode: "category", callback });
  }
  function requestAddSubcategory(forCategory, callback) {
    setCategoryModal({ mode: "subcategory", forCategory, callback });
  }
  function confirmCategoryModal(rawName) {
    const name = (rawName || "").trim();
    if (!name || !categoryModal) return;
    if (categoryModal.mode === "subcategory") {
      const catName = categoryModal.forCategory;
      const cat = categories.find((c) => c.name === catName);
      const existingSubs = (cat && cat.subcategories) || [];
      const existingMatch = existingSubs.find((s) => s.toLowerCase() === name.toLowerCase());
      const finalName = existingMatch || name;
      if (!existingMatch) {
        const nextCats = categories.map((c) => (c.name === catName ? { ...c, subcategories: [...existingSubs, name] } : c));
        persistConfig(nextCats, budgets);
      }
      if (categoryModal.callback) categoryModal.callback(finalName);
    } else {
      const existing = categories.find((c) => c.name.toLowerCase() === name.toLowerCase());
      let finalName = name;
      if (existing) {
        finalName = existing.name;
      } else {
        const color = PALETTE[categories.length % PALETTE.length];
        persistConfig([...categories, { name, color, group: "expense", subcategories: [] }], budgets);
      }
      if (categoryModal.callback) categoryModal.callback(finalName);
    }
    setCategoryModal(null);
  }

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 4500);
    return () => clearTimeout(t);
  }, [toast]);

  const monthTx = useMemo(() => transactions.filter((t) => t.date && t.date.startsWith(month)), [transactions, month]);
  const totalSpent = useMemo(() => monthTx.reduce((s, t) => s + (t.amount > 0 ? t.amount : 0), 0), [monthTx]);
  const totalBudget = useMemo(() => Object.values(budgets).reduce((s, v) => s + (Number(v) || 0), 0), [budgets]);

  const byCategory = useMemo(() => {
    const map = {};
    for (const t of monthTx) {
      if (t.amount <= 0) continue;
      map[t.category] = (map[t.category] || 0) + t.amount;
    }
    return map;
  }, [monthTx]);

  const availableYears = useMemo(() => {
    const set = new Set(transactions.map((t) => (t.date || "").slice(0, 4)).filter(Boolean));
    set.add(String(new Date().getFullYear()));
    return Array.from(set).sort().reverse();
  }, [transactions]);

  const yearTx = useMemo(() => transactions.filter((t) => t.date && t.date.startsWith(ytdYear)), [transactions, ytdYear]);
  const totalSpentYTD = useMemo(() => yearTx.reduce((s, t) => s + (t.amount > 0 ? t.amount : 0), 0), [yearTx]);
  const monthsElapsed = useMemo(() => {
    const currentYear = String(new Date().getFullYear());
    if (ytdYear === currentYear) return new Date().getMonth() + 1;
    return ytdYear < currentYear ? 12 : 0;
  }, [ytdYear]);
  const totalBudgetYTD = useMemo(() => totalBudget * monthsElapsed, [totalBudget, monthsElapsed]);

  const byCategoryYTD = useMemo(() => {
    const map = {};
    for (const t of yearTx) {
      if (t.amount <= 0) continue;
      map[t.category] = (map[t.category] || 0) + t.amount;
    }
    return map;
  }, [yearTx]);

  if (loading) {
    return (
      <div className="hbl-root">
        <Styles />
        <div className="hbl-shell">
          <div className="hbl-empty">Opening the ledger…</div>
        </div>
      </div>
    );
  }

  return (
    <div className="hbl-root">
      <Styles />
      <div className="hbl-shell">
        <div className="hbl-top">
          <div>
            <h1 className="hbl-title">Household Ledger</h1>
            <div className="hbl-sub"><Users size={13} /> Shared budget &amp; spending tracker</div>
          </div>
          <div className="hbl-who">
            <span className="hbl-who-label">Signed&nbsp;as</span>
            <input
              className="hbl-input"
              style={{ width: 130 }}
              placeholder="Your name"
              value={userName}
              onChange={(e) => setUserName(e.target.value)}
            />
            <button className="hbl-btn hbl-btn-sm" title="Sign out" onClick={() => signOut(auth)}>
              <LogOut size={13} />
            </button>
          </div>
        </div>

        {toast && (
          <div className={`hbl-toast ${toast.type === "warn" ? "warn" : ""}`}>
            {toast.type === "warn" ? <AlertCircle size={15} /> : <Check size={15} />}
            {toast.text}
          </div>
        )}

        <div className="hbl-nav">
          <button className={`hbl-nav-tab ${tab === "dashboard" ? "active" : ""}`} onClick={() => setTab("dashboard")}>
            <PiggyBank size={15} /> Dashboard
          </button>
          <button className={`hbl-nav-tab ${tab === "summary" ? "active" : ""}`} onClick={() => setTab("summary")}>
            <BarChart3 size={15} /> Summary
          </button>
          <button className={`hbl-nav-tab ${tab === "transactions" ? "active" : ""}`} onClick={() => setTab("transactions")}>
            <Receipt size={15} /> Transactions
          </button>
          <button className={`hbl-nav-tab ${tab === "import" ? "active" : ""}`} onClick={() => setTab("import")}>
            <Upload size={15} /> Import
          </button>
          <button className={`hbl-nav-tab ${tab === "categories" ? "active" : ""}`} onClick={() => setTab("categories")}>
            <Tags size={15} /> Categories
          </button>
          <button className={`hbl-nav-tab ${tab === "budgets" ? "active" : ""}`} onClick={() => setTab("budgets")}>
            <Settings2 size={15} /> Budgets
          </button>
        </div>

        {tab === "dashboard" && (
          <Dashboard
            month={month} setMonth={setMonth}
            viewMode={viewMode} setViewMode={setViewMode}
            ytdYear={ytdYear} setYtdYear={setYtdYear} availableYears={availableYears}
            categories={categories} budgets={budgets}
            byCategory={byCategory} totalSpent={totalSpent} totalBudget={totalBudget} monthTx={monthTx}
            byCategoryYTD={byCategoryYTD} totalSpentYTD={totalSpentYTD} totalBudgetYTD={totalBudgetYTD}
            yearTx={yearTx} monthsElapsed={monthsElapsed}
            transactions={transactions}
          />
        )}
        {tab === "summary" && (
          <SummaryView transactions={transactions} categories={categories} budgets={budgets} />
        )}
        {tab === "transactions" && (
          <TransactionsView
            transactions={transactions} categories={categories}
            persistTransactions={persistTransactions} userName={userName}
            requestAddCategory={requestAddCategory} requestAddSubcategory={requestAddSubcategory}
          />
        )}
        {tab === "import" && (
          <ImportView
            transactions={transactions} persistTransactions={persistTransactions}
            userName={userName} setToast={setToast} categories={categories}
            requestAddCategory={requestAddCategory} requestAddSubcategory={requestAddSubcategory}
          />
        )}
        {tab === "categories" && (
          <CategoriesView
            categories={categories} budgets={budgets} transactions={transactions}
            persistConfig={persistConfig} persistTransactions={persistTransactions}
            requestAddSubcategory={requestAddSubcategory}
          />
        )}
        {tab === "budgets" && (
          <BudgetsView
            categories={categories} budgets={budgets} transactions={transactions}
            persistConfig={persistConfig}
          />
        )}
      </div>
      {categoryModal && (
        <AddCategoryModal
          title={categoryModal.mode === "subcategory" ? `New subcategory for "${categoryModal.forCategory}"` : "New category"}
          placeholder={categoryModal.mode === "subcategory" ? "e.g. Produce" : "e.g. Home Improvement"}
          buttonLabel={categoryModal.mode === "subcategory" ? "Add subcategory" : "Add category"}
          onConfirm={confirmCategoryModal}
          onCancel={() => setCategoryModal(null)}
        />
      )}
    </div>
  );
}

/* ---------------------------------- add category modal ---------------------------------- */

function AddCategoryModal({ title, placeholder, buttonLabel, onConfirm, onCancel }) {
  const [name, setName] = useState("");
  return (
    <div className="hbl-modal-overlay" onClick={onCancel}>
      <div className="hbl-modal-box" onClick={(e) => e.stopPropagation()}>
        <div className="hbl-section-title" style={{ marginBottom: 12 }}>{title || "New category"}</div>
        <label className="hbl-label">Name</label>
        <input
          className="hbl-input" autoFocus placeholder={placeholder || "e.g. Home Improvement"}
          value={name} onChange={(e) => setName(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && name.trim()) onConfirm(name);
            if (e.key === "Escape") onCancel();
          }}
        />
        <div className="hbl-row" style={{ marginTop: 18, justifyContent: "flex-end", gap: 8 }}>
          <button className="hbl-btn" onClick={onCancel}>Cancel</button>
          <button className="hbl-btn hbl-btn-primary" disabled={!name.trim()} onClick={() => onConfirm(name)}>
            <Plus size={14} /> {buttonLabel || "Add category"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------- dashboard ---------------------------------- */

function Dashboard({
  month, setMonth, viewMode, setViewMode, ytdYear, setYtdYear, availableYears,
  categories, budgets, byCategory, totalSpent, totalBudget, monthTx,
  byCategoryYTD, totalSpentYTD, totalBudgetYTD, yearTx, monthsElapsed, transactions,
}) {
  if (transactions.length === 0) {
    return (
      <div className="hbl-card hbl-empty">
        <Receipt size={34} />
        <div style={{ fontFamily: "var(--display-font)", fontSize: 17, color: "var(--ink)", marginBottom: 6 }}>
          No entries yet
        </div>
        <div style={{ fontSize: 13.5 }}>Import a statement from the Import tab to start tracking spending.</div>
      </div>
    );
  }

  const isYTD = viewMode === "ytd";
  const activeTx = isYTD ? yearTx : monthTx;
  const activeByCategory = isYTD ? byCategoryYTD : byCategory;
  const activeSpent = isYTD ? totalSpentYTD : totalSpent;
  const activeBudgetTotal = isYTD ? totalBudgetYTD : totalBudget;
  const budgetFor = (name) => (isYTD ? (Number(budgets[name]) || 0) * monthsElapsed : Number(budgets[name]) || 0);
  const periodLabel = isYTD ? `year to date (${ytdYear})` : `in ${fmtMonthLabel(month)}`;

  const sorted = [...categories].sort((a, b) => (activeByCategory[b.name] || 0) - (activeByCategory[a.name] || 0));
  const chartData = sorted.filter((c) => (activeByCategory[c.name] || 0) > 0).map((c) => ({ name: c.name, value: activeByCategory[c.name], color: c.color }));
  const overallPct = activeBudgetTotal > 0 ? (activeSpent / activeBudgetTotal) * 100 : null;
  const recent = [...activeTx].sort((a, b) => (b.date > a.date ? 1 : -1) || b.createdAt - a.createdAt).slice(0, 6);

  return (
    <div>
      <div className="hbl-row" style={{ marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <div className="hbl-segment">
            <button className={`hbl-segment-btn ${!isYTD ? "active" : ""}`} onClick={() => setViewMode("month")}>This Month</button>
            <button className={`hbl-segment-btn ${isYTD ? "active" : ""}`} onClick={() => setViewMode("ytd")}>Year to Date</button>
          </div>
          {isYTD ? (
            <select className="hbl-select" style={{ width: 100 }} value={ytdYear} onChange={(e) => setYtdYear(e.target.value)}>
              {availableYears.map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
          ) : (
            <input type="month" className="hbl-input" style={{ width: 170 }} value={month} onChange={(e) => setMonth(e.target.value)} />
          )}
        </div>
        <span className="hbl-pill"><Stamp size={12} /> {activeTx.length} entries {isYTD ? `in ${ytdYear}` : "this month"}</span>
      </div>

      <div className="hbl-grid2">
        <div className="hbl-card">
          <div className="hbl-totalbar">
            <Seal pct={overallPct} size={72} />
            <div>
              <div className="hbl-total-num hbl-mono">{fmtMoney(activeSpent)}</div>
              <div className="hbl-total-label">
                Spent {periodLabel}{activeBudgetTotal > 0 ? ` · budgeted ${fmtMoney(activeBudgetTotal)}${isYTD ? " to date" : ""}` : " · no budget set yet"}
              </div>
            </div>
          </div>
          <div className="hbl-section-title" style={{ marginTop: 4 }}>By category</div>
          {sorted.filter((c) => (activeByCategory[c.name] || 0) > 0 || budgetFor(c.name) > 0).length === 0 ? (
            <div className="hbl-empty" style={{ padding: "20px 0" }}>No spending recorded for this period.</div>
          ) : (
            sorted
              .filter((c) => (activeByCategory[c.name] || 0) > 0 || budgetFor(c.name) > 0)
              .map((c) => (
                <CategoryRow
                  key={c.name} cat={c} spent={activeByCategory[c.name] || 0} budget={budgetFor(c.name)}
                  txForCategory={activeTx.filter((t) => t.category === c.name && t.amount > 0)}
                />
              ))
          )}
        </div>

        <div>
          <div className="hbl-card">
            <div className="hbl-section-title">Breakdown</div>
            {chartData.length === 0 ? (
              <div className="hbl-empty" style={{ padding: "20px 0" }}>Nothing to chart yet.</div>
            ) : (
              <div style={{ width: "100%", height: 230 }}>
                <ResponsiveContainer>
                  <PieChart>
                    <Pie data={chartData} dataKey="value" nameKey="name" innerRadius={48} outerRadius={78} paddingAngle={2}>
                      {chartData.map((d, i) => <Cell key={i} fill={d.color} stroke="var(--surface)" strokeWidth={2} />)}
                    </Pie>
                    <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ background: "#1C2A37", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, color: "#EFEBE1", fontSize: 12 }} />
                    <Legend wrapperStyle={{ fontSize: 11, color: "#9FB0BD" }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>

          <div className="hbl-card">
            <div className="hbl-section-title">Recent entries</div>
            {recent.length === 0 ? (
              <div className="hbl-empty" style={{ padding: "12px 0" }}>None yet.</div>
            ) : (
              recent.map((t) => (
                <div key={t.id} className="hbl-row" style={{ padding: "7px 0", borderBottom: "1px dashed var(--border)", fontSize: 12.5 }}>
                  <div style={{ minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    <div>{t.description}</div>
                    <div style={{ color: "var(--ink-dim)", fontSize: 11 }}>{fmtDate(t.date)} · {t.category}</div>
                  </div>
                  <span className={t.amount > 0 ? "hbl-amt-out" : "hbl-amt-in"}>
                    {t.amount > 0 ? "-" : "+"}{fmtMoney(Math.abs(t.amount))}
                  </span>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------- transactions ---------------------------------- */

/* ---------------------------------- summary ---------------------------------- */

function SummaryCategoryTable({ rows, subBreakdownByCategory, expandedCat, setExpandedCat, spentLabel, emptyText }) {
  if (rows.length === 0) {
    return <div className="hbl-empty" style={{ padding: "20px 0" }}>{emptyText}</div>;
  }
  return (
    <div style={{ overflowX: "auto" }}>
      <table className="hbl-table">
        <thead>
          <tr><th>Category</th><th style={{ textAlign: "right" }}>{spentLabel}</th><th style={{ textAlign: "right" }}>Budget</th><th style={{ textAlign: "right" }}>Variance</th></tr>
        </thead>
        <tbody>
          {rows.map((r) => {
            const variance = r.budget - r.spent;
            const breakdown = subBreakdownByCategory[r.name] || { entries: [], other: 0 };
            const hasSubs = breakdown.entries.length > 0;
            const isExpanded = expandedCat === r.name;
            return (
              <Fragment key={r.name}>
                <tr>
                  <td>
                    {hasSubs && (
                      <button
                        className="hbl-subtoggle" style={{ marginRight: 4 }}
                        onClick={() => setExpandedCat(isExpanded ? null : r.name)}
                      >
                        <ChevronRight size={12} style={{ transform: isExpanded ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
                      </button>
                    )}
                    <span className="hbl-dot" style={{ background: r.color, marginRight: 7 }} />{r.name}
                  </td>
                  <td className="hbl-mono" style={{ textAlign: "right" }}>{fmtMoney(r.spent)}</td>
                  <td className="hbl-mono" style={{ textAlign: "right", color: "var(--ink-dim)" }}>{r.budget > 0 ? fmtMoney(r.budget) : "—"}</td>
                  <td className={`hbl-mono ${r.budget > 0 ? (variance < 0 ? "hbl-amt-out" : "hbl-amt-in") : ""}`} style={{ textAlign: "right" }}>
                    {r.budget > 0 ? `${variance < 0 ? "-" : "+"}${fmtMoney(Math.abs(variance))}` : "—"}
                  </td>
                </tr>
                {isExpanded && breakdown.entries.map(([name, amt]) => (
                  <tr key={`${r.name}-${name}`} style={{ fontSize: 12.5, color: "var(--ink-dim)" }}>
                    <td style={{ paddingLeft: 34 }}>{name}</td>
                    <td className="hbl-mono" style={{ textAlign: "right" }}>{fmtMoney(amt)}</td>
                    <td></td>
                    <td></td>
                  </tr>
                ))}
                {isExpanded && breakdown.other > 0.004 && (
                  <tr style={{ fontSize: 12.5, color: "var(--ink-dim)" }}>
                    <td style={{ paddingLeft: 34 }}>Other {r.name}</td>
                    <td className="hbl-mono" style={{ textAlign: "right" }}>{fmtMoney(breakdown.other)}</td>
                    <td></td>
                    <td></td>
                  </tr>
                )}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function SummaryView({ transactions, categories, budgets }) {
  const [mode, setMode] = useState("month");
  const [month, setMonth] = useState(thisMonth());
  const [year, setYear] = useState(String(new Date().getFullYear()));
  const [compareSpan, setCompareSpan] = useState(6);

  const availableYears = useMemo(() => {
    const set = new Set(transactions.map((t) => (t.date || "").slice(0, 4)).filter(Boolean));
    set.add(String(new Date().getFullYear()));
    return Array.from(set).sort().reverse();
  }, [transactions]);

  const activeTx = useMemo(() => {
    if (mode === "month") return transactions.filter((t) => t.date && t.date.startsWith(month));
    return transactions.filter((t) => t.date && t.date.startsWith(year));
  }, [transactions, mode, month, year]);

  const monthsForBudget = useMemo(() => {
    if (mode === "month") return 1;
    if (mode === "year") return 12;
    const currentYear = String(new Date().getFullYear());
    if (year === currentYear) return new Date().getMonth() + 1;
    return year < currentYear ? 12 : 0;
  }, [mode, year]);

  const totalSpent = useMemo(
    () => activeTx.reduce((s, t) => (groupOf(categories, t.category) === "expense" && t.amount > 0 ? s + t.amount : s), 0),
    [activeTx, categories]
  );
  const totalIncome = useMemo(
    () => activeTx.reduce((s, t) => (groupOf(categories, t.category) === "income" && t.amount < 0 ? s - t.amount : s), 0),
    [activeTx, categories]
  );
  const totalBudget = useMemo(
    () => Object.values(budgets).reduce((s, v) => s + (Number(v) || 0), 0) * monthsForBudget,
    [budgets, monthsForBudget]
  );
  const overallPct = totalBudget > 0 ? (totalSpent / totalBudget) * 100 : totalSpent > 0 ? 999 : null;

  const byCategory = useMemo(() => {
    const map = {};
    for (const t of activeTx) {
      const cat = findCategory(categories, t.category);
      const canonicalName = cat ? cat.name : t.category;
      const isIncome = catGroup(cat) === "income";
      if (isIncome) {
        if (t.amount >= 0) continue;
        map[canonicalName] = (map[canonicalName] || 0) - t.amount;
      } else {
        if (t.amount <= 0) continue;
        map[canonicalName] = (map[canonicalName] || 0) + t.amount;
      }
    }
    return map;
  }, [activeTx, categories]);

  const categoryRows = useMemo(() => {
    return categories
      .map((c) => ({
        name: c.name,
        color: c.color,
        group: catGroup(c),
        spent: byCategory[c.name] || 0,
        budget: (Number(budgets[c.name]) || 0) * monthsForBudget,
      }))
      .sort((a, b) => b.spent - a.spent);
  }, [categories, byCategory, budgets, monthsForBudget]);

  const incomeRows = useMemo(() => categoryRows.filter((r) => r.group === "income"), [categoryRows]);
  const expenseRows = useMemo(() => categoryRows.filter((r) => r.group !== "income"), [categoryRows]);

  const [expandedCat, setExpandedCat] = useState(null);

  const subBreakdownByCategory = useMemo(() => {
    const map = {};
    for (const c of categories) {
      const isIncome = catGroup(c) === "income";
      const txForCat = activeTx.filter((t) => normCatName(t.category) === normCatName(c.name) && (isIncome ? t.amount < 0 : t.amount > 0));
      const bySub = {};
      for (const t of txForCat) {
        if (!t.subcategory || !t.subcategory.trim()) continue;
        const val = isIncome ? -t.amount : t.amount;
        bySub[t.subcategory] = (bySub[t.subcategory] || 0) + val;
      }
      const entries = Object.entries(bySub).sort((a, b) => b[1] - a[1]);
      const total = txForCat.reduce((s, t) => s + (isIncome ? -t.amount : t.amount), 0);
      const other = total - entries.reduce((s, [, v]) => s + v, 0);
      map[c.name] = { entries, other };
    }
    return map;
  }, [categories, activeTx]);

  const monthlyBreakdown = useMemo(() => {
    if (mode !== "year") return [];
    return Array.from({ length: 12 }, (_, i) => {
      const m = `${year}-${String(i + 1).padStart(2, "0")}`;
      const total = transactions
        .filter((t) => t.date && t.date.startsWith(m) && t.amount > 0 && groupOf(categories, t.category) === "expense")
        .reduce((s, t) => s + t.amount, 0);
      return { month: m, label: fmtMonthLabel(m).split(" ")[0], total };
    });
  }, [mode, year, transactions, categories]);

  const compareMonths = useMemo(() => {
    if (mode !== "compare") return [];
    const anchor = new Date();
    const list = [];
    for (let i = compareSpan - 1; i >= 0; i--) {
      const d = new Date(anchor.getFullYear(), anchor.getMonth() - i, 1);
      list.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
    }
    return list;
  }, [mode, compareSpan]);

  const compareData = useMemo(() => {
    if (mode !== "compare") return { months: [], rows: [], totals: {} };
    const totals = {};
    const byCatByMonth = {};
    for (const m of compareMonths) {
      totals[m] = 0;
      byCatByMonth[m] = {};
    }
    for (const t of transactions) {
      if (t.amount <= 0 || !t.date) continue;
      if (groupOf(categories, t.category) === "income") continue;
      const m = t.date.slice(0, 7);
      if (!(m in totals)) continue;
      totals[m] += t.amount;
      byCatByMonth[m][t.category] = (byCatByMonth[m][t.category] || 0) + t.amount;
    }
    const catNames = categories.map((c) => c.name);
    const seen = new Set(catNames);
    for (const m of compareMonths) {
      for (const cn of Object.keys(byCatByMonth[m])) seen.add(cn);
    }
    const rows = Array.from(seen)
      .map((name) => {
        const cat = categories.find((c) => c.name === name);
        const byMonth = compareMonths.map((m) => byCatByMonth[m][name] || 0);
        const rowTotal = byMonth.reduce((s, v) => s + v, 0);
        return { name, color: (cat && cat.color) || "#7A8B99", byMonth, rowTotal };
      })
      .filter((r) => r.rowTotal > 0)
      .sort((a, b) => b.rowTotal - a.rowTotal);
    return { months: compareMonths, rows, totals };
  }, [mode, compareMonths, transactions, categories]);

  const periodLabel = mode === "month" ? `in ${fmtMonthLabel(month)}` : mode === "year" ? `for all of ${year}` : mode === "compare" ? `over the last ${compareSpan} months` : `year to date (${year})`;

  function exportSummaryCSV() {
    const header = ["Category", "Spent", "Budget", "Variance"];
    const lines = [header.join(",")].concat(
      categoryRows.map((r) => [r.name, r.spent.toFixed(2), r.budget.toFixed(2), (r.budget - r.spent).toFixed(2)].join(","))
    );
    const blob = new Blob([lines.join("\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `summary-${mode === "month" ? month : year}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  if (transactions.length === 0) {
    return (
      <div className="hbl-card hbl-empty">
        <BarChart3 size={34} />
        <div style={{ fontFamily: "var(--display-font)", fontSize: 17, color: "var(--ink)", marginBottom: 6 }}>No entries yet</div>
        <div style={{ fontSize: 13.5 }}>Import a statement to see a summary here.</div>
      </div>
    );
  }

  return (
    <div>
      <div className="hbl-row" style={{ marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <div className="hbl-segment">
            <button className={`hbl-segment-btn ${mode === "month" ? "active" : ""}`} onClick={() => setMode("month")}>Month</button>
            <button className={`hbl-segment-btn ${mode === "ytd" ? "active" : ""}`} onClick={() => setMode("ytd")}>Year to Date</button>
            <button className={`hbl-segment-btn ${mode === "year" ? "active" : ""}`} onClick={() => setMode("year")}>Year</button>
            <button className={`hbl-segment-btn ${mode === "compare" ? "active" : ""}`} onClick={() => setMode("compare")}>Compare Months</button>
          </div>
          {mode === "month" ? (
            <input type="month" className="hbl-input" style={{ width: 170 }} value={month} onChange={(e) => setMonth(e.target.value)} />
          ) : mode === "compare" ? (
            <select className="hbl-select" style={{ width: 150 }} value={compareSpan} onChange={(e) => setCompareSpan(Number(e.target.value))}>
              <option value={3}>Last 3 months</option>
              <option value={6}>Last 6 months</option>
              <option value={12}>Last 12 months</option>
            </select>
          ) : (
            <select className="hbl-select" style={{ width: 100 }} value={year} onChange={(e) => setYear(e.target.value)}>
              {availableYears.map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
          )}
        </div>
        {mode !== "compare" && <button className="hbl-btn" onClick={exportSummaryCSV}><Download size={14} /> Export</button>}
      </div>

      {mode === "compare" ? (
        <CompareMonthsView compareData={compareData} />
      ) : (
        <>
      <div className="hbl-card" style={{ marginBottom: 16 }}>
        <div className="hbl-row" style={{ flexWrap: "wrap", gap: 24 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div className="hbl-seal" style={{ borderColor: "var(--green)", color: "var(--green)", width: 60, height: 60 }}>
              <ArrowRight size={20} style={{ transform: "rotate(-45deg)" }} />
            </div>
            <div>
              <div className="hbl-total-num hbl-mono" style={{ fontSize: 26, color: "var(--green)" }}>{fmtMoney(totalIncome)}</div>
              <div className="hbl-total-label">Income {periodLabel}</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <Seal pct={overallPct} size={60} />
            <div>
              <div className="hbl-total-num hbl-mono" style={{ fontSize: 26 }}>{fmtMoney(totalSpent)}</div>
              <div className="hbl-total-label">
                Spent {periodLabel}{totalBudget > 0 ? ` · budgeted ${fmtMoney(totalBudget)}` : " · no budget set"}
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="hbl-card" style={{ marginBottom: 16 }}>
        <div className="hbl-section-title">Income</div>
        <SummaryCategoryTable
          rows={incomeRows} subBreakdownByCategory={subBreakdownByCategory}
          expandedCat={expandedCat} setExpandedCat={setExpandedCat}
          spentLabel="Received" emptyText="No income categories yet — add one on the Categories tab and mark it as Income."
        />
      </div>

      <div className="hbl-card" style={{ marginBottom: 16 }}>
        <div className="hbl-section-title">Expenses</div>
        <SummaryCategoryTable
          rows={expenseRows} subBreakdownByCategory={subBreakdownByCategory}
          expandedCat={expandedCat} setExpandedCat={setExpandedCat}
          spentLabel="Spent" emptyText="No spending recorded for this period."
        />
      </div>

      {mode === "year" && (
        <div className="hbl-card">
          <div className="hbl-section-title">Month by month</div>
          <div style={{ overflowX: "auto" }}>
            <table className="hbl-table">
              <thead><tr><th>Month</th><th style={{ textAlign: "right" }}>Spent</th></tr></thead>
              <tbody>
                {monthlyBreakdown.map((m) => (
                  <tr key={m.month}>
                    <td>{m.label}</td>
                    <td className="hbl-mono" style={{ textAlign: "right" }}>{fmtMoney(m.total)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
        </>
      )}
    </div>
  );
}

function CompareMonthsView({ compareData }) {
  const { months, rows, totals } = compareData;
  const grandTotal = months.reduce((s, m) => s + (totals[m] || 0), 0);
  const chartData = months.map((m) => ({ label: fmtMonthLabel(m).split(" ")[0] + " " + fmtMonthLabel(m).split(" ")[1].slice(2), total: totals[m] || 0 }));

  if (months.length === 0 || grandTotal === 0) {
    return (
      <div className="hbl-card hbl-empty">
        <BarChart3 size={34} />
        <div style={{ fontFamily: "var(--display-font)", fontSize: 17, color: "var(--ink)", marginBottom: 6 }}>Nothing to compare yet</div>
        <div style={{ fontSize: 13.5 }}>Import or add entries across a few months to see them compared here.</div>
      </div>
    );
  }

  return (
    <div>
      <div className="hbl-card" style={{ marginBottom: 16 }}>
        <div className="hbl-section-title">Total spend by month</div>
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <BarChart data={chartData} margin={{ top: 4, right: 8, left: 8, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" vertical={false} />
              <XAxis dataKey="label" tick={{ fill: "#9FB0BD", fontSize: 11 }} axisLine={{ stroke: "rgba(255,255,255,0.1)" }} tickLine={false} />
              <YAxis tick={{ fill: "#9FB0BD", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${v >= 1000 ? `${Math.round(v / 1000)}k` : v}`} />
              <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ background: "#1C2A37", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 8, color: "#EFEBE1", fontSize: 12 }} cursor={{ fill: "rgba(255,255,255,0.04)" }} />
              <Bar dataKey="total" fill="var(--gold)" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="hbl-card">
        <div className="hbl-section-title">By category, month by month</div>
        {rows.length === 0 ? (
          <div className="hbl-empty" style={{ padding: "20px 0" }}>No spending recorded across these months.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table className="hbl-table">
              <thead>
                <tr>
                  <th>Category</th>
                  {months.map((m) => <th key={m} style={{ textAlign: "right", whiteSpace: "nowrap" }}>{fmtMonthLabel(m).split(" ")[0]}</th>)}
                  <th style={{ textAlign: "right" }}>Total</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.name}>
                    <td><span className="hbl-dot" style={{ background: r.color, marginRight: 7 }} />{r.name}</td>
                    {r.byMonth.map((v, i) => (
                      <td key={months[i]} className="hbl-mono" style={{ textAlign: "right", color: v > 0 ? "var(--ink)" : "var(--ink-dim)" }}>
                        {v > 0 ? fmtMoney(v) : "—"}
                      </td>
                    ))}
                    <td className="hbl-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmtMoney(r.rowTotal)}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{ borderTop: "1px solid var(--border)" }}>
                  <td style={{ fontWeight: 600, paddingTop: 10 }}>Total</td>
                  {months.map((m) => (
                    <td key={m} className="hbl-mono" style={{ textAlign: "right", fontWeight: 600, paddingTop: 10 }}>{fmtMoney(totals[m] || 0)}</td>
                  ))}
                  <td className="hbl-mono" style={{ textAlign: "right", fontWeight: 600, paddingTop: 10 }}>{fmtMoney(grandTotal)}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function TransactionsView({ transactions, categories, persistTransactions, userName, requestAddCategory, requestAddSubcategory }) {
  const [search, setSearch] = useState("");
  const [filterCat, setFilterCat] = useState("");
  const [filterSub, setFilterSub] = useState("");
  const [periodMode, setPeriodMode] = useState("all");
  const [filterMonth, setFilterMonth] = useState(thisMonth());
  const [filterYear, setFilterYear] = useState(String(new Date().getFullYear()));
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), description: "", amount: "", category: categories[0]?.name || "Uncategorized", subcategory: "", source: "Manual" });

  const availableYears = useMemo(() => {
    const set = new Set(transactions.map((t) => (t.date || "").slice(0, 4)).filter(Boolean));
    set.add(String(new Date().getFullYear()));
    return Array.from(set).sort().reverse();
  }, [transactions]);

  const [sortField, setSortField] = useState("date");
  const [sortDir, setSortDir] = useState("desc");

  function toggleSort(field) {
    if (sortField === field) {
      setSortDir(sortDir === "asc" ? "desc" : "asc");
    } else {
      setSortField(field);
      setSortDir("asc");
    }
  }

  function sortIndicator(field) {
    if (sortField !== field) return "";
    return sortDir === "asc" ? " ▲" : " ▼";
  }

  function compareField(a, b, field) {
    switch (field) {
      case "date":
        return (a.date || "").localeCompare(b.date || "");
      case "description":
        return (a.description || "").toLowerCase().localeCompare((b.description || "").toLowerCase());
      case "category":
        return (a.category || "").toLowerCase().localeCompare((b.category || "").toLowerCase());
      case "subcategory":
        return (a.subcategory || "").toLowerCase().localeCompare((b.subcategory || "").toLowerCase());
      case "source":
        return (a.source || "").toLowerCase().localeCompare((b.source || "").toLowerCase());
      case "amount":
        return a.amount - b.amount;
      default:
        return 0;
    }
  }

  function handleCategorySelect(value, apply) {
    if (value === "__add__") {
      requestAddCategory(apply);
      return;
    }
    apply(value);
  }
  function handleSubcategorySelect(categoryName, value, apply) {
    if (value === "__add__") {
      requestAddSubcategory(categoryName, apply);
      return;
    }
    apply(value);
  }

  const subsForFilter = filterCat ? ((categories.find((c) => c.name === filterCat) || {}).subcategories || []) : [];

  const filtered = useMemo(() => {
    return transactions
      .filter((t) => !filterCat || t.category === filterCat)
      .filter((t) => !filterSub || t.subcategory === filterSub)
      .filter((t) => {
        if (periodMode === "month") return (t.date || "").startsWith(filterMonth);
        if (periodMode === "year") return (t.date || "").startsWith(filterYear);
        return true;
      })
      .filter((t) => !search || t.description.toLowerCase().includes(search.toLowerCase()))
      .sort((a, b) => {
        const cmp = compareField(a, b, sortField);
        return sortDir === "asc" ? cmp : -cmp;
      });
  }, [transactions, search, filterCat, filterSub, periodMode, filterMonth, filterYear, sortField, sortDir]);

  function updateCategory(id, category) {
    persistTransactions(transactions.map((t) => (t.id === id ? { ...t, category, subcategory: "" } : t)));
  }
  function updateSubcategory(id, subcategory) {
    persistTransactions(transactions.map((t) => (t.id === id ? { ...t, subcategory } : t)));
  }
  function removeTx(id) {
    persistTransactions(transactions.filter((t) => t.id !== id));
  }
  function flipSign(id) {
    persistTransactions(transactions.map((t) => (t.id === id ? { ...t, amount: -t.amount } : t)));
  }
  const [confirmFlipAll, setConfirmFlipAll] = useState(false);
  function bulkFlipSign() {
    const ids = new Set(filtered.map((t) => t.id));
    persistTransactions(transactions.map((t) => (ids.has(t.id) ? { ...t, amount: -t.amount } : t)));
    setConfirmFlipAll(false);
  }
  function addManual() {
    const amt = parseAmount(form.amount);
    const date = parseDateString(form.date);
    if (!date || amt === null || !form.description.trim()) {
      return;
    }
    const tx = { id: makeId(), date, description: form.description.trim(), amount: amt, category: form.category, subcategory: form.subcategory || "", source: form.source || "Manual", addedBy: userName || "Someone", createdAt: Date.now() };
    persistTransactions([...transactions, tx]);
    setForm({ ...form, description: "", amount: "" });
    setShowAdd(false);
  }
  function exportCSV() {
    const header = ["Date", "Description", "Amount", "Category", "Subcategory", "Source", "Added By"];
    const lines = [header.join(",")].concat(
      filtered.map((t) => [t.date, `"${t.description.replace(/"/g, '""')}"`, t.amount, t.category, t.subcategory || "", t.source || "", t.addedBy || ""].join(","))
    );
    const blob = new Blob([lines.join("\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "transactions.csv"; a.click();
    URL.revokeObjectURL(url);
  }

  return (
    <div>
      <div className="hbl-card" style={{ marginBottom: 16 }}>
        <div className="hbl-row" style={{ gap: 10, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 160, position: "relative" }}>
            <Search size={14} style={{ position: "absolute", left: 10, top: 10, color: "var(--ink-dim)" }} />
            <input className="hbl-input" style={{ paddingLeft: 30 }} placeholder="Search description…" value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <select className="hbl-select" style={{ width: 160 }} value={filterCat} onChange={(e) => { setFilterCat(e.target.value); setFilterSub(""); }}>
            <option value="">All categories</option>
            {categories.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
          </select>
          {subsForFilter.length > 0 && (
            <select className="hbl-select" style={{ width: 160 }} value={filterSub} onChange={(e) => setFilterSub(e.target.value)}>
              <option value="">All subcategories</option>
              {subsForFilter.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          )}
          <div className="hbl-segment">
            <button className={`hbl-segment-btn ${periodMode === "all" ? "active" : ""}`} onClick={() => setPeriodMode("all")}>All</button>
            <button className={`hbl-segment-btn ${periodMode === "month" ? "active" : ""}`} onClick={() => setPeriodMode("month")}>Month</button>
            <button className={`hbl-segment-btn ${periodMode === "year" ? "active" : ""}`} onClick={() => setPeriodMode("year")}>Year to Date</button>
          </div>
          {periodMode === "month" && (
            <input type="month" className="hbl-input" style={{ width: 150 }} value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)} />
          )}
          {periodMode === "year" && (
            <select className="hbl-select" style={{ width: 100 }} value={filterYear} onChange={(e) => setFilterYear(e.target.value)}>
              {availableYears.map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
          )}
          <button className="hbl-btn" onClick={exportCSV}><Download size={14} /> Export</button>
          {!confirmFlipAll ? (
            <button className="hbl-btn" disabled={filtered.length === 0} onClick={() => setConfirmFlipAll(true)} title="Flip the sign of every entry currently shown by the filters above">
              ± Flip shown ({filtered.length})
            </button>
          ) : (
            <span className="hbl-pill" style={{ gap: 8, padding: "6px 10px" }}>
              Flip all {filtered.length} shown?
              <button className="hbl-btn hbl-btn-sm hbl-btn-primary" onClick={bulkFlipSign}>Yes, flip</button>
              <button className="hbl-btn hbl-btn-sm" onClick={() => setConfirmFlipAll(false)}>Cancel</button>
            </span>
          )}
          <button className="hbl-btn hbl-btn-primary" onClick={() => setShowAdd(!showAdd)}><Plus size={14} /> Add entry</button>
        </div>
        <div style={{ marginTop: 10, fontSize: 12, color: "var(--ink-dim)" }}>
          {filtered.length} entr{filtered.length === 1 ? "y" : "ies"}
          {periodMode === "month" ? ` in ${fmtMonthLabel(filterMonth)}` : periodMode === "year" ? ` in ${filterYear}` : ""}
          {" · "}<span className="hbl-mono">{fmtMoney(filtered.reduce((s, t) => s + (t.amount > 0 ? t.amount : 0), 0))}</span> spent
        </div>

        {showAdd && (
          <div className="hbl-mapgrid" style={{ marginTop: 16, paddingTop: 16, borderTop: "1px solid var(--border)" }}>
            <div><label className="hbl-label">Date</label><input type="date" className="hbl-input" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
            <div><label className="hbl-label">Amount (positive = spent, negative = refund/income)</label><input className="hbl-input" placeholder="24.50" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></div>
            <div><label className="hbl-label">Description</label><input className="hbl-input" placeholder="e.g. Publix" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} /></div>
            <div><label className="hbl-label">Category</label>
              <select className="hbl-select" value={form.category} onChange={(e) => handleCategorySelect(e.target.value, (cat) => setForm({ ...form, category: cat, subcategory: "" }))}>
                {categories.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
                <option value="__add__">+ Add new category…</option>
              </select>
            </div>
            <div><label className="hbl-label">Subcategory (optional)</label>
              <select className="hbl-select" value={form.subcategory} onChange={(e) => handleSubcategorySelect(form.category, e.target.value, (sub) => setForm({ ...form, subcategory: sub }))}>
                <option value="">—</option>
                {((categories.find((c) => c.name === form.category) || {}).subcategories || []).map((s) => <option key={s} value={s}>{s}</option>)}
                <option value="__add__">+ Add new subcategory…</option>
              </select>
            </div>
            <div><label className="hbl-label">Source</label><input className="hbl-input" placeholder="e.g. Chase Checking" value={form.source} onChange={(e) => setForm({ ...form, source: e.target.value })} /></div>
            <div style={{ display: "flex", alignItems: "flex-end" }}>
              <button className="hbl-btn hbl-btn-primary" onClick={addManual} style={{ width: "100%", justifyContent: "center" }}><Check size={14} /> Save entry</button>
            </div>
          </div>
        )}
      </div>

      <div className="hbl-card">
        {filtered.length === 0 ? (
          <div className="hbl-empty">No transactions match.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table className="hbl-table">
              <thead>
                <tr>
                  <th className="hbl-th-sort" onClick={() => toggleSort("date")}>Date{sortIndicator("date")}</th>
                  <th className="hbl-th-sort" onClick={() => toggleSort("description")}>Description{sortIndicator("description")}</th>
                  <th className="hbl-th-sort" onClick={() => toggleSort("category")}>Category{sortIndicator("category")}</th>
                  <th className="hbl-th-sort" onClick={() => toggleSort("subcategory")}>Subcategory{sortIndicator("subcategory")}</th>
                  <th className="hbl-th-sort" onClick={() => toggleSort("source")}>Source{sortIndicator("source")}</th>
                  <th className="hbl-th-sort" style={{ textAlign: "right" }} onClick={() => toggleSort("amount")}>Amount{sortIndicator("amount")}</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((t) => {
                  const subs = (categories.find((c) => c.name === t.category) || {}).subcategories || [];
                  return (
                    <tr key={t.id}>
                      <td className="hbl-mono" style={{ whiteSpace: "nowrap" }}>{fmtDate(t.date)}</td>
                      <td>{t.description}</td>
                      <td>
                        <select className="hbl-select" style={{ padding: "4px 6px", fontSize: 12 }} value={t.category} onChange={(e) => handleCategorySelect(e.target.value, (cat) => updateCategory(t.id, cat))}>
                          {categories.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
                          <option value="__add__">+ Add new category…</option>
                        </select>
                      </td>
                      <td>
                        <select className="hbl-select" style={{ padding: "4px 6px", fontSize: 12 }} value={t.subcategory || ""} onChange={(e) => handleSubcategorySelect(t.category, e.target.value, (sub) => updateSubcategory(t.id, sub))}>
                          <option value="">—</option>
                          {subs.map((s) => <option key={s} value={s}>{s}</option>)}
                          <option value="__add__">+ Add new subcategory…</option>
                        </select>
                      </td>
                      <td style={{ color: "var(--ink-dim)", fontSize: 12 }}>{t.source}</td>
                      <td style={{ textAlign: "right" }} className={t.amount > 0 ? "hbl-amt-out" : "hbl-amt-in"}>
                        {t.amount > 0 ? "-" : "+"}{fmtMoney(Math.abs(t.amount))}
                      </td>
                      <td>
                        <button className="hbl-btn hbl-btn-sm" title="Flip sign (expense ↔ income)" onClick={() => flipSign(t.id)}>±</button>
                        <button className="hbl-btn hbl-btn-sm hbl-btn-danger" style={{ marginLeft: 4 }} onClick={() => removeTx(t.id)}><Trash2 size={13} /></button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------- import ---------------------------------- */

function ImportView({ transactions, persistTransactions, userName, setToast, categories, requestAddCategory, requestAddSubcategory }) {
  const [queue, setQueue] = useState([]);
  const [idx, setIdx] = useState(0);
  const [mapping, setMapping] = useState(null);
  const [sourceName, setSourceName] = useState("");
  const [error, setError] = useState("");
  const [results, setResults] = useState([]);
  const [drag, setDrag] = useState(false);
  const [categoryOverrides, setCategoryOverrides] = useState({});
  const [subcategoryOverrides, setSubcategoryOverrides] = useState({});

  const currentFile = queue[idx];

  function initMapping(file) {
    const dateCol = guessColumn(file.headers, DATE_CANDS);
    const descCol = guessColumn(file.headers, DESC_CANDS);
    const amountCol = guessColumn(file.headers, AMOUNT_CANDS);
    const debitCol = guessColumn(file.headers, DEBIT_CANDS);
    const creditCol = guessColumn(file.headers, CREDIT_CANDS);
    const categoryCol = guessColumn(file.headers, CATEGORY_CANDS);
    const mode = amountCol >= 0 ? "single" : (debitCol >= 0 || creditCol >= 0) ? "split" : "single";
    setMapping({ dateCol, descCol, amountCol, debitCol, creditCol, categoryCol, mode, negate: false });
    setSourceName(file.name.replace(/\.[^.]+$/, ""));
    setCategoryOverrides({});
    setSubcategoryOverrides({});
    setError("");
  }

  async function handleFiles(fileList) {
    const files = Array.from(fileList);
    const parsed = [];
    for (const file of files) {
      try {
        const ext = file.name.split(".").pop().toLowerCase();
        let headers = [], rows = [];
        if (ext === "csv") {
          const text = await file.text();
          const result = Papa.parse(text, { skipEmptyLines: true });
          const data = result.data;
          headers = (data[0] || []).map((h) => String(h));
          rows = data.slice(1);
        } else if (ext === "xlsx" || ext === "xls") {
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(buf, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          headers = (data[0] || []).map((h) => String(h));
          rows = data.slice(1);
        } else {
          continue;
        }
        rows = rows.filter((r) => Array.isArray(r) && r.some((c) => String(c).trim() !== ""));
        if (headers.length && rows.length) parsed.push({ name: file.name, headers, rows });
      } catch (e) {
        console.error("parse error", e);
      }
    }
    if (parsed.length) {
      setQueue(parsed);
      setIdx(0);
      setResults([]);
      initMapping(parsed[0]);
    } else {
      setError("Couldn't read that file. Upload a CSV or Excel (.xlsx) export from your bank or card.");
    }
  }

  function buildFromMapping(file, map, src, catOverrides, subOverrides) {
    const out = [];
    let invalid = 0;
    for (const r of file.rows) {
      const rawDate = map.dateCol >= 0 ? r[map.dateCol] : "";
      const rawDesc = map.descCol >= 0 ? r[map.descCol] : "";
      let amount = null;
      if (map.mode === "single") {
        amount = parseAmount(map.amountCol >= 0 ? r[map.amountCol] : null);
        if (amount !== null && map.negate) amount = -amount;
      } else {
        const debit = map.debitCol >= 0 ? parseAmount(r[map.debitCol]) : null;
        const credit = map.creditCol >= 0 ? parseAmount(r[map.creditCol]) : null;
        if (debit) amount = Math.abs(debit);
        else if (credit) amount = -Math.abs(credit);
      }
      const date = parseDateString(rawDate);
      if (!date || amount === null || amount === 0) { invalid++; continue; }
      const descKey = String(rawDesc || "").trim().toLowerCase();
      let category = map.categoryCol >= 0 && r[map.categoryCol] ? String(r[map.categoryCol]).trim() : guessCategory(rawDesc);
      if (catOverrides && catOverrides[descKey]) category = catOverrides[descKey];
      const subcategory = (subOverrides && subOverrides[descKey]) || "";
      out.push({ id: makeId(), date, description: String(rawDesc || "").trim() || "(no description)", amount, category: category || "Uncategorized", subcategory, source: src, addedBy: userName || "Someone", createdAt: Date.now() });
    }
    return { built: out, invalid };
  }

  const preview = useMemo(() => {
    if (!currentFile || !mapping) return [];
    const { built } = buildFromMapping({ ...currentFile, rows: currentFile.rows.slice(0, 8) }, mapping, sourceName, categoryOverrides, subcategoryOverrides);
    return built;
  }, [currentFile, mapping, sourceName, categoryOverrides, subcategoryOverrides]);

  function setPreviewCategory(description, category) {
    const key = String(description || "").trim().toLowerCase();
    const nextSub = { ...subcategoryOverrides };
    delete nextSub[key];
    setSubcategoryOverrides(nextSub);
    setCategoryOverrides({ ...categoryOverrides, [key]: category });
  }
  function setPreviewSubcategory(description, subcategory) {
    const key = String(description || "").trim().toLowerCase();
    setSubcategoryOverrides({ ...subcategoryOverrides, [key]: subcategory });
  }

  function confirmImport() {
    if (!currentFile || !mapping) return;
    const { built, invalid } = buildFromMapping(currentFile, mapping, sourceName || currentFile.name, categoryOverrides, subcategoryOverrides);
    const existingKeys = new Set(transactions.map(dedupeKey));
    const seen = new Set();
    const keep = [];
    let dup = 0;
    for (const t of built) {
      const k = dedupeKey(t);
      if (existingKeys.has(k) || seen.has(k)) { dup++; continue; }
      seen.add(k);
      keep.push(t);
    }
    const next = [...transactions, ...keep];
    persistTransactions(next);
    setResults((r) => [...r, { name: currentFile.name, imported: keep.length, duplicates: dup, invalid }]);
    if (idx + 1 < queue.length) {
      setIdx(idx + 1);
      initMapping(queue[idx + 1]);
    } else {
      setQueue([]);
      setIdx(0);
      setMapping(null);
      setToast({ type: "ok", text: "Import finished." });
    }
  }

  function skipFile() {
    if (idx + 1 < queue.length) {
      setIdx(idx + 1);
      initMapping(queue[idx + 1]);
    } else {
      setQueue([]);
      setMapping(null);
    }
  }

  if (!currentFile) {
    return (
      <div>
        <div
          className={`hbl-dropzone ${drag ? "drag" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); handleFiles(e.dataTransfer.files); }}
          onClick={() => document.getElementById("hbl-file-input").click()}
        >
          <input id="hbl-file-input" type="file" accept=".csv,.xlsx,.xls" multiple style={{ display: "none" }} onChange={(e) => handleFiles(e.target.files)} />
          <FileSpreadsheet size={30} style={{ color: "var(--gold)", marginBottom: 10 }} />
          <div style={{ fontFamily: "var(--display-font)", fontSize: 16, marginBottom: 4 }}>Drop statements here, or click to browse</div>
          <div style={{ fontSize: 12.5, color: "var(--ink-dim)" }}>CSV or Excel exports from any bank or card — upload several sources at once</div>
        </div>
        {error && <div className="hbl-toast warn" style={{ marginTop: 14 }}><AlertCircle size={15} /> {error}</div>}
        {results.length > 0 && (
          <div className="hbl-card" style={{ marginTop: 16 }}>
            <div className="hbl-section-title">Last import</div>
            {results.map((r, i) => (
              <div key={i} className="hbl-row" style={{ fontSize: 13, padding: "6px 0" }}>
                <span>{r.name}</span>
                <span className="hbl-pill">{r.imported} added · {r.duplicates} duplicates skipped{r.invalid ? ` · ${r.invalid} unreadable rows` : ""}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="hbl-card">
      <div className="hbl-row" style={{ marginBottom: 6 }}>
        <div className="hbl-section-title" style={{ margin: 0 }}>
          <span className="hbl-step-badge">{idx + 1}</span>
          Mapping “{currentFile.name}”{queue.length > 1 ? ` (${idx + 1} of ${queue.length})` : ""}
        </div>
        <button className="hbl-btn hbl-btn-sm" onClick={skipFile}>Skip this file</button>
      </div>
      <div style={{ fontSize: 12.5, color: "var(--ink-dim)", marginBottom: 4 }}>
        Tell us which columns hold what — we guessed based on the headers, double check before importing.
      </div>

      <div className="hbl-mapgrid">
        <div>
          <label className="hbl-label">Source name (e.g. “Chase Checking”)</label>
          <input className="hbl-input" value={sourceName} onChange={(e) => setSourceName(e.target.value)} />
        </div>
        <div>
          <label className="hbl-label">Date column</label>
          <select className="hbl-select" value={mapping.dateCol} onChange={(e) => setMapping({ ...mapping, dateCol: Number(e.target.value) })}>
            <option value={-1}>— none —</option>
            {currentFile.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
          </select>
        </div>
        <div>
          <label className="hbl-label">Description column</label>
          <select className="hbl-select" value={mapping.descCol} onChange={(e) => setMapping({ ...mapping, descCol: Number(e.target.value) })}>
            <option value={-1}>— none —</option>
            {currentFile.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
          </select>
        </div>
        <div>
          <label className="hbl-label">Category column (optional)</label>
          <select className="hbl-select" value={mapping.categoryCol} onChange={(e) => setMapping({ ...mapping, categoryCol: Number(e.target.value) })}>
            <option value={-1}>— guess from description —</option>
            {currentFile.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
          </select>
        </div>

        <div>
          <label className="hbl-label">Amount format</label>
          <select className="hbl-select" value={mapping.mode} onChange={(e) => setMapping({ ...mapping, mode: e.target.value })}>
            <option value="single">Single amount column</option>
            <option value="split">Separate debit / credit columns</option>
          </select>
        </div>

        {mapping.mode === "single" ? (
          <>
            <div>
              <label className="hbl-label">Amount column</label>
              <select className="hbl-select" value={mapping.amountCol} onChange={(e) => setMapping({ ...mapping, amountCol: Number(e.target.value) })}>
                <option value={-1}>— none —</option>
                {currentFile.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", alignItems: "flex-end", paddingBottom: 8 }}>
              <label style={{ fontSize: 12.5, display: "flex", alignItems: "center", gap: 6, color: "var(--ink-dim)" }}>
                <input type="checkbox" checked={mapping.negate} onChange={(e) => setMapping({ ...mapping, negate: e.target.checked })} />
                Flip sign (purchases show as negative in this file)
              </label>
            </div>
          </>
        ) : (
          <>
            <div>
              <label className="hbl-label">Debit column (money out)</label>
              <select className="hbl-select" value={mapping.debitCol} onChange={(e) => setMapping({ ...mapping, debitCol: Number(e.target.value) })}>
                <option value={-1}>— none —</option>
                {currentFile.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
              </select>
            </div>
            <div>
              <label className="hbl-label">Credit column (money in)</label>
              <select className="hbl-select" value={mapping.creditCol} onChange={(e) => setMapping({ ...mapping, creditCol: Number(e.target.value) })}>
                <option value={-1}>— none —</option>
                {currentFile.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
              </select>
            </div>
          </>
        )}
      </div>

      <div className="hbl-section-title">Preview (first rows)</div>
      <div style={{ fontSize: 12, color: "var(--ink-dim)", marginTop: -8, marginBottom: 10 }}>
        Fix a category here and it's applied to every row with that same description.
      </div>
      <div style={{ overflowX: "auto", marginBottom: 16 }}>
        <table className="hbl-table">
          <thead><tr><th>Date</th><th>Description</th><th>Category</th><th>Subcategory</th><th style={{ textAlign: "right" }}>Amount</th></tr></thead>
          <tbody>
            {preview.length === 0 ? (
              <tr><td colSpan={5} style={{ color: "var(--ink-dim)" }}>No rows matched this mapping — adjust the columns above.</td></tr>
            ) : preview.map((t, i) => {
              const subs = (categories.find((c) => c.name === t.category) || {}).subcategories || [];
              return (
                <tr key={i}>
                  <td className="hbl-mono">{fmtDate(t.date)}</td>
                  <td>{t.description}</td>
                  <td>
                    <select
                      className="hbl-select" style={{ padding: "4px 6px", fontSize: 12 }}
                      value={t.category}
                      onChange={(e) => {
                        const value = e.target.value;
                        if (value === "__add__") {
                          requestAddCategory((created) => setPreviewCategory(t.description, created));
                          return;
                        }
                        setPreviewCategory(t.description, value);
                      }}
                    >
                      {!categories.some((c) => c.name === t.category) && <option value={t.category}>{t.category}</option>}
                      {categories.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
                      <option value="__add__">+ Add new category…</option>
                    </select>
                  </td>
                  <td>
                    <select
                      className="hbl-select" style={{ padding: "4px 6px", fontSize: 12 }}
                      value={t.subcategory || ""}
                      onChange={(e) => {
                        const value = e.target.value;
                        if (value === "__add__") {
                          requestAddSubcategory(t.category, (created) => setPreviewSubcategory(t.description, created));
                          return;
                        }
                        setPreviewSubcategory(t.description, value);
                      }}
                    >
                      <option value="">—</option>
                      {subs.map((s) => <option key={s} value={s}>{s}</option>)}
                      <option value="__add__">+ Add new subcategory…</option>
                    </select>
                  </td>
                  <td style={{ textAlign: "right" }} className={t.amount > 0 ? "hbl-amt-out" : "hbl-amt-in"}>{t.amount > 0 ? "-" : "+"}{fmtMoney(Math.abs(t.amount))}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="hbl-row">
        <span style={{ fontSize: 12, color: "var(--ink-dim)" }}>{currentFile.rows.length} rows in file · category and source can be edited after import</span>
        <button className="hbl-btn hbl-btn-primary" onClick={confirmImport} disabled={preview.length === 0}>
          Import this file <ArrowRight size={14} />
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------- budgets ---------------------------------- */

function CategoriesView({ categories, budgets, transactions, persistConfig, persistTransactions, requestAddSubcategory }) {
  const [newCat, setNewCat] = useState("");
  const [colorIdx, setColorIdx] = useState(0);
  const [editing, setEditing] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [editColor, setEditColor] = useState("");
  const [expandedCat, setExpandedCat] = useState(null);
  const [editingSub, setEditingSub] = useState(null);
  const [editSubValue, setEditSubValue] = useState("");

  const last3Avg = useMemo(() => {
    const now = new Date();
    const months = [0, 1, 2].map((n) => {
      const d = new Date(now.getFullYear(), now.getMonth() - n, 1);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    });
    const sums = {};
    for (const t of transactions) {
      if (t.amount <= 0 || !months.some((m) => t.date.startsWith(m))) continue;
      sums[t.category] = (sums[t.category] || 0) + t.amount;
    }
    const avg = {};
    for (const [k, v] of Object.entries(sums)) avg[k] = v / 3;
    return avg;
  }, [transactions]);

  function addCategory() {
    const name = newCat.trim();
    if (!name || categories.some((c) => c.name.toLowerCase() === name.toLowerCase())) return;
    const color = PALETTE[colorIdx % PALETTE.length];
    persistConfig([...categories, { name, color, group: "expense", subcategories: [] }], budgets);
    setNewCat("");
    setColorIdx(colorIdx + 1);
  }
  function removeCategory(name) {
    const nextCats = categories.filter((c) => c.name !== name);
    const nextBudgets = { ...budgets };
    delete nextBudgets[name];
    persistConfig(nextCats, nextBudgets);
  }
  function removeSubcategory(catName, sub) {
    const nextCats = categories.map((c) => (c.name === catName ? { ...c, subcategories: (c.subcategories || []).filter((s) => s !== sub) } : c));
    persistConfig(nextCats, budgets);
    const nextTx = transactions.map((t) => (t.category === catName && t.subcategory === sub ? { ...t, subcategory: "" } : t));
    persistTransactions(nextTx);
  }
  function startEditSub(catName, sub) {
    setEditingSub({ category: catName, sub });
    setEditSubValue(sub);
  }
  function cancelEditSub() {
    setEditingSub(null);
    setEditSubValue("");
  }
  function saveEditSub() {
    if (!editingSub) return;
    const newName = editSubValue.trim();
    if (!newName) return;
    const cat = categories.find((c) => c.name === editingSub.category);
    if (!cat) return;
    const subs = cat.subcategories || [];
    const dupe = newName.toLowerCase() !== editingSub.sub.toLowerCase() &&
      subs.some((s) => s.toLowerCase() === newName.toLowerCase());
    if (dupe) return;
    const nextCats = categories.map((c) => (c.name === editingSub.category ? { ...c, subcategories: subs.map((s) => (s === editingSub.sub ? newName : s)) } : c));
    persistConfig(nextCats, budgets);
    if (newName !== editingSub.sub) {
      const nextTx = transactions.map((t) => (t.category === editingSub.category && t.subcategory === editingSub.sub ? { ...t, subcategory: newName } : t));
      persistTransactions(nextTx);
    }
    cancelEditSub();
  }
  function startEdit(c) {
    setEditing(c.name);
    setEditValue(c.name);
    setEditColor(c.color);
  }
  function cancelEdit() {
    setEditing(null);
    setEditValue("");
    setEditColor("");
  }
  function saveEdit(c) {
    const newName = editValue.trim();
    if (!newName) return;
    const dupe = newName.toLowerCase() !== c.name.toLowerCase() &&
      categories.some((cc) => cc.name.toLowerCase() === newName.toLowerCase());
    if (dupe) return;
    const nextCats = categories.map((cc) => (cc.name === c.name ? { name: newName, color: editColor, group: c.group, subcategories: c.subcategories || [] } : cc));
    const nextBudgets = { ...budgets };
    if (newName !== c.name && c.name in nextBudgets) {
      nextBudgets[newName] = nextBudgets[c.name];
      delete nextBudgets[c.name];
    }
    persistConfig(nextCats, nextBudgets);
    if (newName !== c.name) {
      const nextTx = transactions.map((t) => (t.category === c.name ? { ...t, category: newName } : t));
      persistTransactions(nextTx);
    }
    cancelEdit();
  }

  function setCategoryGroup(name, group) {
    const nextCats = categories.map((c) => (c.name === name ? { ...c, group } : c));
    persistConfig(nextCats, budgets);
  }

  return (
    <div className="hbl-card">
      <div className="hbl-section-title">Categories</div>
      <div style={{ fontSize: 12.5, color: "var(--ink-dim)", marginBottom: 14 }}>
        Rename, recolor, add, or delete categories and subcategories here. Set how much to budget for each on the Budgets tab.
      </div>
      {categories.map((c) => {
        const isEditing = editing === c.name;
        const dupe = isEditing && editValue.trim().toLowerCase() !== c.name.toLowerCase() &&
          categories.some((cc) => cc.name.toLowerCase() === editValue.trim().toLowerCase());
        if (isEditing) {
          return (
            <div key={c.name} style={{ padding: "12px 0", borderBottom: "1px dashed var(--border)" }}>
              <div className="hbl-row" style={{ gap: 10, marginBottom: 10 }}>
                <input className="hbl-input" style={{ flex: 1 }} value={editValue} onChange={(e) => setEditValue(e.target.value)} autoFocus />
                <button className="hbl-btn hbl-btn-sm hbl-btn-primary" onClick={() => saveEdit(c)} disabled={!editValue.trim() || dupe}><Check size={13} /></button>
                <button className="hbl-btn hbl-btn-sm" onClick={cancelEdit}><X size={13} /></button>
              </div>
              {dupe && <div style={{ fontSize: 11.5, color: "var(--coral)", marginBottom: 8 }}>A category with that name already exists.</div>}
              <div style={{ display: "flex", gap: 7 }}>
                {PALETTE.map((p) => (
                  <span key={p} className={`hbl-swatch ${editColor === p ? "sel" : ""}`} style={{ background: p }} onClick={() => setEditColor(p)} />
                ))}
              </div>
            </div>
          );
        }
        const subs = c.subcategories || [];
        const isExpanded = expandedCat === c.name;
        return (
          <div key={c.name} style={{ borderBottom: "1px dashed var(--border)" }}>
            <div className="hbl-row" style={{ padding: "10px 0" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, flex: 1, minWidth: 0 }}>
                <button className="hbl-subtoggle" onClick={() => setExpandedCat(isExpanded ? null : c.name)} title="Manage subcategories">
                  <ChevronRight size={13} style={{ transform: isExpanded ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
                </button>
                <span className="hbl-dot" style={{ background: c.color }} />
                <span style={{ fontSize: 13.5 }}>{c.name}</span>
                {subs.length > 0 && <span className="hbl-pill" style={{ fontSize: 10.5 }}>{subs.length} sub</span>}
                {last3Avg[c.name] ? (
                  <span style={{ fontSize: 11, color: "var(--ink-dim)" }}>avg {fmtMoney(last3Avg[c.name])}/mo</span>
                ) : null}
              </div>
              <div className="hbl-segment">
                <button
                  className={`hbl-segment-btn ${catGroup(c) === "income" ? "active" : ""}`}
                  style={catGroup(c) === "income" ? { background: "var(--green)", color: "#0E1A13" } : undefined}
                  onClick={() => setCategoryGroup(c.name, "income")}
                >
                  Income
                </button>
                <button
                  className={`hbl-segment-btn ${catGroup(c) !== "income" ? "active" : ""}`}
                  onClick={() => setCategoryGroup(c.name, "expense")}
                >
                  Expense
                </button>
              </div>
              <button className="hbl-btn hbl-btn-sm" onClick={() => startEdit(c)}><Pencil size={13} /></button>
              <button className="hbl-btn hbl-btn-sm hbl-btn-danger" onClick={() => removeCategory(c.name)}><Trash2 size={13} /></button>
            </div>
            {isExpanded && (
              <div style={{ padding: "0 0 14px 29px" }}>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
                  {subs.length === 0 && <span style={{ fontSize: 12, color: "var(--ink-dim)" }}>No subcategories yet.</span>}
                  {subs.map((s) => {
                    const isEditingThis = editingSub && editingSub.category === c.name && editingSub.sub === s;
                    if (isEditingThis) {
                      const subDupe = editSubValue.trim().toLowerCase() !== s.toLowerCase() &&
                        subs.some((ss) => ss.toLowerCase() === editSubValue.trim().toLowerCase());
                      return (
                        <span key={s} className="hbl-chip" style={{ paddingLeft: 4, gap: 4 }}>
                          <input
                            className="hbl-input" style={{ width: 96, padding: "2px 6px", fontSize: 11.5 }}
                            value={editSubValue} autoFocus
                            onChange={(e) => setEditSubValue(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter" && editSubValue.trim() && !subDupe) saveEditSub();
                              if (e.key === "Escape") cancelEditSub();
                            }}
                          />
                          <button onClick={saveEditSub} disabled={!editSubValue.trim() || subDupe}><Check size={11} /></button>
                          <button onClick={cancelEditSub}><X size={11} /></button>
                        </span>
                      );
                    }
                    return (
                      <span key={s} className="hbl-chip">
                        {s}
                        <button onClick={() => startEditSub(c.name, s)}><Pencil size={10} /></button>
                        <button onClick={() => removeSubcategory(c.name, s)}><X size={10} /></button>
                      </span>
                    );
                  })}
                </div>
                {editingSub && editingSub.category === c.name &&
                  editSubValue.trim().toLowerCase() !== editingSub.sub.toLowerCase() &&
                  subs.some((ss) => ss.toLowerCase() === editSubValue.trim().toLowerCase()) && (
                    <div style={{ fontSize: 11.5, color: "var(--coral)", marginBottom: 8 }}>A subcategory with that name already exists.</div>
                  )}
                <button className="hbl-btn hbl-btn-sm" onClick={() => requestAddSubcategory(c.name, () => {})}>
                  <Plus size={12} /> Add subcategory
                </button>
              </div>
            )}
          </div>
        );
      })}

      <div className="hbl-row" style={{ marginTop: 16, gap: 10 }}>
        <span className="hbl-swatch sel" style={{ background: PALETTE[colorIdx % PALETTE.length] }} />
        <input className="hbl-input" placeholder="New category name" value={newCat} onChange={(e) => setNewCat(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addCategory()} />
        <button className="hbl-btn hbl-btn-primary" onClick={addCategory}><Plus size={14} /> Add category</button>
      </div>
    </div>
  );
}

function BudgetsView({ categories, budgets, transactions, persistConfig }) {
  const last3Avg = useMemo(() => {
    const now = new Date();
    const months = [0, 1, 2].map((n) => {
      const d = new Date(now.getFullYear(), now.getMonth() - n, 1);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    });
    const sums = {};
    for (const t of transactions) {
      if (t.amount <= 0 || !months.some((m) => t.date.startsWith(m))) continue;
      sums[t.category] = (sums[t.category] || 0) + t.amount;
    }
    const avg = {};
    for (const [k, v] of Object.entries(sums)) avg[k] = v / 3;
    return avg;
  }, [transactions]);

  function setBudget(name, value) {
    persistConfig(categories, { ...budgets, [name]: value === "" ? "" : Number(value) });
  }

  if (categories.length === 0) {
    return (
      <div className="hbl-card hbl-empty">
        <Settings2 size={34} />
        <div style={{ fontFamily: "var(--display-font)", fontSize: 17, color: "var(--ink)", marginBottom: 6 }}>No categories yet</div>
        <div style={{ fontSize: 13.5 }}>Add some on the Categories tab first, then set budgets here.</div>
      </div>
    );
  }

  return (
    <div className="hbl-card">
      <div className="hbl-section-title">Monthly budgets</div>
      <div style={{ fontSize: 12.5, color: "var(--ink-dim)", marginBottom: 14 }}>
        Set what you want to spend per category each month. The 3-month average is shown as a guide. To rename, recolor,
        add, or delete categories and subcategories, use the Categories tab.
      </div>
      {categories.map((c) => (
        <div key={c.name} className="hbl-row" style={{ padding: "10px 0", borderBottom: "1px dashed var(--border)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1, minWidth: 0 }}>
            <span className="hbl-dot" style={{ background: c.color }} />
            <span style={{ fontSize: 13.5 }}>{c.name}</span>
            {last3Avg[c.name] ? (
              <span style={{ fontSize: 11, color: "var(--ink-dim)" }}>avg {fmtMoney(last3Avg[c.name])}/mo</span>
            ) : null}
          </div>
          <input
            className="hbl-input hbl-mono" style={{ width: 110, textAlign: "right" }}
            placeholder="0.00" value={budgets[c.name] ?? ""} onChange={(e) => setBudget(c.name, e.target.value)}
          />
        </div>
      ))}
    </div>
  );
}

/* ---------------------------------- auth gate ---------------------------------- */

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      await signInWithEmailAndPassword(auth, email.trim(), password);
    } catch (err) {
      setError("Could not sign in — check the email and password.");
    }
    setBusy(false);
  }

  return (
    <div className="hbl-root" style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <Styles />
      <form onSubmit={handleSubmit} className="hbl-modal-box" style={{ maxWidth: 320, position: "static" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <Lock size={16} style={{ color: "var(--gold)" }} />
          <div className="hbl-section-title" style={{ margin: 0 }}>Sign in to the ledger</div>
        </div>
        <label className="hbl-label">Email</label>
        <input
          className="hbl-input" style={{ marginBottom: 12 }} type="email" autoFocus required
          value={email} onChange={(e) => setEmail(e.target.value)}
        />
        <label className="hbl-label">Password</label>
        <input
          className="hbl-input" style={{ marginBottom: 16 }} type="password" required
          value={password} onChange={(e) => setPassword(e.target.value)}
        />
        {error && (
          <div className="hbl-toast warn" style={{ marginBottom: 14 }}><AlertCircle size={15} /> {error}</div>
        )}
        <button className="hbl-btn hbl-btn-primary" type="submit" disabled={busy} style={{ width: "100%", justifyContent: "center" }}>
          {busy ? "Signing in…" : "Sign in"}
        </button>
        <div style={{ fontSize: 11.5, color: "var(--ink-dim)", marginTop: 14 }}>
          New team member? Ask whoever set this up to add your email under Authentication → Users in the Firebase console.
        </div>
      </form>
    </div>
  );
}

export default function App() {
  const [user, setUser] = useState(undefined); // undefined = checking, null = signed out

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return unsub;
  }, []);

  if (user === undefined) {
    return (
      <div className="hbl-root" style={{ minHeight: "100vh" }}>
        <Styles />
        <div className="hbl-empty">Loading…</div>
      </div>
    );
  }
  if (!user) {
    return <LoginScreen />;
  }
  return <LedgerApp user={user} />;
}
