import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
  LineChart, Line,
} from "recharts";
import {
  Home, Users, UserRound, Landmark, CalendarClock, TrendingUp, LogOut, Plus,
  MapPin, Phone, ChevronRight, ArrowLeft, X, Search, AlertTriangle, IndianRupee,
  RotateCcw, Check, Wallet, Printer, MessageCircle, Smartphone, Copy, Download, Mail, Upload, RefreshCw, Trash2,
  Archive, PiggyBank,
} from "lucide-react";
import * as XLSX from "xlsx";

/* ============================== SHARED DATA SYNC ==============================
   This app shares its data across every agent and the owner by reading/writing
   a Google Sheet through a small Google Apps Script "bridge" web app.

   SETUP (one-time, done by the owner before deploying this site):
   1. Create a new Google Sheet in your Google Drive.
   2. In the Sheet, go to Extensions -> Apps Script.
   3. Delete anything in the editor and paste the contents of the
      "google-apps-script/Code.gs" file included alongside this project.
   4. Click Deploy -> New deployment -> Web app.
        Execute as: Me
        Who has access: Anyone
   5. Click Deploy, authorize it with your Google account, and copy the
      "Web app URL" it gives you (ends in /exec).
   6. Paste that URL below, replacing the placeholder text.
   7. Redeploy this website (e.g. push to GitHub, Vercel redeploys automatically).

   Until a real URL is pasted in, the app falls back to saving data only on
   this device (same as before), so it still works while you're setting this up.
================================================================================ */

const SYNC_URL = "https://script.google.com/macros/s/AKfycbwg7MubtOgrDvMRbQqxDqzEDJbd0lCKovJUn30hTTsLqHCDh-pqS58YUdmkcXZ-c0xuig/exec";

function syncConfigured() {
  return typeof SYNC_URL === "string" && SYNC_URL.startsWith("http");
}

/* ---- Storage client: shared Google Sheet when configured, else local device only ---- */
const storage = {
  async get(key) {
    if (syncConfigured()) {
      try {
        const res = await fetch(`${SYNC_URL}?t=${Date.now()}`, { method: "GET" });
        const text = await res.text();
        return text ? { key, value: text } : null;
      } catch (e) {
        console.error("shared sync load failed, falling back to this device's saved copy", e);
      }
    }
    try {
      const v = window.localStorage.getItem(key);
      return v !== null ? { key, value: v } : null;
    } catch (e) {
      return null;
    }
  },
  async set(key, value) {
    // Always keep a local copy too, so the app still works offline / mid-setup.
    try {
      window.localStorage.setItem(key, value);
    } catch (e) {}
    if (syncConfigured()) {
      try {
        await fetch(SYNC_URL, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: value,
        });
      } catch (e) {
        console.error("shared sync save failed (saved locally on this device only)", e);
        return null;
      }
    }
    return { key, value };
  },
};

/* ---- Conflict-safe merge for multi-device / patchy-connectivity use -----------
   The Sheet stores one JSON blob. If two Android devices are offline at the
   same time and both come back online, a naive "just overwrite with mine"
   save would silently erase whichever device synced first. This merges the
   freshest remote copy with this device's local changes record-by-record
   (by id) instead of blindly replacing the whole document, so:
     - New agents/clients/loans/reinvestments added on either device survive.
     - A payment recorded on installment X by one agent doesn't get
       overwritten by an unrelated change on another device.
   It is not a perfect distributed database - if the SAME installment is
   paid differently on two devices before either syncs, the higher paid
   amount wins - but for this app's real usage (each agent works their own
   clients) that scenario is rare, and this removes the everyday risk. --- */
function byId(arr) {
  const out = {};
  (arr || []).forEach((x) => { if (x && x.id) out[x.id] = x; });
  return out;
}

function mergeLoan(remoteLoan, localLoan) {
  if (!remoteLoan) return localLoan;
  if (!localLoan) return remoteLoan;
  const rInst = byId(remoteLoan.schedule);
  const lInst = byId(localLoan.schedule);
  const ids = new Set([...Object.keys(rInst), ...Object.keys(lInst)]);
  const schedule = [...ids]
    .map((id) => {
      const ri = rInst[id], li = lInst[id];
      if (!ri) return li;
      if (!li) return ri;
      // Payments only ever go up, so the higher paid amount is the freshest state.
      return (li.paidAmount || 0) >= (ri.paidAmount || 0) ? li : ri;
    })
    .sort((a, b) => a.seq - b.seq);
  // Prefer whichever copy has the most payment activity recorded as the base
  // (covers non-schedule fields like status), then attach the merged schedule.
  const base = (localLoan.schedule || []).reduce((s, i) => s + (i.paidAmount || 0), 0)
    >= (remoteLoan.schedule || []).reduce((s, i) => s + (i.paidAmount || 0), 0)
    ? localLoan : remoteLoan;
  return { ...base, schedule };
}

function mergeData(remote, local) {
  if (!remote) return local;
  if (!local) return remote;

  const agents = { ...byId(remote.agents), ...byId(local.agents) };
  const clients = { ...byId(remote.clients), ...byId(local.clients) };
  const reinvestments = { ...byId(remote.reinvestments), ...byId(local.reinvestments) };

  const remoteLoans = byId(remote.loans);
  const localLoans = byId(local.loans);
  const loanIds = new Set([...Object.keys(remoteLoans), ...Object.keys(localLoans)]);
  const loans = [...loanIds].map((id) => mergeLoan(remoteLoans[id], localLoans[id]));

  return {
    agents: Object.values(agents),
    clients: Object.values(clients),
    loans,
    reinvestments: Object.values(reinvestments),
  };
}

/* ============================== UTILITIES ============================== */

const STORAGE_KEY = "loan-manager-data-v1";

const INR = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 });
const INR_COMPACT = new Intl.NumberFormat("en-IN", { notation: "compact", maximumFractionDigits: 1 });

function money(n) { return INR.format(Math.round(n || 0)); }
function moneyCompact(n) { return "₹" + INR_COMPACT.format(Math.round(n || 0)); }
function uid(prefix) { return prefix + "_" + Math.random().toString(36).slice(2, 9); }

function addDays(date, days) { const d = new Date(date); d.setDate(d.getDate() + days); return d; }
function addMonths(date, months) { const d = new Date(date); d.setMonth(d.getMonth() + months); return d; }
// For weekly loans, collection day is standardised to Sunday: given the
// disbursal date, returns the following Sunday (always at least 1 day out,
// even if disbursed on a Sunday, so the first instalment isn't due same-day).
function nextSunday(date) {
  const d = startOfDay(date);
  const day = d.getDay(); // 0 = Sunday
  const diff = day === 0 ? 7 : 7 - day;
  return addDays(d, diff);
}
function fmtDate(d) {
  const dt = new Date(d);
  return dt.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}
function fmtDateShort(d) {
  const dt = new Date(d);
  return dt.toLocaleDateString("en-IN", { day: "2-digit", month: "short" });
}
function startOfDay(d) { const nd = new Date(d); nd.setHours(0, 0, 0, 0); return nd; }
function getWeekStart(d) {
  const nd = startOfDay(d);
  const day = nd.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  nd.setDate(nd.getDate() + diff);
  return nd;
}

/* ---- Loan amortization (compound interest, reducing balance / EMI method) ---- */
function periodicRate(annualRatePct, frequency, customDays) {
  const r = annualRatePct / 100;
  if (frequency === "weekly") return r / 52;
  if (frequency === "monthly") return r / 12;
  return r * ((customDays || 30) / 365);
}

// Given a principal, a FIXED installment amount (chosen by the owner) and a
// number of installments, back-solve for the flat/simple periodic interest
// rate implied by those figures (total interest spread evenly, charged once
// on the original principal). Returns null if the fixed EMI is too small to
// ever repay the principal (i.e. no interest rate, however low, works).
function solvePeriodicRateFromEmi(principal, emi, n) {
  if (!(principal > 0) || !(emi > 0) || !(n > 0)) return null;
  const totalRepayment = emi * n;
  if (totalRepayment <= principal) return null;
  const totalInterest = totalRepayment - principal;
  return totalInterest / (principal * n); // simple rate per instalment period
}

// Converts a periodic rate back into the annual % this app stores on the
// loan record, so the fixed-EMI path can reuse generateSchedule() unchanged.
function periodicRateToAnnualPct(r, frequency, customDays) {
  if (frequency === "weekly") return r * 52 * 100;
  if (frequency === "monthly") return r * 12 * 100;
  return r * (365 / (customDays || 30)) * 100;
}

function generateSchedule(loan) {
  const { principal, annualRatePct, installments, frequency, customDays, startDate } = loan;
  const n = installments;
  const r = periodicRate(annualRatePct, frequency, customDays); // simple rate per instalment period
  // Simple/flat interest: total interest = principal x rate x number of periods,
  // charged once on the original principal - not on a reducing balance. Split
  // evenly across every instalment (last one absorbs any rounding).
  const totalInterest = principal * r * n;
  const emi = (principal + totalInterest) / n;
  const principalPerInst = principal / n;
  const interestPerInst = totalInterest / n;

  const schedule = [];
  const start = new Date(startDate);
  const firstWeeklyDue = frequency === "weekly" ? nextSunday(start) : null;
  let principalAllocated = 0, interestAllocated = 0;
  for (let i = 1; i <= n; i++) {
    let principalComp = principalPerInst;
    let interestComp = interestPerInst;
    if (i === n) {
      principalComp = principal - principalAllocated;
      interestComp = totalInterest - interestAllocated;
    }
    principalAllocated += principalComp;
    interestAllocated += interestComp;
    const due =
      frequency === "weekly" ? addDays(firstWeeklyDue, 7 * (i - 1))
      : frequency === "monthly" ? addMonths(start, i)
      : addDays(start, (customDays || 30) * i);
    schedule.push({
      id: uid("inst"),
      seq: i,
      dueDate: due.toISOString(),
      principalComponent: Math.round(principalComp * 100) / 100,
      interestComponent: Math.round(interestComp * 100) / 100,
      amount: Math.round((principalComp + interestComp) * 100) / 100,
      status: "pending",
      paidAmount: 0,
      paidDate: null,
    });
  }
  return { schedule, emi };
}

function getInstMeta(inst, today) {
  if (inst.status === "paid") return { key: "paid", label: "PAID" };
  if (inst.status === "written_off") return { key: "written_off", label: "WRITTEN OFF" };
  const due = new Date(inst.dueDate);
  const diffDays = Math.floor((due - today) / 86400000);
  if (diffDays < 0) return { key: "overdue", label: `OVERDUE ${Math.abs(diffDays)}D`, diffDays };
  if (inst.status === "partial") return { key: "partial", label: "PARTIAL", diffDays };
  if (diffDays <= 7) return { key: "duesoon", label: "DUE SOON", diffDays };
  return { key: "upcoming", label: "UPCOMING", diffDays };
}

function loanOutstanding(loan) {
  return loan.schedule.reduce((s, i) => s + Math.max(0, i.amount - (i.paidAmount || 0)), 0);
}
function loanIsClosed(loan) {
  return loan.schedule.every((i) => i.status === "paid" || i.status === "written_off");
}
function nextDueInstallment(loan, today) {
  return loan.schedule.find((i) => i.status !== "paid" && i.status !== "written_off");
}
// True if any installment on this loan was ever struck off as a bad debt
// (whole loan write-off, or written off because the client was closed).
function loanIsWrittenOff(loan) {
  return loan.schedule.some((i) => i.status === "written_off");
}
// Three-way status for display: an "Active" loan is still being collected;
// a "Closed" loan was fully repaid; a "Struck Off" loan was closed early
// via a bad-debt write-off. Lets the UI tell a normally-closed loan apart
// from one that was struck off, and both apart from a fresh loan on the
// same client.
function loanStatusMeta(loan) {
  if (!loanIsClosed(loan)) return { key: "active", label: "Active" };
  return loanIsWrittenOff(loan) ? { key: "struck_off", label: "Struck Off" } : { key: "closed", label: "Closed" };
}
// The date a closed/struck-off loan actually finished, for "as on <date>"
// labels - the latest paid/written-off date across its installments.
function loanClosedDate(loan) {
  if (!loanIsClosed(loan)) return null;
  const dates = loan.schedule.map((i) => i.paidDate).filter(Boolean).map((d) => new Date(d));
  if (!dates.length) return null;
  return new Date(Math.max(...dates.map((d) => d.getTime())));
}
// Total capital the owner has ever put out as loan principal - every loan
// ever issued, active, closed, or struck off, counts as money invested.
function totalPrincipalInvested(data) {
  return (data.loans || []).reduce((s, l) => s + (l.principal || 0), 0);
}
function totalBadDebt(data) {
  return (data.writeOffs || []).reduce((s, w) => s + (w.amount || 0), 0);
}
// Bad debt written off so far, grouped by client.
function badDebtByClient(data) {
  const map = {};
  (data.writeOffs || []).forEach((w) => {
    const loan = w.loanId ? data.loans.find((l) => l.id === w.loanId) : null;
    const clientId = w.clientId || loan?.clientId || null;
    const key = clientId || `name:${w.clientName || "Unknown"}`;
    if (!map[key]) map[key] = { clientId, name: w.clientName || "Unknown", amount: 0, count: 0 };
    map[key].amount += w.amount || 0;
    map[key].count += 1;
  });
  return Object.values(map).sort((a, b) => b.amount - a.amount);
}
// Bad debt written off so far, grouped by the agent whose client it was.
function badDebtByAgent(data) {
  const map = {};
  (data.writeOffs || []).forEach((w) => {
    const loan = w.loanId ? data.loans.find((l) => l.id === w.loanId) : null;
    const clientId = w.clientId || loan?.clientId || null;
    const client = clientId ? data.clients.find((c) => c.id === clientId) : null;
    const agent = client ? data.agents.find((a) => a.id === client.agentId) : null;
    const key = agent?.id || "unassigned";
    if (!map[key]) map[key] = { agentId: agent?.id || null, name: agent?.name || "Unassigned / closed client", amount: 0, count: 0 };
    map[key].amount += w.amount || 0;
    map[key].count += 1;
  });
  return Object.values(map).sort((a, b) => b.amount - a.amount);
}

function allRows(data) {
  const rows = [];
  data.loans.forEach((loan) => {
    const client = data.clients.find((c) => c.id === loan.clientId);
    const agent = client ? data.agents.find((a) => a.id === client.agentId) : null;
    loan.schedule.forEach((inst) => rows.push({ inst, loan, client, agent }));
  });
  return rows;
}

function getWeeklyStats(loans, today, weeksBack = 8) {
  const currentWeekStart = getWeekStart(today);
  const weeks = [];
  for (let w = weeksBack - 1; w >= 0; w--) {
    const start = addDays(currentWeekStart, -7 * w);
    const end = addDays(start, 6);
    let collected = 0, due = 0;
    loans.forEach((loan) =>
      loan.schedule.forEach((inst) => {
        const dueDate = new Date(inst.dueDate);
        if (dueDate >= start && dueDate <= end) due += inst.amount;
        if (inst.paidDate) {
          const pd = new Date(inst.paidDate);
          if (pd >= start && pd <= end) collected += inst.paidAmount || 0;
        }
      })
    );
    weeks.push({ label: `${start.getDate()}/${start.getMonth() + 1}`, collected: Math.round(collected), due: Math.round(due), start, end });
  }
  return weeks;
}

/* ---- Reminders (WhatsApp / SMS) ---- */
function phoneDigitsOf(phone) { return (phone || "").replace(/\D/g, ""); }
function intlPhoneOf(phone) {
  const d = phoneDigitsOf(phone);
  return d.length === 10 ? "91" + d : d;
}
function buildReminderMessage(client, inst) {
  const due = fmtDate(inst.dueDate);
  const amt = money(inst.amount - (inst.paidAmount || 0));
  return `Dear ${client?.name || "Customer"}, a reminder that your installment of ${amt} is due on ${due}. Kindly keep the amount ready for collection. - Annapurna Finance`;
}

/* ---- Full data backup (download + email) ---- */
function downloadBackupFile(data) {
  const filename = `annapurna-finance-backup-${new Date().toISOString().slice(0, 10)}.json`;
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return filename;
}

function openBackupEmailDraft(filename, toEmail) {
  const subject = encodeURIComponent(`Annapurna Finance Ledger backup - ${new Date().toLocaleDateString("en-IN")}`);
  const body = encodeURIComponent(
    `A backup file named "${filename}" has just been downloaded to this device's Downloads folder.\n\n` +
    `Please attach that file to this email before sending, so the data is safely stored in this mailbox.\n\n` +
    `(Mobile/browser security rules do not allow attaching files automatically - this one extra step keeps your data private.)`
  );
  const to = toEmail ? encodeURIComponent(toEmail) : "";
  window.location.href = `mailto:${to}?subject=${subject}&body=${body}`;
}

function parseBackupFile(text) {
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    throw new Error("That file isn't valid JSON. Please choose the backup .json file.");
  }
  if (!parsed || !Array.isArray(parsed.agents) || !Array.isArray(parsed.clients) || !Array.isArray(parsed.loans)) {
    throw new Error("That file doesn't look like an Annapurna Finance Ledger backup.");
  }
  return parsed;
}

/* ---- Excel export of the collection register ---- */
function exportCollectionRegister(rows, weeks) {
  const registerRows = rows.map((r) => ({
    Client: r.client?.name || "",
    Area: r.client?.area || "",
    Agent: r.agent?.name || "",
    "Loan ID": r.loan.id,
    Installment: r.inst.seq,
    "Due Date": fmtDate(r.inst.dueDate),
    Amount: r.inst.amount,
    "Principal Component": r.inst.principalComponent,
    "Interest Component": r.inst.interestComponent,
    "Paid Amount": r.inst.paidAmount || 0,
    "Paid Date": r.inst.paidDate ? fmtDate(r.inst.paidDate) : "",
    Status: r.inst.status,
  }));
  const weekRows = (weeks || []).map((w) => ({
    "Week Starting": fmtDate(w.start),
    "Week Ending": fmtDate(w.end),
    "Due": w.due,
    "Collected": w.collected,
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(registerRows), "Collection Register");
  if (weekRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(weekRows), "Weekly Summary");
  XLSX.writeFile(wb, `collection-register-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ---- Seed / demo data ---- */
function seedData() {
  const today = new Date();
  const agents = [
    { id: "ag1", name: "Rakesh Sahoo", phone: "9861023456", area: "Patia" },
    { id: "ag2", name: "Sunita Behera", phone: "9437087123", area: "Khandagiri" },
  ];
  const clients = [
    { id: "cl1", name: "Debasis Nayak", phone: "9861111111", address: "Plot 12, Patia", area: "Patia", agentId: "ag1" },
    { id: "cl2", name: "Manoj Pradhan", phone: "9861222222", address: "Near Big Bazar, Patia", area: "Patia", agentId: "ag1" },
    { id: "cl3", name: "Pramila Swain", phone: "9437333333", address: "Lane 4, Khandagiri", area: "Khandagiri", agentId: "ag2" },
    { id: "cl4", name: "Ashok Mallick", phone: "9437444444", address: "Near Temple, Khandagiri", area: "Khandagiri", agentId: "ag2" },
    { id: "cl5", name: "Snehalata Jena", phone: "9861555555", address: "Sector 6, Patia", area: "Patia", agentId: "ag1" },
  ];
  const loanDefs = [
    { id: "ln1", clientId: "cl1", principal: 20000, annualRatePct: 24, installments: 20, frequency: "weekly", startOffsetDays: -90 },
    { id: "ln2", clientId: "cl2", principal: 15000, annualRatePct: 22, installments: 10, frequency: "monthly", startOffsetDays: -150 },
    { id: "ln3", clientId: "cl3", principal: 30000, annualRatePct: 26, installments: 24, frequency: "weekly", startOffsetDays: -60 },
    { id: "ln4", clientId: "cl4", principal: 10000, annualRatePct: 20, installments: 6, frequency: "monthly", startOffsetDays: -120 },
    { id: "ln5", clientId: "cl5", principal: 25000, annualRatePct: 24, installments: 16, frequency: "weekly", startOffsetDays: -40 },
  ];
  const loans = loanDefs.map((def) => {
    const startDate = addDays(today, def.startOffsetDays).toISOString();
    const { schedule } = generateSchedule({ ...def, customDays: null, startDate });
    return { id: def.id, clientId: def.clientId, principal: def.principal, annualRatePct: def.annualRatePct, installments: def.installments, frequency: def.frequency, customDays: null, startDate, status: "active", schedule };
  });

  loans.forEach((loan) => {
    loan.schedule.forEach((inst) => {
      const due = new Date(inst.dueDate);
      const daysOverdue = (today - due) / 86400000;
      if (daysOverdue > 10) {
        inst.status = "paid";
        inst.paidAmount = inst.amount;
        inst.paidDate = addDays(due, Math.floor(Math.random() * 3)).toISOString();
      }
    });
  });
  const ln3 = loans.find((l) => l.id === "ln3");
  if (ln3) {
    const overdueInst = ln3.schedule.find((i) => {
      const d = (today - new Date(i.dueDate)) / 86400000;
      return d > 0 && d <= 10;
    });
    if (overdueInst) { overdueInst.status = "pending"; overdueInst.paidAmount = 0; overdueInst.paidDate = null; }
  }
  const ln1 = loans.find((l) => l.id === "ln1");
  if (ln1) {
    const recentPaid = [...ln1.schedule].reverse().find((i) => i.status === "paid");
    if (recentPaid) { recentPaid.status = "partial"; recentPaid.paidAmount = Math.round(recentPaid.amount * 0.6); }
  }

  return { agents, clients, loans, reinvestments: [] };
}

/* ============================== SMALL UI PIECES ============================== */

const FONT_STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Lora:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
.font-display { font-family: 'Lora', serif; }
.font-ledger { font-family: 'IBM Plex Mono', monospace; font-variant-numeric: tabular-nums; }
.app-root { font-family: 'Inter', ui-sans-serif, system-ui, sans-serif; }
.ledger-rule { background-image: repeating-linear-gradient(to bottom, transparent, transparent 39px, #e7e2d9 39px, #e7e2d9 40px); }
@media print {
  .no-print { display: none !important; }
  body, .app-root { background: white !important; }
  /* A fixed-position, scrolling overlay gets clipped to one viewport-height
     page by most browsers' print engines. Force it back to normal document
     flow only while printing, so every page of the passbook comes out. */
  .print-overlay {
    position: static !important;
    inset: auto !important;
    overflow: visible !important;
    height: auto !important;
  }
}
`;

const STAMP_STYLES = {
  paid: "text-emerald-700 border-emerald-600 bg-emerald-50",
  partial: "text-amber-700 border-amber-600 bg-amber-50",
  overdue: "text-rose-700 border-rose-600 bg-rose-50",
  duesoon: "text-amber-700 border-amber-500 bg-amber-50",
  upcoming: "text-stone-500 border-stone-400 bg-stone-50",
  written_off: "text-stone-500 border-stone-400 bg-stone-100 line-through",
};

function StatusStamp({ meta }) {
  return (
    <span className={`inline-block px-2 py-0.5 border-2 border-dashed rounded-sm font-ledger text-[10px] tracking-wider -rotate-2 whitespace-nowrap ${STAMP_STYLES[meta.key]}`}>
      {meta.label}
    </span>
  );
}

function StatCard({ icon: Icon, label, value, sub, tone = "slate" }) {
  const toneMap = {
    slate: "text-slate-900 bg-slate-900",
    emerald: "text-emerald-700 bg-emerald-600",
    amber: "text-amber-700 bg-amber-500",
    rose: "text-rose-700 bg-rose-600",
  };
  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 flex flex-col gap-2 shadow-sm">
      <div className="flex items-center justify-between">
        <span className="text-xs uppercase tracking-wider text-stone-500 font-medium">{label}</span>
        <span className={`w-7 h-7 rounded-full flex items-center justify-center ${toneMap[tone].split(" ")[1]} bg-opacity-10`}>
          <Icon className={`w-4 h-4 ${toneMap[tone].split(" ")[0]}`} />
        </span>
      </div>
      <div className="font-ledger text-2xl font-semibold text-stone-900">{value}</div>
      {sub && <div className="text-xs text-stone-500">{sub}</div>}
    </div>
  );
}

function SectionHeader({ title, subtitle, action }) {
  return (
    <div className="flex items-start justify-between mb-4 gap-3 flex-wrap">
      <div>
        <h2 className="font-display text-xl font-semibold text-stone-900">{title}</h2>
        {subtitle && <p className="text-sm text-stone-500 mt-0.5">{subtitle}</p>}
      </div>
      {action}
    </div>
  );
}

function Btn({ children, onClick, variant = "primary", size = "md", icon: Icon, type = "button", disabled }) {
  const base = "inline-flex items-center gap-1.5 rounded-md font-medium transition-colors disabled:opacity-40 disabled:cursor-not-allowed";
  const sizes = { md: "px-3.5 py-2 text-sm", sm: "px-2.5 py-1.5 text-xs" };
  const variants = {
    primary: "bg-slate-900 text-white hover:bg-slate-800",
    outline: "border border-stone-300 text-stone-700 hover:bg-stone-100 bg-white",
    ghost: "text-stone-600 hover:bg-stone-100",
    danger: "bg-rose-600 text-white hover:bg-rose-700",
    subtle: "bg-stone-100 text-stone-700 hover:bg-stone-200",
  };
  return (
    <button type={type} disabled={disabled} onClick={onClick} className={`${base} ${sizes[size]} ${variants[variant]}`}>
      {Icon && <Icon className="w-4 h-4" />}
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="text-stone-600 font-medium text-xs uppercase tracking-wide">{label}</span>
      {children}
    </label>
  );
}
const inputCls = "border border-stone-300 rounded-md px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-slate-900/20 focus:border-slate-400 bg-white";

/* Inline pay control used across owner + agent views */
function PayRow({ inst, today, onPay }) {
  const [open, setOpen] = useState(false);
  const [amount, setAmount] = useState(Math.max(0, inst.amount - (inst.paidAmount || 0)));
  const [collectedOn, setCollectedOn] = useState(startOfDay(today).toISOString().slice(0, 10));
  const meta = getInstMeta(inst, today);
  if (inst.status === "paid" || inst.status === "written_off") return <StatusStamp meta={meta} />;
  if (!open) {
    return (
      <div className="flex items-center gap-2">
        <StatusStamp meta={meta} />
        <button onClick={() => setOpen(true)} className="text-xs font-medium text-emerald-700 hover:text-emerald-800 underline decoration-dashed underline-offset-2">
          Collect
        </button>
      </div>
    );
  }
  return (
    <div className="flex items-center gap-1.5">
      <input
        type="date"
        value={collectedOn}
        onChange={(e) => setCollectedOn(e.target.value)}
        className={inputCls + " w-32"}
        title="Date of collection"
      />
      <input
        type="number"
        value={amount}
        onChange={(e) => setAmount(Number(e.target.value))}
        className={inputCls + " w-24"}
        min={0}
      />
      <button
        onClick={() => { onPay(inst.id, amount, new Date(collectedOn).toISOString()); setOpen(false); }}
        className="w-7 h-7 rounded-md bg-emerald-600 text-white flex items-center justify-center hover:bg-emerald-700"
        title="Confirm collection"
      >
        <Check className="w-4 h-4" />
      </button>
      <button onClick={() => setOpen(false)} className="w-7 h-7 rounded-md bg-stone-100 text-stone-500 flex items-center justify-center hover:bg-stone-200">
        <X className="w-4 h-4" />
      </button>
    </div>
  );
}

function ReminderButton({ client, inst }) {
  const [copied, setCopied] = useState(false);
  if (!client) return null;
  const message = buildReminderMessage(client, inst);
  const digits = phoneDigitsOf(client.phone);
  const intl = intlPhoneOf(client.phone);

  async function copyMsg() {
    try {
      await navigator.clipboard.writeText(message);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch (e) {
      setCopied("manual");
    }
  }

  return (
    <div className="relative inline-flex items-center gap-1">
      <a
        href={`https://wa.me/${intl}?text=${encodeURIComponent(message)}`}
        target="_blank" rel="noopener noreferrer" title="Send WhatsApp reminder"
        className="w-6 h-6 rounded flex items-center justify-center text-emerald-600 hover:bg-emerald-50"
      >
        <MessageCircle className="w-3.5 h-3.5" />
      </a>
      <a
        href={`sms:${digits}?body=${encodeURIComponent(message)}`} title="Send SMS reminder"
        className="w-6 h-6 rounded flex items-center justify-center text-sky-600 hover:bg-sky-50"
      >
        <Smartphone className="w-3.5 h-3.5" />
      </a>
      <button onClick={copyMsg} title="Copy reminder text" className="w-6 h-6 rounded flex items-center justify-center text-stone-500 hover:bg-stone-100">
        <Copy className="w-3.5 h-3.5" />
      </button>
      {copied && (
        <div className="absolute z-20 top-7 right-0 bg-white border border-stone-200 rounded-md shadow-md p-2 text-[11px] w-52 text-stone-600">
          {copied === "manual" ? (
            <>
              <div className="mb-1">Select &amp; copy manually:</div>
              <textarea readOnly value={message} rows={3} className="w-full text-[10px] border border-stone-200 rounded p-1" onClick={(e) => e.target.select()} />
            </>
          ) : "Copied to clipboard"}
        </div>
      )}
    </div>
  );
}

/* ============================== PASSBOOK (PRINT / PDF) ============================== */

function PassbookOverlay({ client, agent, loans, onClose }) {
  return (
    <div className="print-overlay fixed inset-0 z-50 bg-white overflow-y-auto app-root">
      <style>{FONT_STYLE}</style>
      <div className="no-print sticky top-0 bg-slate-900 text-white flex items-center justify-between px-4 py-3 gap-3 flex-wrap">
        <span className="font-display font-semibold">Passbook — {client.name}</span>
        <div className="flex items-center gap-2">
          <Btn variant="outline" icon={Printer} onClick={() => window.print()}>Print / Save as PDF</Btn>
          <button onClick={onClose} className="w-8 h-8 rounded-md bg-slate-800 flex items-center justify-center hover:bg-slate-700">
            <X className="w-4 h-4" />
          </button>
        </div>
      </div>
      <div className="max-w-3xl mx-auto p-6 sm:p-10">
        <div className="text-center mb-6 border-b-2 border-stone-800 pb-4">
          <div className="font-display text-2xl font-semibold">Annapurna Finance Ledger</div>
          <div className="text-xs text-stone-500 mt-1 tracking-wide uppercase">Customer Loan Passbook</div>
        </div>
        <div className="grid grid-cols-2 gap-2 mb-8 text-xs font-ledger">
          <div><span className="text-stone-400">Customer: </span><span className="font-medium">{client.name}</span></div>
          <div><span className="text-stone-400">Phone: </span>{client.phone || "—"}</div>
          <div><span className="text-stone-400">Address: </span>{client.address || "—"}</div>
          <div><span className="text-stone-400">Area: </span>{client.area}</div>
          <div><span className="text-stone-400">Agent: </span>{agent?.name || "—"}</div>
          <div><span className="text-stone-400">Printed: </span>{fmtDate(new Date())}</div>
        </div>
        {loans.length === 0 && <p className="text-sm text-stone-500">No loans on record for this customer.</p>}
        {loans.map((loan) => (
          <div key={loan.id} className="mb-8" style={{ breakInside: "avoid" }}>
            <div className="flex justify-between items-baseline border-b border-dashed border-stone-300 pb-1 mb-2">
              <span className="font-display font-semibold">Loan #{loan.id.slice(-5).toUpperCase()}</span>
              <span className="text-xs text-stone-500 font-ledger">{money(loan.principal)} · {loan.annualRatePct}% p.a. · {loan.installments} × {loan.frequency}</span>
            </div>
            <table className="w-full text-xs font-ledger">
              <thead>
                <tr className="text-left text-stone-400 border-b border-stone-200">
                  <th className="py-1 pr-2">#</th><th className="py-1 pr-2">Due</th><th className="py-1 pr-2">Amount</th>
                  <th className="py-1 pr-2">Paid</th><th className="py-1 pr-2">Paid On</th><th className="py-1 pr-2">Status</th>
                </tr>
              </thead>
              <tbody>
                {loan.schedule.map((inst) => (
                  <tr key={inst.id} className="border-b border-stone-100">
                    <td className="py-1 pr-2">{inst.seq}</td>
                    <td className="py-1 pr-2">{fmtDateShort(inst.dueDate)}</td>
                    <td className="py-1 pr-2">{money(inst.amount)}</td>
                    <td className="py-1 pr-2">{inst.paidAmount ? money(inst.paidAmount) : "—"}</td>
                    <td className="py-1 pr-2">{inst.paidDate ? fmtDateShort(inst.paidDate) : "—"}</td>
                    <td className="py-1 pr-2">{inst.status.toUpperCase()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="flex justify-end gap-6 text-xs mt-2 font-medium font-ledger">
              <span>Total paid: {money(loan.schedule.reduce((s, i) => s + (i.paidAmount || 0), 0))}</span>
              <span>Outstanding: {money(loanOutstanding(loan))}</span>
            </div>
          </div>
        ))}
        <div className="text-center text-[11px] text-stone-400 mt-10 border-t border-stone-200 pt-3">
          This passbook is a record of your loan account with Annapurna Finance. Please retain it for your reference.
        </div>
      </div>
    </div>
  );
}

/* ============================== LOGIN ============================== */

function LoginScreen({ data, onLogin }) {
  const [pickingAgent, setPickingAgent] = useState(false);
  return (
    <div className="min-h-screen bg-stone-50 flex items-center justify-center p-6 app-root">
      <style>{FONT_STYLE}</style>
      <div className="w-full max-w-md">
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-lg bg-slate-900 text-white mb-3">
            <Landmark className="w-6 h-6" />
          </div>
          <h1 className="font-display text-2xl font-semibold text-stone-900">Annapurna Finance Ledger</h1>
          <p className="text-sm text-stone-500 mt-1">Loan &amp; collection register</p>
        </div>

        {!pickingAgent ? (
          <div className="grid grid-cols-1 gap-3">
            <button
              onClick={() => onLogin({ role: "owner" })}
              className="bg-white border border-stone-200 rounded-lg p-5 text-left hover:border-slate-400 hover:shadow-sm transition-all flex items-center justify-between group"
            >
              <div>
                <div className="font-display text-lg font-semibold text-stone-900">Owner</div>
                <div className="text-sm text-stone-500">Full access — clients, loans, rates &amp; reinvestment</div>
              </div>
              <ChevronRight className="w-5 h-5 text-stone-400 group-hover:text-slate-900" />
            </button>
            <button
              onClick={() => setPickingAgent(true)}
              className="bg-white border border-stone-200 rounded-lg p-5 text-left hover:border-slate-400 hover:shadow-sm transition-all flex items-center justify-between group"
            >
              <div>
                <div className="font-display text-lg font-semibold text-stone-900">Field Agent</div>
                <div className="text-sm text-stone-500">Your customers, arrears &amp; collections only</div>
              </div>
              <ChevronRight className="w-5 h-5 text-stone-400 group-hover:text-slate-900" />
            </button>
            <p className="text-[11px] text-stone-400 text-center mt-2">This is a role selector for the app, not a secured login.</p>
          </div>
        ) : (
          <div className="bg-white border border-stone-200 rounded-lg p-5">
            <button onClick={() => setPickingAgent(false)} className="flex items-center gap-1 text-xs text-stone-500 hover:text-stone-800 mb-3">
              <ArrowLeft className="w-3.5 h-3.5" /> Back
            </button>
            <div className="text-sm font-medium text-stone-600 mb-3">Select your name</div>
            <div className="flex flex-col gap-2">
              {data.agents.map((a) => (
                <button
                  key={a.id}
                  onClick={() => onLogin({ role: "agent", agentId: a.id })}
                  className="border border-stone-200 rounded-md px-3 py-2.5 text-left hover:border-slate-400 hover:bg-stone-50 flex items-center justify-between"
                >
                  <span>
                    <span className="font-medium text-stone-900">{a.name}</span>
                    <span className="text-stone-400 text-xs ml-2">{a.area}</span>
                  </span>
                  <ChevronRight className="w-4 h-4 text-stone-400" />
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================== NAV ============================== */

const OWNER_NAV = [
  { key: "dashboard", label: "Dashboard", icon: Home },
  { key: "clients", label: "Clients", icon: Users },
  { key: "agents", label: "Agents", icon: UserRound },
  { key: "loans", label: "Loans", icon: Landmark },
  { key: "collections", label: "Collections", icon: CalendarClock },
  { key: "reinvest", label: "Reinvest", icon: TrendingUp },
];
const AGENT_NAV = [
  { key: "customers", label: "My Customers", icon: Users },
  { key: "arrears", label: "Arrears", icon: AlertTriangle },
  { key: "duesoon", label: "Due Soon", icon: CalendarClock },
];

function Sidebar({ role, tab, setTab, onLogout, agentLabel, onResetDemo, onEraseAll, data, onRestoreBackup, shared, lastSynced, syncing, onSyncNow }) {
  const items = role === "owner" ? OWNER_NAV : AGENT_NAV;
  const [confirmReset, setConfirmReset] = useState(false);
  const [confirmErase, setConfirmErase] = useState(false);
  const [confirmRestore, setConfirmRestore] = useState(false);
  const [backingUp, setBackingUp] = useState(false);
  const [restoreMsg, setRestoreMsg] = useState(null); // { type: "ok"|"error", text }
  const fileInputRef = useRef(null);

  function handleEmailBackup() {
    setBackingUp(true);
    const filename = downloadBackupFile(data);
    setTimeout(() => {
      openBackupEmailDraft(filename);
      setBackingUp(false);
    }, 400);
  }

  function handleRestoreClick() {
    setConfirmRestore(false);
    fileInputRef.current?.click();
  }

  function handleFileChosen(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = ""; // allow choosing the same file again later
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const restored = parseBackupFile(reader.result);
        onRestoreBackup(restored);
        setRestoreMsg({ type: "ok", text: "Backup restored." });
      } catch (err) {
        setRestoreMsg({ type: "error", text: err.message });
      }
      setTimeout(() => setRestoreMsg(null), 4000);
    };
    reader.onerror = () => {
      setRestoreMsg({ type: "error", text: "Could not read that file." });
      setTimeout(() => setRestoreMsg(null), 4000);
    };
    reader.readAsText(file);
  }
  return (
    <div className="w-16 sm:w-56 shrink-0 bg-slate-900 text-slate-200 flex flex-col justify-between min-h-screen">
      <div>
        <div className="h-16 flex items-center gap-2 px-3 sm:px-4 border-b border-slate-700/60">
          <Landmark className="w-5 h-5 text-amber-400 shrink-0" />
          <span className="hidden sm:block font-display text-sm font-semibold text-white leading-tight">Annapurna Finance<br /><span className="text-slate-400 font-sans text-[11px] font-normal">Ledger</span></span>
        </div>
        <nav className="p-2 flex flex-col gap-1 mt-2">
          {items.map((it) => {
            const Icon = it.icon;
            const active = tab === it.key;
            return (
              <button
                key={it.key}
                onClick={() => setTab(it.key)}
                className={`flex items-center gap-3 px-3 py-2.5 rounded-md text-sm transition-colors ${active ? "bg-white text-slate-900 font-medium" : "text-slate-300 hover:bg-slate-800"}`}
              >
                <Icon className="w-4 h-4 shrink-0" />
                <span className="hidden sm:block">{it.label}</span>
              </button>
            );
          })}
        </nav>
      </div>
      <div className="p-2 border-t border-slate-700/60 flex flex-col gap-1">
        <div className="flex items-center justify-between px-2 py-1">
          <span className="hidden sm:flex items-center gap-1.5 text-[11px] text-slate-400">
            <span className={`w-1.5 h-1.5 rounded-full ${shared ? "bg-emerald-400" : "bg-amber-400"}`} />
            {shared ? (lastSynced ? `Synced ${lastSynced.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })}` : "Connecting…") : "Not shared (this device only)"}
          </span>
          {shared && (
            <button onClick={onSyncNow} disabled={syncing} title="Sync now" className="w-6 h-6 rounded flex items-center justify-center text-slate-400 hover:bg-slate-800 hover:text-slate-200 disabled:opacity-50">
              <RefreshCw className={`w-3.5 h-3.5 ${syncing ? "animate-spin" : ""}`} />
            </button>
          )}
        </div>
        <button onClick={handleEmailBackup} disabled={backingUp} className="flex items-center gap-3 px-3 py-2 rounded-md text-xs text-slate-400 hover:bg-slate-800 hover:text-slate-200 disabled:opacity-50">
          <Mail className="w-3.5 h-3.5 shrink-0" />
          <span className="hidden sm:block">{backingUp ? "Preparing backup…" : "Backup by email"}</span>
        </button>

        <input ref={fileInputRef} type="file" accept="application/json,.json" onChange={handleFileChosen} className="hidden" />
        {confirmRestore ? (
          <div className="px-1 py-1 flex flex-col gap-1">
            <span className="hidden sm:block text-[11px] text-slate-400 px-2">Replace current data with the chosen backup file?</span>
            <div className="flex gap-1 px-1">
              <button onClick={handleRestoreClick} className="flex-1 text-xs bg-emerald-600 text-white rounded px-2 py-1.5">Choose file</button>
              <button onClick={() => setConfirmRestore(false)} className="flex-1 text-xs bg-slate-700 text-white rounded px-2 py-1.5">Cancel</button>
            </div>
          </div>
        ) : (
          <button onClick={() => setConfirmRestore(true)} className="flex items-center gap-3 px-3 py-2 rounded-md text-xs text-slate-400 hover:bg-slate-800 hover:text-slate-200">
            <Upload className="w-3.5 h-3.5 shrink-0" />
            <span className="hidden sm:block">Restore backup</span>
          </button>
        )}
        {restoreMsg && (
          <div className={`px-2 py-1 text-[11px] rounded ${restoreMsg.type === "ok" ? "text-emerald-300" : "text-rose-300"} hidden sm:block`}>
            {restoreMsg.text}
          </div>
        )}

        {role === "owner" && (
          confirmReset ? (
            <div className="px-1 py-1 flex flex-col gap-1">
              <span className="hidden sm:block text-[11px] text-slate-400 px-2">Reset all demo data?</span>
              <div className="flex gap-1 px-1">
                <button onClick={() => { onResetDemo(); setConfirmReset(false); }} className="flex-1 text-xs bg-rose-600 text-white rounded px-2 py-1.5">Yes</button>
                <button onClick={() => setConfirmReset(false)} className="flex-1 text-xs bg-slate-700 text-white rounded px-2 py-1.5">No</button>
              </div>
            </div>
          ) : (
            <button onClick={() => setConfirmReset(true)} className="flex items-center gap-3 px-3 py-2 rounded-md text-xs text-slate-400 hover:bg-slate-800 hover:text-slate-200">
              <RotateCcw className="w-3.5 h-3.5 shrink-0" />
              <span className="hidden sm:block">Reset demo data</span>
            </button>
          )
        )}
        {role === "owner" && (
          confirmErase ? (
            <div className="px-1 py-1 flex flex-col gap-1">
              <span className="hidden sm:block text-[11px] text-slate-400 px-2">Erase every agent, client, loan &amp; write-off? This cannot be undone — take a backup first if unsure.</span>
              <div className="flex gap-1 px-1">
                <button onClick={() => { onEraseAll(); setConfirmErase(false); }} className="flex-1 text-xs bg-rose-600 text-white rounded px-2 py-1.5">Yes, erase</button>
                <button onClick={() => setConfirmErase(false)} className="flex-1 text-xs bg-slate-700 text-white rounded px-2 py-1.5">No</button>
              </div>
            </div>
          ) : (
            <button onClick={() => setConfirmErase(true)} className="flex items-center gap-3 px-3 py-2 rounded-md text-xs text-slate-400 hover:bg-slate-800 hover:text-slate-200">
              <Trash2 className="w-3.5 h-3.5 shrink-0" />
              <span className="hidden sm:block">Erase all data</span>
            </button>
          )
        )}
        <button onClick={onLogout} className="flex items-center gap-3 px-3 py-2.5 rounded-md text-sm text-slate-300 hover:bg-slate-800">
          <LogOut className="w-4 h-4 shrink-0" />
          <span className="hidden sm:block">{agentLabel ? `Switch (${agentLabel})` : "Switch role"}</span>
        </button>
      </div>
    </div>
  );
}

/* ============================== OWNER: DASHBOARD ============================== */

function OwnerDashboard({ data, today, onPay }) {
  const weeks = useMemo(() => getWeeklyStats(data.loans, today, 8), [data.loans, today]);
  const thisWeek = weeks[weeks.length - 1];
  const lastWeek = weeks[weeks.length - 2];
  const rows = useMemo(() => allRows(data), [data]);

  const totalOutstanding = data.loans.reduce((s, l) => s + loanOutstanding(l), 0);
  const activeLoans = data.loans.filter((l) => !loanIsClosed(l)).length;

  const overdueRows = rows.filter((r) => r.inst.status !== "paid" && new Date(r.inst.dueDate) < today);
  const overdueTotal = overdueRows.reduce((s, r) => s + (r.inst.amount - (r.inst.paidAmount || 0)), 0);

  const dueSoonRows = rows
    .filter((r) => {
      const diff = Math.floor((new Date(r.inst.dueDate) - today) / 86400000);
      return r.inst.status !== "paid" && diff >= 0 && diff <= 7;
    })
    .sort((a, b) => new Date(a.inst.dueDate) - new Date(b.inst.dueDate));

  const dueSoonTotal = dueSoonRows.reduce((s, r) => s + (r.inst.amount - (r.inst.paidAmount || 0)), 0);

  const totalInvested = totalPrincipalInvested(data);
  const badDebt = totalBadDebt(data);
  const badDebtPct = totalInvested > 0 ? (badDebt / totalInvested) * 100 : 0;
  const byClient = useMemo(() => badDebtByClient(data), [data]);
  const byAgent = useMemo(() => badDebtByAgent(data), [data]);

  return (
    <div>
      <SectionHeader title="Dashboard" subtitle={`As of ${fmtDate(today)}`} />
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-6">
        <StatCard icon={Landmark} label="Active Loans" value={activeLoans} sub={`${data.loans.length} total issued`} tone="slate" />
        <StatCard icon={IndianRupee} label="Outstanding" value={moneyCompact(totalOutstanding)} sub="Principal + interest receivable" tone="slate" />
        <StatCard icon={Wallet} label="This Week Collected" value={moneyCompact(thisWeek.collected)} sub={`vs ${moneyCompact(lastWeek.collected)} last week`} tone="emerald" />
        <StatCard icon={AlertTriangle} label="In Arrears" value={moneyCompact(overdueTotal)} sub={`${overdueRows.length} installment(s) overdue`} tone="rose" />
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 mb-6">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <h3 className="font-display text-base font-semibold text-stone-900">Capital &amp; bad debt overview</h3>
          <span className="text-xs text-stone-400">All-time, across every loan ever issued</span>
        </div>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
          <StatCard icon={PiggyBank} label="Total Invested by Owner" value={moneyCompact(totalInvested)} sub={`${data.loans.length} loan(s) disbursed`} tone="slate" />
          <StatCard icon={AlertTriangle} label="Total Bad Debt" value={moneyCompact(badDebt)} sub={`${(data.writeOffs || []).length} write-off record(s)`} tone="rose" />
          <StatCard icon={TrendingUp} label="Bad Debt % of Invested" value={`${badDebtPct.toFixed(2)}%`} sub="Bad debt ÷ total invested" tone="rose" />
          <StatCard icon={Wallet} label="Net Collected So Far" value={moneyCompact(rows.reduce((s, r) => s + (r.inst.paidAmount || 0), 0))} sub="All installments paid to date" tone="emerald" />
        </div>

        {byClient.length === 0 ? (
          <p className="text-sm text-stone-500 py-2">No bad debt recorded yet.</p>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <div>
              <h4 className="text-xs uppercase tracking-wide text-stone-400 font-medium mb-2">Bad debt — client-wise</h4>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-xs uppercase tracking-wide text-stone-400 border-b border-stone-200">
                      <th className="py-1.5 pr-3">Client</th><th className="py-1.5 pr-3">Records</th><th className="py-1.5 pr-3">Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {byClient.map((r) => (
                      <tr key={r.clientId || r.name} className="border-b border-stone-100 last:border-0">
                        <td className="py-1.5 pr-3 text-stone-700">{r.name}</td>
                        <td className="py-1.5 pr-3 text-stone-500">{r.count}</td>
                        <td className="py-1.5 pr-3 font-ledger text-rose-700">{money(r.amount)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div>
              <h4 className="text-xs uppercase tracking-wide text-stone-400 font-medium mb-2">Bad debt — agent-wise</h4>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-xs uppercase tracking-wide text-stone-400 border-b border-stone-200">
                      <th className="py-1.5 pr-3">Agent</th><th className="py-1.5 pr-3">Records</th><th className="py-1.5 pr-3">Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {byAgent.map((r) => (
                      <tr key={r.agentId || r.name} className="border-b border-stone-100 last:border-0">
                        <td className="py-1.5 pr-3 text-stone-700">{r.name}</td>
                        <td className="py-1.5 pr-3 text-stone-500">{r.count}</td>
                        <td className="py-1.5 pr-3 font-ledger text-rose-700">{money(r.amount)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 mb-6">
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-display text-base font-semibold text-stone-900">Weekly collection — due vs collected</h3>
        </div>
        <ResponsiveContainer width="100%" height={240}>
          <BarChart data={weeks} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e7e2d9" vertical={false} />
            <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#78716c" }} axisLine={{ stroke: "#e7e2d9" }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: "#78716c" }} axisLine={false} tickLine={false} tickFormatter={(v) => moneyCompact(v)} width={55} />
            <Tooltip formatter={(v) => money(v)} contentStyle={{ fontSize: 12, borderRadius: 8, borderColor: "#e7e2d9" }} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="due" name="Due" fill="#d6d3cc" radius={[3, 3, 0, 0]} />
            <Bar dataKey="collected" name="Collected" fill="#059669" radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <h3 className="font-display text-base font-semibold text-stone-900">Due before due date (next 7 days)</h3>
          <span className="font-ledger text-sm text-stone-600">{moneyCompact(dueSoonTotal)} expected</span>
        </div>
        {dueSoonRows.length === 0 ? (
          <p className="text-sm text-stone-500 py-4 text-center">Nothing due in the next 7 days.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs uppercase tracking-wide text-stone-400 border-b border-stone-200">
                  <th className="py-2 pr-3">Customer</th>
                  <th className="py-2 pr-3">Agent</th>
                  <th className="py-2 pr-3">Area</th>
                  <th className="py-2 pr-3">Due date</th>
                  <th className="py-2 pr-3">Amount</th>
                  <th className="py-2 pr-3">Status</th>
                  <th className="py-2 pr-3">Remind</th>
                </tr>
              </thead>
              <tbody>
                {dueSoonRows.slice(0, 12).map((r) => (
                  <tr key={r.inst.id} className="border-b border-stone-100 last:border-0">
                    <td className="py-2 pr-3 font-medium text-stone-800">{r.client?.name}</td>
                    <td className="py-2 pr-3 text-stone-600">{r.agent?.name}</td>
                    <td className="py-2 pr-3 text-stone-500">{r.client?.area}</td>
                    <td className="py-2 pr-3 text-stone-600">{fmtDateShort(r.inst.dueDate)}</td>
                    <td className="py-2 pr-3 font-ledger">{money(r.inst.amount - (r.inst.paidAmount || 0))}</td>
                    <td className="py-2 pr-3"><PayRow inst={r.inst} today={today} onPay={(id, amt, dt) => onPay(r.loan.id, id, amt, dt)} /></td>
                    <td className="py-2 pr-3"><ReminderButton client={r.client} inst={r.inst} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================== OWNER: CLIENTS ============================== */

function ClientsPage({ data, actions }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const [form, setForm] = useState({ name: "", phone: "", address: "", area: "", agentId: data.agents[0]?.id || "" });
  const [passbookClientId, setPassbookClientId] = useState(null);

  const filtered = data.clients
    .filter((c) => c.name.toLowerCase().includes(q.toLowerCase()) || c.area.toLowerCase().includes(q.toLowerCase()))
    // Active clients first, closed/struck-off ones trail behind so the
    // working list stays clean while history is still just a scroll away.
    .sort((a, b) => (a.closed === b.closed ? 0 : a.closed ? 1 : -1));

  const activeCount = data.clients.filter((c) => !c.closed).length;
  const closedCount = data.clients.length - activeCount;

  function submit() {
    if (!form.name.trim() || !form.agentId) return;
    actions.addClient({ id: uid("cl"), ...form });
    setForm({ name: "", phone: "", address: "", area: "", agentId: data.agents[0]?.id || "" });
    setOpen(false);
  }

  return (
    <div>
      <SectionHeader
        title="Clients"
        subtitle={`${activeCount} active${closedCount ? ` · ${closedCount} closed/struck off` : ""} on the register`}
        action={<Btn icon={Plus} onClick={() => setOpen((v) => !v)}>{open ? "Close" : "Add Client"}</Btn>}
      />

      {(data.writeOffs || []).length > 0 && (
        <div className="bg-rose-50 border border-rose-200 rounded-lg px-4 py-2.5 mb-4 flex items-center justify-between text-sm">
          <span className="text-rose-700">Total written off to date, across {(data.writeOffs || []).length} record(s)</span>
          <span className="font-ledger font-semibold text-rose-700">{money(totalBadDebt(data))}</span>
        </div>
      )}

      {open && (
        <div className="bg-white border border-stone-200 rounded-lg p-4 mb-4 grid grid-cols-1 sm:grid-cols-2 gap-3">
          <Field label="Full name"><input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
          <Field label="Phone"><input className={inputCls} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label="Address"><input className={inputCls} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></Field>
          <Field label="Area"><input className={inputCls} value={form.area} onChange={(e) => setForm({ ...form, area: e.target.value })} placeholder="e.g. Patia" /></Field>
          <Field label="Assign agent">
            <select className={inputCls} value={form.agentId} onChange={(e) => setForm({ ...form, agentId: e.target.value })}>
              {data.agents.map((a) => <option key={a.id} value={a.id}>{a.name} — {a.area}</option>)}
            </select>
          </Field>
          <div className="flex items-end"><Btn onClick={submit}>Save client</Btn></div>
        </div>
      )}

      <div className="bg-white border border-stone-200 rounded-lg p-4">
        <div className="flex items-center gap-2 mb-3 border border-stone-200 rounded-md px-2.5 py-1.5 bg-stone-50 max-w-xs">
          <Search className="w-4 h-4 text-stone-400" />
          <input className="bg-transparent outline-none text-sm flex-1" placeholder="Search name or area" value={q} onChange={(e) => setQ(e.target.value)} />
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs uppercase tracking-wide text-stone-400 border-b border-stone-200">
                <th className="py-2 pr-3">Name</th><th className="py-2 pr-3">Area</th><th className="py-2 pr-3">Agent</th>
                <th className="py-2 pr-3">Loans</th><th className="py-2 pr-3">Outstanding</th><th className="py-2 pr-3"></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((c) => {
                const loans = data.loans.filter((l) => l.clientId === c.id);
                const outstanding = loans.reduce((s, l) => s + loanOutstanding(l), 0);
                const agent = data.agents.find((a) => a.id === c.agentId);
                const closed = !!c.closed;
                return (
                  <tr key={c.id} className={`border-b border-stone-100 last:border-0 ${closed ? "opacity-60" : ""}`}>
                    <td className="py-2.5 pr-3">
                      <div className={`font-medium text-stone-800 ${closed ? "line-through decoration-stone-400" : ""}`}>{c.name}</div>
                      <div className="text-xs text-stone-400 flex items-center gap-1"><Phone className="w-3 h-3" />{c.phone || "—"}</div>
                      {closed && (
                        <div className="text-[10px] uppercase tracking-wide text-rose-600 mt-0.5">Closed as on {fmtDateShort(c.closedAt || Date.now())}</div>
                      )}
                    </td>
                    <td className="py-2.5 pr-3 text-stone-600"><span className="flex items-center gap-1"><MapPin className="w-3.5 h-3.5 text-stone-400" />{c.area}</span></td>
                    <td className="py-2.5 pr-3 text-stone-600">{agent?.name}</td>
                    <td className="py-2.5 pr-3 text-stone-600">{loans.length}</td>
                    <td className="py-2.5 pr-3 font-ledger">{money(outstanding)}</td>
                    <td className="py-2.5 pr-3">
                      <div className="flex items-center gap-1.5">
                        <button onClick={() => setPassbookClientId(c.id)} title="View / print passbook" className="w-7 h-7 rounded-md border border-stone-200 flex items-center justify-center text-stone-500 hover:bg-stone-100">
                          <Printer className="w-3.5 h-3.5" />
                        </button>
                        {closed ? (
                          <button
                            onClick={() => { if (window.confirm(`Reopen ${c.name}? They will show as an active client again.`)) actions.reopenClient(c.id); }}
                            title="Reopen client"
                            className="w-7 h-7 rounded-md border border-stone-200 flex items-center justify-center text-stone-400 hover:text-emerald-600 hover:border-emerald-200"
                          >
                            <RotateCcw className="w-3.5 h-3.5" />
                          </button>
                        ) : (
                          <button
                            onClick={() => {
                              const msg = outstanding > 0
                                ? `${c.name} still has ${money(outstanding)} outstanding across ${loans.length} loan(s).\n\nClose this client? The remaining balance will be written off as bad debt, and the client/loan history will be kept on record as closed/struck off as on today.`
                                : `Close client ${c.name}? Their record and loan history is kept on file, marked closed as on today.`;
                              if (window.confirm(msg)) actions.closeClient(c.id);
                            }}
                            title="Close client (owner only) — record is kept, marked struck off as on today"
                            className="w-7 h-7 rounded-md border border-stone-200 flex items-center justify-center text-stone-400 hover:text-rose-600 hover:border-rose-200"
                          >
                            <Archive className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      {passbookClientId && (() => {
        const pbClient = data.clients.find((c) => c.id === passbookClientId);
        const pbAgent = data.agents.find((a) => a.id === pbClient?.agentId);
        const pbLoans = data.loans.filter((l) => l.clientId === passbookClientId);
        return <PassbookOverlay client={pbClient} agent={pbAgent} loans={pbLoans} onClose={() => setPassbookClientId(null)} />;
      })()}
    </div>
  );
}

/* ============================== OWNER: AGENTS ============================== */

function AgentsPage({ data, today, actions }) {
  const [open, setOpen] = useState(false);
  const [expanded, setExpanded] = useState(null);
  const [form, setForm] = useState({ name: "", phone: "", area: "" });

  function submit() {
    if (!form.name.trim()) return;
    actions.addAgent({ id: uid("ag"), ...form });
    setForm({ name: "", phone: "", area: "" });
    setOpen(false);
  }

  function remove(agent) {
    const activeClientCount = data.clients.filter((c) => c.agentId === agent.id && !c.closed).length;
    if (activeClientCount > 0) {
      alert(`${agent.name} still has ${activeClientCount} active client(s) assigned. Reassign or close those clients first.`);
      return;
    }
    if (window.confirm(`Close agent ${agent.name}? Their record and collection history is kept on file, marked closed as on today.`)) actions.closeAgent(agent.id);
  }

  const activeAgents = data.agents.filter((a) => !a.closed).length;
  const closedAgents = data.agents.length - activeAgents;

  return (
    <div>
      <SectionHeader
        title="Agents"
        subtitle={`${activeAgents} active${closedAgents ? ` · ${closedAgents} closed` : ""} field agents — additions and closures are owner-only`}
        action={<Btn icon={Plus} onClick={() => setOpen((v) => !v)}>{open ? "Close" : "Add Agent"}</Btn>}
      />
      {open && (
        <div className="bg-white border border-stone-200 rounded-lg p-4 mb-4 grid grid-cols-1 sm:grid-cols-3 gap-3">
          <Field label="Full name"><input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
          <Field label="Phone"><input className={inputCls} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label="Area"><input className={inputCls} value={form.area} onChange={(e) => setForm({ ...form, area: e.target.value })} /></Field>
          <div className="flex items-end sm:col-span-3"><Btn onClick={submit}>Save agent</Btn></div>
        </div>
      )}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        {[...data.agents].sort((a, b) => (a.closed === b.closed ? 0 : a.closed ? 1 : -1)).map((a) => {
          const clients = data.clients.filter((c) => c.agentId === a.id);
          const loans = data.loans.filter((l) => clients.some((c) => c.id === l.clientId));
          const outstanding = loans.reduce((s, l) => s + loanOutstanding(l), 0);
          const collectedSoFar = loans.reduce((s, l) => s + l.schedule.reduce((ss, i) => ss + (i.paidAmount || 0), 0), 0);
          const pending = loans.flatMap((l) => l.schedule.filter((i) => i.status !== "paid"));
          let nextDueDate = null, nextDueAmount = 0;
          if (pending.length) {
            nextDueDate = pending.reduce((min, i) => (!min || new Date(i.dueDate) < new Date(min) ? i.dueDate : min), null);
            nextDueAmount = pending
              .filter((i) => new Date(i.dueDate).toDateString() === new Date(nextDueDate).toDateString())
              .reduce((s, i) => s + (i.amount - (i.paidAmount || 0)), 0);
          }
          const isExpanded = expanded === a.id;
          const closed = !!a.closed;
          const agentBadDebt = badDebtByAgent(data).find((r) => r.agentId === a.id)?.amount || 0;
          return (
            <div key={a.id} className={`bg-white border border-stone-200 rounded-lg p-4 ${closed ? "opacity-60" : ""}`}>
              <div className="flex items-center justify-between mb-2">
                <div>
                  <div className={`font-display font-semibold text-stone-900 ${closed ? "line-through decoration-stone-400" : ""}`}>{a.name}</div>
                  <div className="text-xs text-stone-500 flex items-center gap-1 mt-0.5"><MapPin className="w-3 h-3" />{a.area} · {a.phone}</div>
                  {closed && <div className="text-[10px] uppercase tracking-wide text-rose-600 mt-0.5">Closed as on {fmtDateShort(a.closedAt || Date.now())}</div>}
                </div>
                <div className="flex items-center gap-2">
                  <span className="w-9 h-9 rounded-full bg-slate-900 text-white flex items-center justify-center font-display text-sm">{a.name.split(" ").map((s) => s[0]).slice(0, 2).join("")}</span>
                  {closed ? (
                    <button onClick={() => { if (window.confirm(`Reopen ${a.name}? They will show as an active agent again.`)) actions.reopenAgent(a.id); }} title="Reopen agent" className="w-7 h-7 rounded-md border border-stone-200 flex items-center justify-center text-stone-400 hover:text-emerald-600 hover:border-emerald-200">
                      <RotateCcw className="w-3.5 h-3.5" />
                    </button>
                  ) : (
                    <button onClick={() => remove(a)} title="Close agent (owner only) — record is kept, marked closed as on today" className="w-7 h-7 rounded-md border border-stone-200 flex items-center justify-center text-stone-400 hover:text-rose-600 hover:border-rose-200">
                      <Archive className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 mt-3 text-center">
                <div className="bg-stone-50 rounded-md py-2"><div className="font-ledger text-sm font-semibold">{clients.length}</div><div className="text-[10px] uppercase text-stone-400">Clients</div></div>
                <div className="bg-stone-50 rounded-md py-2"><div className="font-ledger text-sm font-semibold">{moneyCompact(outstanding)}</div><div className="text-[10px] uppercase text-stone-400">Outstanding</div></div>
                <div className="bg-stone-50 rounded-md py-2"><div className="font-ledger text-sm font-semibold text-emerald-700">{moneyCompact(collectedSoFar)}</div><div className="text-[10px] uppercase text-stone-400">Collected so far</div></div>
                <div className="bg-stone-50 rounded-md py-2"><div className="font-ledger text-sm font-semibold text-amber-700">{nextDueDate ? moneyCompact(nextDueAmount) : "—"}</div><div className="text-[10px] uppercase text-stone-400">{nextDueDate ? `Due ${fmtDateShort(nextDueDate)}` : "Nothing due"}</div></div>
                <div className="bg-stone-50 rounded-md py-2"><div className="font-ledger text-sm font-semibold text-rose-700">{agentBadDebt ? moneyCompact(agentBadDebt) : "—"}</div><div className="text-[10px] uppercase text-stone-400">Bad debt</div></div>
              </div>
              <button onClick={() => setExpanded(isExpanded ? null : a.id)} className="text-xs text-stone-500 hover:text-stone-800 mt-3 flex items-center gap-1">
                {isExpanded ? "Hide client details" : "Show client details"}
              </button>
              {isExpanded && (
                <div className="mt-2 border-t border-stone-100 pt-2 flex flex-col divide-y divide-stone-100">
                  {clients.length === 0 ? <p className="text-xs text-stone-400 py-2">No clients assigned.</p> : clients.map((c) => {
                    const cLoans = data.loans.filter((l) => l.clientId === c.id);
                    const cOutstanding = cLoans.reduce((s, l) => s + loanOutstanding(l), 0);
                    return (
                      <div key={c.id} className="py-2 flex items-center justify-between gap-2 text-sm">
                        <div>
                          <div className="text-stone-800">{c.name}</div>
                          <div className="text-xs text-stone-400">{c.phone} · {c.area}</div>
                        </div>
                        <div className="text-right">
                          <div className="font-ledger text-stone-700">{money(cOutstanding)}</div>
                          <div className="text-[10px] uppercase text-stone-400">{cLoans.length} loan(s)</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ============================== OWNER: LOANS ============================== */

function LoanDetail({ loan, client, today, onPay, onClose, onWriteOff }) {
  const outstanding = loanOutstanding(loan);
  const closed = loanIsClosed(loan);
  const meta = loanStatusMeta(loan);
  const closedDate = loanClosedDate(loan);
  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <button onClick={onClose} className="flex items-center gap-1 text-xs text-stone-500 hover:text-stone-800 mb-3">
        <ArrowLeft className="w-3.5 h-3.5" /> Back to loans
      </button>
      <div className="flex flex-wrap items-start justify-between gap-3 mb-4">
        <div>
          <h3 className={`font-display text-lg font-semibold text-stone-900 ${closed ? "line-through decoration-stone-400" : ""}`}>{client?.name}</h3>
          <p className="text-xs text-stone-500">{client?.area} · Started {fmtDate(loan.startDate)} · {loan.annualRatePct.toFixed(2)}% p.a. compounding, {loan.frequency}</p>
          <p className="text-xs mt-1">
            {meta.key === "active" && <span className="text-slate-700 font-medium">Active</span>}
            {meta.key === "closed" && <span className="text-emerald-700 font-medium">Closed{closedDate ? ` as on ${fmtDate(closedDate)}` : ""}</span>}
            {meta.key === "struck_off" && <span className="text-rose-700 font-medium">Struck Off (bad debt){closedDate ? ` as on ${fmtDate(closedDate)}` : ""}</span>}
          </p>
        </div>
        <div className="text-right">
          <div className="text-xs uppercase text-stone-400">Outstanding</div>
          <div className="font-ledger text-xl font-semibold">{money(outstanding)}</div>
          {!closed && outstanding > 0 && (
            <button
              onClick={() => {
                const note = window.prompt(`Write off ${money(outstanding)} remaining on this loan as a bad debt?\n\nOptional note (reason):`, "");
                if (note !== null) onWriteOff(note);
              }}
              className="text-[11px] text-rose-600 hover:text-rose-800 underline decoration-dashed underline-offset-2 mt-1"
            >
              Write off remaining balance
            </button>
          )}
        </div>
      </div>
      <div className="overflow-x-auto ledger-rule">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-xs uppercase tracking-wide text-stone-400 border-b border-stone-200">
              <th className="py-2 pr-3">#</th><th className="py-2 pr-3">Due date</th><th className="py-2 pr-3">Principal</th>
              <th className="py-2 pr-3">Interest</th><th className="py-2 pr-3">Amount</th>
              <th className="py-2 pr-3">Date of collection</th><th className="py-2 pr-3">Status</th>
            </tr>
          </thead>
          <tbody>
            {loan.schedule.map((inst) => (
              <tr key={inst.id} className="h-10">
                <td className="pr-3 text-stone-500">{inst.seq}</td>
                <td className="pr-3 text-stone-600">{fmtDate(inst.dueDate)}</td>
                <td className="pr-3 font-ledger text-stone-600">{money(inst.principalComponent)}</td>
                <td className="pr-3 font-ledger text-stone-600">{money(inst.interestComponent)}</td>
                <td className="pr-3 font-ledger font-medium">{money(inst.amount)}</td>
                <td className="pr-3 font-ledger text-stone-600">{inst.paidDate ? fmtDate(inst.paidDate) : "—"}</td>
                <td className="pr-3"><PayRow inst={inst} today={today} onPay={onPay} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function LoansPage({ data, today, actions }) {
  const [open, setOpen] = useState(false);
  const [selected, setSelected] = useState(null);
  const [form, setForm] = useState({
    clientId: data.clients[0]?.id || "", principal: 20000, annualRatePct: 24,
    installments: 20, frequency: "weekly", customDays: 30, startDate: new Date().toISOString().slice(0, 10),
    mode: "rate", fixedEmi: 5000,
  });

  // Live preview of the implied interest rate when the owner fixes the EMI.
  const emiPreview = useMemo(() => {
    if (form.mode !== "emi") return null;
    const P = Number(form.principal), emi = Number(form.fixedEmi), n = Number(form.installments);
    const r = solvePeriodicRateFromEmi(P, emi, n);
    if (r == null) return { error: true };
    const annualRatePct = periodicRateToAnnualPct(r, form.frequency, Number(form.customDays));
    return { annualRatePct };
  }, [form.mode, form.principal, form.fixedEmi, form.installments, form.frequency, form.customDays]);

  function submit() {
    if (!form.clientId || form.principal <= 0 || form.installments <= 0) return;
    let annualRatePct;
    if (form.mode === "emi") {
      const r = solvePeriodicRateFromEmi(Number(form.principal), Number(form.fixedEmi), Number(form.installments));
      if (r == null) { alert("That fixed EMI is too low to ever repay the loan amount over this tenure. Raise the EMI or reduce the tenure."); return; }
      annualRatePct = periodicRateToAnnualPct(r, form.frequency, Number(form.customDays));
    } else {
      annualRatePct = Number(form.annualRatePct);
    }
    const loanInput = {
      id: uid("ln"), clientId: form.clientId, principal: Number(form.principal), annualRatePct,
      installments: Number(form.installments), frequency: form.frequency,
      customDays: form.frequency === "custom" ? Number(form.customDays) : null,
      startDate: new Date(form.startDate).toISOString(), status: "active",
    };
    const { schedule } = generateSchedule(loanInput);
    actions.addLoan({ ...loanInput, schedule });
    setOpen(false);
  }

  if (selected) {
    const loan = data.loans.find((l) => l.id === selected);
    const client = data.clients.find((c) => c.id === loan.clientId);
    return <LoanDetail loan={loan} client={client} today={today} onClose={() => setSelected(null)} onPay={(instId, amt, dt) => actions.recordPayment(loan.id, instId, amt, dt)} onWriteOff={(note) => actions.writeOffLoan(loan.id, note)} />;
  }

  return (
    <div>
      <SectionHeader
        title="Loans"
        subtitle={`${data.loans.length} issued · ${data.loans.filter((l) => !loanIsClosed(l)).length} active · ${data.loans.filter((l) => loanIsWrittenOff(l)).length} struck off`}
        action={<Btn icon={Plus} onClick={() => setOpen((v) => !v)}>{open ? "Close" : "New Loan"}</Btn>}
      />
      {open && (
        <div className="bg-white border border-stone-200 rounded-lg p-4 mb-4 grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div className="sm:col-span-3 flex items-center gap-2 -mt-1 mb-1">
            <button
              type="button"
              onClick={() => setForm({ ...form, mode: "rate" })}
              className={`text-xs px-3 py-1.5 rounded-md border ${form.mode !== "emi" ? "bg-slate-900 text-white border-slate-900" : "border-stone-200 text-stone-600"}`}
            >I set the interest rate</button>
            <button
              type="button"
              onClick={() => setForm({ ...form, mode: "emi" })}
              className={`text-xs px-3 py-1.5 rounded-md border ${form.mode === "emi" ? "bg-slate-900 text-white border-slate-900" : "border-stone-200 text-stone-600"}`}
            >I fix the EMI amount</button>
          </div>
          <Field label="Client">
            <select className={inputCls} value={form.clientId} onChange={(e) => setForm({ ...form, clientId: e.target.value })}>
              {data.clients.filter((c) => !c.closed).map((c) => <option key={c.id} value={c.id}>{c.name} — {c.area}</option>)}
            </select>
          </Field>
          <Field label="Loan amount (₹)"><input type="number" className={inputCls} value={form.principal} onChange={(e) => setForm({ ...form, principal: e.target.value })} /></Field>
          {form.mode === "emi" ? (
            <Field label="Fixed EMI amount (₹ per installment)">
              <input type="number" className={inputCls} value={form.fixedEmi} onChange={(e) => setForm({ ...form, fixedEmi: e.target.value })} />
            </Field>
          ) : (
            <Field label="Annual interest rate (%)"><input type="number" className={inputCls} value={form.annualRatePct} onChange={(e) => setForm({ ...form, annualRatePct: e.target.value })} /></Field>
          )}
          <Field label="Repayment frequency">
            <select className={inputCls} value={form.frequency} onChange={(e) => setForm({ ...form, frequency: e.target.value })}>
              <option value="weekly">Weekly</option>
              <option value="monthly">Monthly</option>
              <option value="custom">Custom (days)</option>
            </select>
          </Field>
          {form.frequency === "custom" && (
            <Field label="Days per installment"><input type="number" className={inputCls} value={form.customDays} onChange={(e) => setForm({ ...form, customDays: e.target.value })} /></Field>
          )}
          <Field label="Number of installments (tenure)"><input type="number" className={inputCls} value={form.installments} onChange={(e) => setForm({ ...form, installments: e.target.value })} /></Field>
          <Field label="Start date"><input type="date" className={inputCls} value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} /></Field>
          {form.mode === "emi" && (
            <div className="sm:col-span-3 -mt-1">
              {emiPreview?.error ? (
                <p className="text-xs text-rose-600">This EMI is too low to ever repay {money(form.principal)} over {form.installments} installments. Raise the EMI or shorten the tenure.</p>
              ) : emiPreview ? (
                <p className="text-xs text-stone-500">Implied interest rate: <span className="font-ledger font-medium text-stone-800">{emiPreview.annualRatePct.toFixed(2)}% p.a.</span> (compounding {form.frequency}), based on a fixed {money(form.fixedEmi)} per installment.</p>
              ) : null}
            </div>
          )}
          <div className="flex items-end"><Btn onClick={submit}>Create loan</Btn></div>
        </div>
      )}

      <div className="bg-white border border-stone-200 rounded-lg p-4">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs uppercase tracking-wide text-stone-400 border-b border-stone-200">
                <th className="py-2 pr-3">Client</th><th className="py-2 pr-3">Principal</th><th className="py-2 pr-3">Rate</th>
                <th className="py-2 pr-3">Tenure</th><th className="py-2 pr-3">Next due</th><th className="py-2 pr-3">Outstanding</th>
                <th className="py-2 pr-3">Status</th><th className="py-2 pr-3"></th>
              </tr>
            </thead>
            <tbody>
              {/* Active loans first, then closed/struck-off ones - so a fresh
                  loan on a client who has old closed loans is easy to spot,
                  and doesn't get lost among their history. */}
              {[...data.loans].sort((a, b) => {
                const ca = loanIsClosed(a), cb = loanIsClosed(b);
                if (ca !== cb) return ca ? 1 : -1;
                return new Date(b.startDate) - new Date(a.startDate);
              }).map((loan) => {
                const client = data.clients.find((c) => c.id === loan.clientId);
                const closed = loanIsClosed(loan);
                const meta = loanStatusMeta(loan);
                const closedDate = loanClosedDate(loan);
                const next = nextDueInstallment(loan, today);
                return (
                  <tr key={loan.id} className={`border-b border-stone-100 last:border-0 cursor-pointer hover:bg-stone-50 ${closed ? "opacity-70" : ""}`} onClick={() => setSelected(loan.id)}>
                    <td className="py-2.5 pr-3 font-medium text-stone-800">
                      <span className={closed ? "line-through decoration-stone-400" : ""}>{client?.name}</span>
                      {client?.closed && <span className="ml-1.5 text-[10px] uppercase text-rose-600">(client closed)</span>}
                    </td>
                    <td className="py-2.5 pr-3 font-ledger">{money(loan.principal)}</td>
                    <td className="py-2.5 pr-3 text-stone-600">{loan.annualRatePct}%</td>
                    <td className="py-2.5 pr-3 text-stone-600">{loan.installments} × {loan.frequency}</td>
                    <td className="py-2.5 pr-3 text-stone-600">{closed ? "—" : fmtDateShort(next.dueDate)}</td>
                    <td className="py-2.5 pr-3 font-ledger">{closed ? <span className="text-stone-400">—</span> : money(loanOutstanding(loan))}</td>
                    <td className="py-2.5 pr-3">
                      {meta.key === "active" && <span className="text-slate-700">Active</span>}
                      {meta.key === "closed" && <span className="text-emerald-700">Closed{closedDate ? ` (${fmtDateShort(closedDate)})` : ""}</span>}
                      {meta.key === "struck_off" && <span className="text-rose-700">Struck Off{closedDate ? ` (${fmtDateShort(closedDate)})` : ""}</span>}
                    </td>
                    <td className="py-2.5 pr-3"><ChevronRight className="w-4 h-4 text-stone-400" /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ============================== OWNER: COLLECTIONS ============================== */

function CollectionsPage({ data, today, actions }) {
  const rows = useMemo(() => allRows(data), [data]);
  const weeks = useMemo(() => getWeeklyStats(data.loans, today, 8), [data.loans, today]);
  const [agentFilter, setAgentFilter] = useState("all");

  const filteredRows = rows.filter((r) => agentFilter === "all" || r.agent?.id === agentFilter);
  const dueSoon = filteredRows
    .filter((r) => { const d = Math.floor((new Date(r.inst.dueDate) - today) / 86400000); return r.inst.status !== "paid" && d >= 0 && d <= 7; })
    .sort((a, b) => new Date(a.inst.dueDate) - new Date(b.inst.dueDate));
  const overdue = filteredRows
    .filter((r) => r.inst.status !== "paid" && new Date(r.inst.dueDate) < today)
    .sort((a, b) => new Date(a.inst.dueDate) - new Date(b.inst.dueDate));

  return (
    <div>
      <SectionHeader title="Collections" subtitle="Weekly totals, upcoming dues and arrears across all agents" />
      <div className="bg-white border border-stone-200 rounded-lg p-4 mb-6">
        <ResponsiveContainer width="100%" height={220}>
          <LineChart data={weeks} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e7e2d9" vertical={false} />
            <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#78716c" }} axisLine={{ stroke: "#e7e2d9" }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: "#78716c" }} axisLine={false} tickLine={false} tickFormatter={(v) => moneyCompact(v)} width={55} />
            <Tooltip formatter={(v) => money(v)} contentStyle={{ fontSize: 12, borderRadius: 8, borderColor: "#e7e2d9" }} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Line type="monotone" dataKey="due" name="Due" stroke="#a8a29e" strokeWidth={2} dot={false} />
            <Line type="monotone" dataKey="collected" name="Collected" stroke="#059669" strokeWidth={2} dot={{ r: 3 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <div className="mb-4">
        <Field label="Filter by agent">
          <select className={inputCls + " max-w-xs"} value={agentFilter} onChange={(e) => setAgentFilter(e.target.value)}>
            <option value="all">All agents</option>
            {data.agents.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
          </select>
        </Field>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          <h3 className="font-display font-semibold text-stone-900 mb-3">Due before due date</h3>
          {dueSoon.length === 0 ? <p className="text-sm text-stone-500">Nothing upcoming.</p> : (
            <div className="flex flex-col divide-y divide-stone-100">
              {dueSoon.map((r) => (
                <div key={r.inst.id} className="py-2 flex items-center justify-between text-sm gap-2">
                  <div>
                    <div className="font-medium text-stone-800">{r.client?.name}</div>
                    <div className="text-xs text-stone-400">{r.agent?.name} · {fmtDateShort(r.inst.dueDate)}</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="font-ledger">{money(r.inst.amount - (r.inst.paidAmount || 0))}</span>
                    <PayRow inst={r.inst} today={today} onPay={(id, amt, dt) => actions.recordPayment(r.loan.id, id, amt, dt)} />
                    <ReminderButton client={r.client} inst={r.inst} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          <h3 className="font-display font-semibold text-stone-900 mb-3">In arrears</h3>
          {overdue.length === 0 ? <p className="text-sm text-stone-500">No arrears. Clean book!</p> : (
            <div className="flex flex-col divide-y divide-stone-100">
              {overdue.map((r) => (
                <div key={r.inst.id} className="py-2 flex items-center justify-between text-sm gap-2">
                  <div>
                    <div className="font-medium text-stone-800">{r.client?.name}</div>
                    <div className="text-xs text-stone-400">{r.agent?.name} · was due {fmtDateShort(r.inst.dueDate)}</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="font-ledger">{money(r.inst.amount - (r.inst.paidAmount || 0))}</span>
                    <PayRow inst={r.inst} today={today} onPay={(id, amt, dt) => actions.recordPayment(r.loan.id, id, amt, dt)} />
                    <ReminderButton client={r.client} inst={r.inst} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================== OWNER: REINVEST ============================== */

function ReinvestPage({ data, actions }) {
  const rows = useMemo(() => allRows(data), [data]);
  const totalCollected = rows.reduce((s, r) => s + (r.inst.paidAmount || 0), 0);

  const [mode, setMode] = useState("auto"); // "auto" = derive rate from weekly collection; "manual" = enter a rate directly
  const [amount, setAmount] = useState(Math.round(totalCollected / 2) || 100000);
  const [rate, setRate] = useState(12);
  const [compounding, setCompounding] = useState(12);
  const [months, setMonths] = useState(12);
  const [weeklyCollection, setWeeklyCollection] = useState(5000);
  const [cycleWeeks, setCycleWeeks] = useState(25);
  const [reinvestWeeks, setReinvestWeeks] = useState(25);

  // Auto mode: you receive `weeklyCollection` back every week on a loan of
  // `amount`, running `cycleWeeks` weeks. Back-solve the weekly rate that's
  // implied by those real collection figures, so the calculator matches
  // what you're actually being paid rather than a guessed annual rate.
  const autoCalc = useMemo(() => {
    if (mode !== "auto") return null;
    const r = solvePeriodicRateFromEmi(Number(amount), Number(weeklyCollection), Number(cycleWeeks));
    if (r == null) return { error: true };
    return { weeklyRate: r, annualRatePct: periodicRateToAnnualPct(r, "weekly") };
  }, [mode, amount, weeklyCollection, cycleWeeks]);

  // Unify both modes into one periodic rate + one period count. Both modes
  // model money being redeployed each period (weekly collections rolled
  // straight into a fresh cycle in auto mode; the chosen compounding
  // frequency in manual mode) - so growth compounds period over period,
  // not just linearly on the original amount.
  const periodicRate = mode === "auto" ? (autoCalc && !autoCalc.error ? autoCalc.weeklyRate : 0) : (rate / 100) / compounding;
  const periodsElapsed = mode === "auto" ? Number(reinvestWeeks) : Math.round(compounding * (months / 12));
  const displayAnnualPct = mode === "auto" ? (autoCalc && !autoCalc.error ? autoCalc.annualRatePct : 0) : rate;

  // Compound interest: each period's return is calculated on the balance
  // including everything reinvested so far, not just the original amount -
  // this is what "redeploying every collection" actually means financially.
  const fv = amount * Math.pow(1 + periodicRate, periodsElapsed);
  const interestEarned = fv - amount;
  const totalReturnPct = amount > 0 ? (fv / amount - 1) * 100 : 0;
  // The effective annual rate once compounding is accounted for - this is
  // higher than the quoted/nominal rate above whenever periods > 1, since
  // each redeployed rupee starts earning its own return.
  const periodsPerYear = mode === "auto" ? 52 : compounding;
  const effectiveAnnualPct = (Math.pow(1 + periodicRate, periodsPerYear) - 1) * 100;

  const chartData = useMemo(() => {
    const pts = [];
    const steps = Math.max(1, Math.round(periodsElapsed));
    for (let i = 0; i <= steps; i++) {
      const val = amount * Math.pow(1 + periodicRate, i);
      pts.push({ period: i, value: Math.round(val) });
    }
    return pts;
  }, [amount, periodicRate, periodsElapsed]);

  return (
    <div>
      <SectionHeader title="Reinvestment Calculator" subtitle="Model the return from redeploying collected funds, entirely or in part" />
      <div className="bg-slate-900 text-white rounded-lg p-4 mb-4 flex flex-wrap items-center justify-between gap-2">
        <span className="text-sm text-slate-300">Total collected to date (available to reinvest)</span>
        <span className="font-ledger text-xl font-semibold">{money(totalCollected)}</span>
      </div>

      <div className="flex items-center gap-2 mb-4">
        <button
          type="button"
          onClick={() => setMode("auto")}
          className={`text-xs px-3 py-1.5 rounded-md border ${mode === "auto" ? "bg-slate-900 text-white border-slate-900" : "border-stone-200 text-stone-600"}`}
        >Auto-calc from weekly collection</button>
        <button
          type="button"
          onClick={() => setMode("manual")}
          className={`text-xs px-3 py-1.5 rounded-md border ${mode !== "auto" ? "bg-slate-900 text-white border-slate-900" : "border-stone-200 text-stone-600"}`}
        >I'll enter a rate manually</button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="bg-white border border-stone-200 rounded-lg p-4 flex flex-col gap-3">
          <Field label="Amount reinvested / principal (₹)">
            <input type="number" className={inputCls} value={amount} onChange={(e) => setAmount(Number(e.target.value))} />
          </Field>

          {mode === "auto" ? (
            <>
              <Field label="Weekly collection you receive back (₹)">
                <input type="number" className={inputCls} value={weeklyCollection} onChange={(e) => setWeeklyCollection(Number(e.target.value))} />
              </Field>
              <Field label="Tenure of each loan cycle (weeks)">
                <input type="number" className={inputCls} value={cycleWeeks} onChange={(e) => setCycleWeeks(Number(e.target.value))} />
              </Field>
              <Field label="Weeks you keep rolling the collections over">
                <input type="number" className={inputCls} value={reinvestWeeks} onChange={(e) => setReinvestWeeks(Number(e.target.value))} />
              </Field>
              {autoCalc?.error ? (
                <p className="text-xs text-rose-600">A weekly collection of {money(weeklyCollection)} on {money(amount)} over {cycleWeeks} weeks never repays the principal — raise the weekly amount or shorten the cycle.</p>
              ) : (
                <p className="text-xs text-stone-500">Implied weekly rate: <span className="font-ledger font-medium text-stone-800">{(periodicRate * 100).toFixed(3)}%</span> flat per cycle, i.e. <span className="font-ledger font-medium text-stone-800">{displayAnnualPct.toFixed(2)}% p.a.</span> nominal (quoted) — but since every week's collection is immediately redeployed, it compounds weekly to an effective <span className="font-ledger font-medium text-stone-800">{effectiveAnnualPct.toFixed(2)}% p.a.</span> over a full year.</p>
              )}
            </>
          ) : (
            <>
              <Field label="Annual rate offered (%)">
                <input type="number" className={inputCls} value={rate} onChange={(e) => setRate(Number(e.target.value))} />
              </Field>
              <Field label="Compounding frequency">
                <select className={inputCls} value={compounding} onChange={(e) => setCompounding(Number(e.target.value))}>
                  <option value={12}>Monthly</option>
                  <option value={4}>Quarterly</option>
                  <option value={2}>Half-yearly</option>
                  <option value={1}>Annually</option>
                </select>
              </Field>
              <Field label="Reinvestment period (months)">
                <input type="number" className={inputCls} value={months} onChange={(e) => setMonths(Number(e.target.value))} />
              </Field>
            </>
          )}
        </div>

        <div className="lg:col-span-2 grid grid-cols-2 gap-3 content-start">
          <StatCard icon={IndianRupee} label="Value after the period" value={moneyCompact(fv)} sub={money(fv)} tone="emerald" />
          <StatCard icon={TrendingUp} label="Interest earned" value={moneyCompact(interestEarned)} sub={`${totalReturnPct.toFixed(2)}% total return`} tone="emerald" />
          <StatCard icon={Landmark} label="Nominal Rate (quoted)" value={`${displayAnnualPct.toFixed(2)}%`} sub={mode === "auto" ? "Implied by your weekly collection, before compounding" : "Rate as quoted, per annum"} tone="slate" />
          <StatCard icon={TrendingUp} label="Effective Annual Rate" value={`${effectiveAnnualPct.toFixed(2)}%`} sub={`After compounding every ${mode === "auto" ? "week" : "period"} — the rate you actually earn`} tone="amber" />

          <div className="col-span-2 bg-white border border-stone-200 rounded-lg p-4 mt-1">
            <h3 className="text-sm font-semibold text-stone-700 mb-2">Growth of reinvested amount (compounding)</h3>
            <ResponsiveContainer width="100%" height={180}>
              <LineChart data={chartData} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e7e2d9" vertical={false} />
                <XAxis dataKey="period" tick={{ fontSize: 11, fill: "#78716c" }} axisLine={{ stroke: "#e7e2d9" }} tickLine={false} label={{ value: mode === "auto" ? "week" : "compounding period", position: "insideBottom", offset: -3, fontSize: 10, fill: "#a8a29e" }} />
                <YAxis tick={{ fontSize: 11, fill: "#78716c" }} axisLine={false} tickLine={false} tickFormatter={(v) => moneyCompact(v)} width={55} />
                <Tooltip formatter={(v) => money(v)} contentStyle={{ fontSize: 12, borderRadius: 8, borderColor: "#e7e2d9" }} />
                <Line type="monotone" dataKey="value" stroke="#0f172a" strokeWidth={2} dot={false} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <ReinvestmentLog data={data} actions={actions} />
    </div>
  );
}

function ReinvestmentLog({ data, actions }) {
  const log = data.reinvestments || [];
  const defaultEnd = addDays(new Date(), 175); // 25 weeks, matches the worked example
  const [form, setForm] = useState({
    amount: 100000,
    startDate: new Date().toISOString().slice(0, 10),
    endDate: defaultEnd.toISOString().slice(0, 10),
    returnAmount: "",
  });

  function addEntry() {
    const amount = Number(form.amount), returnAmount = Number(form.returnAmount);
    if (!(amount > 0) || !(returnAmount > 0) || !form.startDate || !form.endDate) return;
    if (new Date(form.endDate) <= new Date(form.startDate)) return;
    actions.addReinvestment({
      id: uid("ri"), amount, returnAmount,
      startDate: new Date(form.startDate).toISOString(),
      endDate: new Date(form.endDate).toISOString(),
    });
    setForm({ ...form, returnAmount: "" });
  }

  // Each logged cycle's own return, plus that return annualised on its own
  // using simple (linear) annualisation - not compounded.
  const rows = useMemo(() => {
    return [...log]
      .sort((a, b) => new Date(a.startDate) - new Date(b.startDate))
      .map((e) => {
        const days = Math.max(1, Math.round((new Date(e.endDate) - new Date(e.startDate)) / 86400000));
        const growth = e.returnAmount / e.amount;
        const totalReturnPct = (growth - 1) * 100;
        const annualPct = totalReturnPct * (365 / days);
        return { ...e, days, growth, totalReturnPct, annualPct };
      });
  }, [log]);

  // Chains every logged cycle together (as if each payout were rolled straight
  // into the next cycle) to get the total value grown across the whole
  // tracked history, then annualises that total return with simple interest
  // (linearly, by day-count) rather than compounding it.
  const overall = useMemo(() => {
    if (rows.length === 0) return null;
    const first = rows[0], last = rows[rows.length - 1];
    const totalDays = Math.max(1, Math.round((new Date(last.endDate) - new Date(first.startDate)) / 86400000));
    const overallGrowth = rows.reduce((g, e) => g * e.growth, 1);
    const totalGrowthPct = (overallGrowth - 1) * 100;
    const overallAnnualPct = totalGrowthPct * (365 / totalDays);
    return {
      totalDays, overallAnnualPct, cycles: rows.length,
      totalInvested: first.amount,
      finalValue: first.amount * overallGrowth,
      totalGrowthPct,
    };
  }, [rows]);

  return (
    <div className="mt-6">
      <SectionHeader
        title="Reinvestment Log (actual)"
        subtitle="Log what you actually reinvested and got back each cycle — the effective annual rate below is calculated from these real figures, not assumptions"
      />
      <div className="bg-white border border-stone-200 rounded-lg p-4 mb-4 grid grid-cols-1 sm:grid-cols-4 gap-3">
        <Field label="Amount reinvested (₹)"><input type="number" className={inputCls} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></Field>
        <Field label="Cycle start date"><input type="date" className={inputCls} value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} /></Field>
        <Field label="Cycle end date (when repaid)"><input type="date" className={inputCls} value={form.endDate} onChange={(e) => setForm({ ...form, endDate: e.target.value })} /></Field>
        <Field label="Amount received back (₹)"><input type="number" className={inputCls} value={form.returnAmount} onChange={(e) => setForm({ ...form, returnAmount: e.target.value })} placeholder="e.g. 112000" /></Field>
        <div className="sm:col-span-4 flex items-end"><Btn onClick={addEntry}>Add cycle</Btn></div>
      </div>

      {overall && (
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4">
          <StatCard icon={RefreshCw} label="Cycles logged" value={overall.cycles} tone="slate" />
          <StatCard icon={IndianRupee} label="Value today" value={moneyCompact(overall.finalValue)} sub={money(overall.finalValue)} tone="emerald" />
          <StatCard icon={TrendingUp} label="Total growth" value={`${overall.totalGrowthPct.toFixed(2)}%`} sub={`over ${overall.totalDays} days, ${overall.cycles} cycle(s)`} tone="amber" />
          <StatCard icon={TrendingUp} label="Effective annual rate" value={`${overall.overallAnnualPct.toFixed(2)}%`} sub="Actual return, annualised on a simple-interest (linear) basis" tone="amber" />
        </div>
      )}

      {rows.length > 0 && (
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs uppercase tracking-wide text-stone-400 border-b border-stone-200">
                  <th className="py-2 pr-3">Start</th><th className="py-2 pr-3">End</th><th className="py-2 pr-3">Days</th>
                  <th className="py-2 pr-3">Invested</th><th className="py-2 pr-3">Received back</th>
                  <th className="py-2 pr-3">Cycle return</th><th className="py-2 pr-3">Annualised</th><th className="py-2 pr-3"></th>
                </tr>
              </thead>
              <tbody>
                {rows.map((e) => (
                  <tr key={e.id} className="border-b border-stone-100 last:border-0">
                    <td className="py-2.5 pr-3 text-stone-600">{fmtDateShort(e.startDate)}</td>
                    <td className="py-2.5 pr-3 text-stone-600">{fmtDateShort(e.endDate)}</td>
                    <td className="py-2.5 pr-3 text-stone-600">{e.days}</td>
                    <td className="py-2.5 pr-3 font-ledger">{money(e.amount)}</td>
                    <td className="py-2.5 pr-3 font-ledger">{money(e.returnAmount)}</td>
                    <td className="py-2.5 pr-3 font-ledger text-emerald-700">{e.totalReturnPct.toFixed(2)}%</td>
                    <td className="py-2.5 pr-3 font-ledger text-amber-700">{e.annualPct.toFixed(2)}%</td>
                    <td className="py-2.5 pr-3">
                      <button onClick={() => actions.deleteReinvestment(e.id)} title="Remove" className="text-stone-400 hover:text-rose-600">
                        <X className="w-3.5 h-3.5" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================== AGENT VIEW ============================== */

function AgentHome({ data, agentId, today, subview, actions }) {
  const agent = data.agents.find((a) => a.id === agentId);
  const myClients = data.clients.filter((c) => c.agentId === agentId);
  const rows = allRows(data).filter((r) => r.client?.agentId === agentId);

  if (subview === "arrears") {
    const overdue = rows.filter((r) => r.inst.status !== "paid" && new Date(r.inst.dueDate) < today)
      .sort((a, b) => new Date(a.inst.dueDate) - new Date(b.inst.dueDate));
    const total = overdue.reduce((s, r) => s + (r.inst.amount - (r.inst.paidAmount || 0)), 0);
    return (
      <div>
        <SectionHeader title="Arrears" subtitle={`Overdue collections for ${agent?.name}`} />
        <div className="bg-rose-600 text-white rounded-lg p-4 mb-4 flex items-center justify-between">
          <span className="text-sm">Total in arrears</span>
          <span className="font-ledger text-xl font-semibold">{money(total)}</span>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          {overdue.length === 0 ? <p className="text-sm text-stone-500 text-center py-6">No arrears — every customer is up to date.</p> : (
            <div className="flex flex-col divide-y divide-stone-100">
              {overdue.map((r) => (
                <div key={r.inst.id} className="py-3 flex items-center justify-between gap-2 flex-wrap">
                  <div>
                    <div className="font-medium text-stone-800">{r.client?.name}</div>
                    <div className="text-xs text-stone-400 flex items-center gap-1"><MapPin className="w-3 h-3" />{r.client?.area} · was due {fmtDateShort(r.inst.dueDate)}</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="font-ledger font-medium">{money(r.inst.amount - (r.inst.paidAmount || 0))}</span>
                    <PayRow inst={r.inst} today={today} onPay={(id, amt, dt) => actions.recordPayment(r.loan.id, id, amt, dt)} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }

  if (subview === "duesoon") {
    const dueSoon = rows.filter((r) => { const d = Math.floor((new Date(r.inst.dueDate) - today) / 86400000); return r.inst.status !== "paid" && d >= 0 && d <= 7; })
      .sort((a, b) => new Date(a.inst.dueDate) - new Date(b.inst.dueDate));
    const total = dueSoon.reduce((s, r) => s + (r.inst.amount - (r.inst.paidAmount || 0)), 0);
    return (
      <div>
        <SectionHeader title="Due Soon" subtitle="Collections expected in the next 7 days" />
        <div className="bg-amber-500 text-white rounded-lg p-4 mb-4 flex items-center justify-between">
          <span className="text-sm">Expected this week</span>
          <span className="font-ledger text-xl font-semibold">{money(total)}</span>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          {dueSoon.length === 0 ? <p className="text-sm text-stone-500 text-center py-6">Nothing due in the next 7 days.</p> : (
            <div className="flex flex-col divide-y divide-stone-100">
              {dueSoon.map((r) => (
                <div key={r.inst.id} className="py-3 flex items-center justify-between gap-2 flex-wrap">
                  <div>
                    <div className="font-medium text-stone-800">{r.client?.name}</div>
                    <div className="text-xs text-stone-400 flex items-center gap-1"><MapPin className="w-3 h-3" />{r.client?.area} · due {fmtDateShort(r.inst.dueDate)}</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="font-ledger font-medium">{money(r.inst.amount - (r.inst.paidAmount || 0))}</span>
                    <PayRow inst={r.inst} today={today} onPay={(id, amt, dt) => actions.recordPayment(r.loan.id, id, amt, dt)} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }

  // customers view (default), grouped by area
  const activeClientCount = myClients.filter((c) => !c.closed).length;
  const areas = [...new Set(myClients.map((c) => c.area))];
  return (
    <div>
      <SectionHeader title="My Customers" subtitle={`${activeClientCount} active customers assigned to ${agent?.name}`} />
      {areas.map((area) => {
        const areaClients = [...myClients.filter((c) => c.area === area)].sort((a, b) => (a.closed === b.closed ? 0 : a.closed ? 1 : -1));
        return (
          <div key={area} className="mb-5">
            <h3 className="text-xs uppercase tracking-wide text-stone-400 font-medium mb-2 flex items-center gap-1"><MapPin className="w-3.5 h-3.5" />{area}</h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {areaClients.map((c) => {
                const loans = data.loans.filter((l) => l.clientId === c.id);
                return (
                  <div key={c.id} className={`bg-white border border-stone-200 rounded-lg p-3.5 ${c.closed ? "opacity-60" : ""}`}>
                    <div className="flex items-center justify-between mb-2">
                      <div className={`font-medium text-stone-900 ${c.closed ? "line-through decoration-stone-400" : ""}`}>{c.name}</div>
                      <span className="text-xs text-stone-400">{c.phone}</span>
                    </div>
                    {c.closed && <div className="text-[10px] uppercase tracking-wide text-rose-600 -mt-1 mb-1.5">Closed as on {fmtDateShort(c.closedAt || Date.now())}</div>}
                    {loans.length === 0 ? <p className="text-xs text-stone-400">No active loans</p> : [...loans].sort((a, b) => {
                      const ca = loanIsClosed(a), cb = loanIsClosed(b);
                      if (ca !== cb) return ca ? 1 : -1;
                      return new Date(b.startDate) - new Date(a.startDate);
                    }).map((loan) => {
                      const closed = loanIsClosed(loan);
                      const meta = loanStatusMeta(loan);
                      const next = nextDueInstallment(loan, today);
                      return (
                        <div key={loan.id} className="flex items-center justify-between text-sm border-t border-stone-100 pt-2 mt-2 first:border-0 first:pt-0 first:mt-0">
                          <div>
                            <div className={`text-stone-600 ${closed ? "line-through decoration-stone-400" : ""}`}>{money(loanOutstanding(loan))} outstanding</div>
                            {!closed && <div className="text-xs text-stone-400">next: {fmtDateShort(next.dueDate)} · {money(next.amount - (next.paidAmount || 0))}</div>}
                          </div>
                          {meta.key === "struck_off" ? <StatusStamp meta={{ key: "written_off", label: "STRUCK OFF" }} />
                            : closed ? <StatusStamp meta={{ key: "paid", label: "CLOSED" }} />
                            : <PayRow inst={next} today={today} onPay={(id, amt, dt) => actions.recordPayment(loan.id, id, amt, dt)} />}
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ============================== ROOT APP ============================== */

export default function App() {
  const [data, setData] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [session, setSession] = useState(null);
  const [ownerTab, setOwnerTab] = useState("dashboard");
  const [agentTab, setAgentTab] = useState("customers");
  const [lastSynced, setLastSynced] = useState(null);
  const [syncing, setSyncing] = useState(false);

  const today = useMemo(() => startOfDay(new Date()), []);

  async function pullLatest(showSpinner) {
    if (showSpinner) setSyncing(true);
    try {
      const res = await storage.get(STORAGE_KEY);
      if (res && res.value) {
        const remote = JSON.parse(res.value);
        setData((prev) => (JSON.stringify(prev) !== JSON.stringify(remote) ? remote : prev));
        setLastSynced(new Date());
      }
    } catch (e) {
      // keep showing whatever is already on screen if a pull fails
    } finally {
      if (showSpinner) setSyncing(false);
    }
  }

  useEffect(() => {
    (async () => {
      try {
        const res = await storage.get(STORAGE_KEY);
        if (res && res.value) { setData(JSON.parse(res.value)); setLastSynced(new Date()); }
        else {
          const seeded = seedData();
          setData(seeded);
          try { await storage.set(STORAGE_KEY, JSON.stringify(seeded)); setLastSynced(new Date()); } catch (e) {}
        }
      } catch (e) {
        const seeded = seedData();
        setData(seeded);
        try { await storage.set(STORAGE_KEY, JSON.stringify(seeded)); } catch (e2) {}
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  // Keep everyone's screen up to date: re-check the shared sheet every 15s,
  // and immediately whenever this tab/app regains focus.
  useEffect(() => {
    if (!loaded || !syncConfigured()) return;
    const interval = setInterval(() => pullLatest(false), 15000);
    const onFocus = () => pullLatest(false);
    window.addEventListener("focus", onFocus);
    return () => { clearInterval(interval); window.removeEventListener("focus", onFocus); };
  }, [loaded]);

  async function persist(newData, opts) {
    setData(newData); // show the change instantly, even if this device is offline
    try {
      let toSave = newData;
      if (syncConfigured() && !(opts && opts.overwrite)) {
        // Pull whatever the shared Sheet has right now (may include changes
        // from another agent's device that synced while this one was
        // offline) and merge record-by-record instead of overwriting it.
        const latest = await storage.get(STORAGE_KEY);
        if (latest && latest.value) {
          try {
            const remote = JSON.parse(latest.value);
            toSave = mergeData(remote, newData);
          } catch (e) { /* remote value unreadable - fall back to just saving ours */ }
        }
      }
      await storage.set(STORAGE_KEY, JSON.stringify(toSave));
      setData(toSave);
      setLastSynced(new Date());
    } catch (e) {
      console.error("save failed", e);
    }
  }

  const actions = {
    addClient: (client) => persist({ ...data, clients: [...data.clients, client] }),
    addAgent: (agent) => persist({ ...data, agents: [...data.agents, agent] }),
    // Both of these are only ever called from owner-only screens (ClientsPage,
    // AgentsPage) - agents have no route to reach them.
    //
    // "Deleting" a client or agent never actually erases their record - it
    // marks them closed/struck-off as on the date of closure, so their loan
    // and collection history stays intact for the books. Use reopenClient /
    // reopenAgent to undo a closure.
    closeClient: (clientId, note) => {
      const newData = JSON.parse(JSON.stringify(data));
      const client = newData.clients.find((c) => c.id === clientId);
      if (!client) return;
      const clientLoans = newData.loans.filter((l) => l.clientId === clientId);
      const now = new Date().toISOString();
      const newWriteOffs = [];
      clientLoans.forEach((loan) => {
        let amountWrittenOff = 0;
        loan.schedule.forEach((i) => {
          if (i.status !== "paid" && i.status !== "written_off") {
            amountWrittenOff += Math.max(0, i.amount - (i.paidAmount || 0));
            i.status = "written_off";
            i.paidDate = i.paidDate || now;
          }
        });
        if (amountWrittenOff > 0) {
          newWriteOffs.push({
            id: uid("wo"), loanId: loan.id, clientId, clientName: client.name,
            amount: Math.round(amountWrittenOff * 100) / 100, date: now, note: note || "Client closed",
          });
        }
      });
      if (newWriteOffs.length) newData.writeOffs = [...(newData.writeOffs || []), ...newWriteOffs];
      client.closed = true;
      client.closedAt = now;
      persist(newData);
    },
    reopenClient: (clientId) => {
      const newData = JSON.parse(JSON.stringify(data));
      const client = newData.clients.find((c) => c.id === clientId);
      if (!client) return;
      client.closed = false;
      client.closedAt = null;
      persist(newData);
    },
    closeAgent: (id) => {
      const newData = JSON.parse(JSON.stringify(data));
      const agent = newData.agents.find((a) => a.id === id);
      if (!agent) return;
      agent.closed = true;
      agent.closedAt = new Date().toISOString();
      persist(newData);
    },
    reopenAgent: (id) => {
      const newData = JSON.parse(JSON.stringify(data));
      const agent = newData.agents.find((a) => a.id === id);
      if (!agent) return;
      agent.closed = false;
      agent.closedAt = null;
      persist(newData);
    },
    // Marks whatever is still unpaid on a loan as a bad-debt write-off
    // (owner-only, from the loan detail screen) - zeroes its outstanding
    // balance without pretending the cash was actually collected, and keeps
    // a running record of total write-offs for bookkeeping. The loan itself
    // is never deleted - it stays on record, closed/struck-off as on today.
    writeOffLoan: (loanId, note) => {
      const newData = JSON.parse(JSON.stringify(data));
      const loan = newData.loans.find((l) => l.id === loanId);
      if (!loan) return;
      const client = newData.clients.find((c) => c.id === loan.clientId);
      let amountWrittenOff = 0;
      const now = new Date().toISOString();
      loan.schedule.forEach((i) => {
        if (i.status !== "paid") {
          amountWrittenOff += Math.max(0, i.amount - (i.paidAmount || 0));
          i.status = "written_off";
          i.paidDate = i.paidDate || now;
        }
      });
      if (amountWrittenOff > 0) {
        newData.writeOffs = [...(newData.writeOffs || []), {
          id: uid("wo"), loanId, clientId: loan.clientId, clientName: client?.name || "Unknown", amount: Math.round(amountWrittenOff * 100) / 100, date: now, note: note || "",
        }];
      }
      persist(newData);
    },
    addLoan: (loan) => persist({ ...data, loans: [...data.loans, loan] }),
    recordPayment: (loanId, instId, amount, collectedDate) => {
      const newData = JSON.parse(JSON.stringify(data));
      const loan = newData.loans.find((l) => l.id === loanId);
      if (!loan) return;
      const inst = loan.schedule.find((i) => i.id === instId);
      if (!inst) return;
      const collDate = collectedDate ? new Date(collectedDate) : new Date();
      const newPaid = (inst.paidAmount || 0) + Number(amount || 0);
      inst.paidAmount = newPaid;
      inst.paidDate = collDate.toISOString();
      inst.status = newPaid >= inst.amount - 0.5 ? "paid" : "partial";
      persist(newData);
    },
    addReinvestment: (entry) => persist({ ...data, reinvestments: [...(data.reinvestments || []), entry] }),
    deleteReinvestment: (id) => persist({ ...data, reinvestments: (data.reinvestments || []).filter((r) => r.id !== id) }),
    resetDemo: () => {
      const seeded = seedData();
      persist(seeded, { overwrite: true });
    },
    // Wipes everything - agents, clients, loans, write-offs, reinvestment
    // log - back to a completely empty ledger, ready for real data. Unlike
    // resetDemo, this does NOT reseed the sample records.
    eraseAllData: () => {
      persist({ agents: [], clients: [], loans: [], reinvestments: [], writeOffs: [] }, { overwrite: true });
    },
    restoreBackup: (restoredData) => {
      persist(restoredData, { overwrite: true });
    },
  };

  if (!loaded || !data) {
    return (
      <div className="min-h-screen bg-stone-50 flex items-center justify-center app-root">
        <style>{FONT_STYLE}</style>
        <p className="text-stone-400 text-sm">Loading ledger…</p>
      </div>
    );
  }

  if (!session) {
    return <LoginScreen data={data} onLogin={setSession} />;
  }

  const isOwner = session.role === "owner";
  const currentAgent = !isOwner ? data.agents.find((a) => a.id === session.agentId) : null;

  return (
    <div className="app-root min-h-screen bg-stone-50 flex text-stone-900">
      <style>{FONT_STYLE}</style>
      <Sidebar
        role={session.role}
        tab={isOwner ? ownerTab : agentTab}
        setTab={isOwner ? setOwnerTab : setAgentTab}
        onLogout={() => setSession(null)}
        agentLabel={currentAgent?.name}
        onResetDemo={actions.resetDemo}
        onEraseAll={actions.eraseAllData}
        data={data}
        onRestoreBackup={actions.restoreBackup}
        shared={syncConfigured()}
        lastSynced={lastSynced}
        syncing={syncing}
        onSyncNow={() => pullLatest(true)}
      />
      <main className="flex-1 min-h-screen overflow-y-auto">
        <div className="h-16 border-b border-stone-200 bg-white flex items-center px-4 sm:px-6 justify-between">
          <div className="font-display font-semibold text-stone-900">
            {isOwner ? "Owner" : `Agent — ${currentAgent?.name}`}
          </div>
          <div className="text-xs text-stone-400 font-ledger">{fmtDate(today)}</div>
        </div>
        <div className="p-4 sm:p-6 max-w-6xl mx-auto">
          {isOwner ? (
            ownerTab === "dashboard" ? <OwnerDashboard data={data} today={today} onPay={actions.recordPayment} />
            : ownerTab === "clients" ? <ClientsPage data={data} actions={actions} />
            : ownerTab === "agents" ? <AgentsPage data={data} today={today} actions={actions} />
            : ownerTab === "loans" ? <LoansPage data={data} today={today} actions={actions} />
            : ownerTab === "collections" ? <CollectionsPage data={data} today={today} actions={actions} />
            : ownerTab === "reinvest" ? <ReinvestPage data={data} actions={actions} />
            : null
          ) : (
            <AgentHome data={data} agentId={session.agentId} today={today} subview={agentTab} actions={actions} />
          )}
        </div>
      </main>
    </div>
  );
}
