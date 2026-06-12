import React, { useState, useMemo, useCallback, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { X, Search, FileSpreadsheet, Download, Info, Edit3, AlertCircle, FileText, CheckCircle2, ChevronDown, ChevronUp, Filter, XCircle, ArrowUp, ArrowDown, ArrowUpDown, ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight, RefreshCw } from 'lucide-react';
import { storage, dataUrl } from './storage';

// ==================== Constants ====================
const PIPELINE_STAGES = ['Requirement Gathering', 'Deal Solutioning', 'Proposal', 'Negotiation', 'Final Approval', 'Future Follow Up'];
const PREPIPE_STAGES = ['Discovery', 'Deal Owner Qualification', 'Pre-Sales Qualification'];

const VERTICAL_GROUPS: Record<string, string[]> = {
  'US Education':   ['us education', 'us - education', 'education us', 'education'],
  'US Healthcare':  ['us healthcare', 'us - healthcare', 'healthcare us', 'healthcare'],
  'US Others':      ['us others', 'us - others', 'kam us mm', 'kam-us mm', 'others'],
  'BFSI India':     ['bfsi india', 'a.bfsi', 'key growth b', 'bfsi'],
  'EV MM':          ['ev midmarket', 'ev mid market', 'ev mid-market', 'ev - midmarket', 'ev - mid market', 'mid market ev', 'midmarket ev', 'mid-market ev', 'ev mm'],
  'EV ENT':         ['ev enterprise', 'ev - enterprise', 'enterprise ev', 'ev ent'],
  'ME':             ['mena', 'roe', 'me'],
  'Partner Sales':  ['partner sales', 'smb', 'partner'],
};

const TARGET_LABEL_MATCHERS: Record<string, string[]> = {
  'US Education':   ['us education'],
  'US Healthcare':  ['us healthcare'],
  'US Others':      ['kam-us mm', 'us others', 'others'],
  'BFSI India':     ['bfsi india', 'bfsi'],
  'EV MM':          ['mid market', 'midmarket', 'mid-market', 'ev mm'],
  'EV ENT':         ['enterprise', 'ev ent'],
  'ME':             ['mena', 'me', 'roe'],
  'Partner Sales':  ['smb', 'partner sales', 'partner'],
};

const INDIA_VERTICALS = ['BFSI India', 'EV MM', 'EV ENT', 'ME', 'Partner Sales'];
const US_VERTICALS    = ['US Education', 'US Healthcare', 'US Others'];
const isUSLabel = (label: string) => US_VERTICALS.includes(label);

const RENEWAL_BU_MAP: Record<string, string> = {
  'Key Growth FS':          'BFSI India',
  'Key Growth A':           'BFSI India',
  'Key Growth B':           'BFSI India',
  'Key Growth C':           'BFSI India',
  'Key Growth ROW':         'ME',
  'Partner KAM':            'Partner Sales',
  'Key Growth EV':          'EV (MM + ENT)',
  'Retention US Education': 'US Education',
  'Retention US Healthcare':'US Healthcare',
  'Retention US Others':    'US Others',
  'Key Growth US Education':'US Education',
  'Key Growth US Healthcare':'US Healthcare',
};
const RENEWAL_INDIA_VERTICALS = ['BFSI India', 'EV (MM + ENT)', 'ME', 'Partner Sales'];
const RENEWAL_US_VERTICALS    = ['US Education', 'US Healthcare', 'US Others'];

// MRR Performance tab uses a combined EV row (since target is given as a single EV figure).
const MRR_INDIA_VERTICALS = ['BFSI India', 'EV', 'ME', 'Partner Sales'];
const MRR_US_VERTICALS    = ['US Education', 'US Healthcare', 'US Others'];

const MRR_NET_TARGETS: Record<string, Record<string, Record<string, number>>> = {
  India: {
    Q1: { 'BFSI India': 2441743,  'EV': 193887,   'ME': -1366389, 'Partner Sales': -845610 },
    Q2: { 'BFSI India': 5067138,  'EV': 6960744,  'ME': -571263,  'Partner Sales': -1133781 },
    Q3: { 'BFSI India': 0, 'EV': 0, 'ME': 0, 'Partner Sales': 0 },
    Q4: { 'BFSI India': 0, 'EV': 0, 'ME': 0, 'Partner Sales': 0 },
  },
  US: {
    Q1: { 'US Education': 43750, 'US Healthcare': 39586, 'US Others': 0 },
    Q2: { 'US Education': 43750, 'US Healthcare': 39595, 'US Others': 0 },
    Q3: { 'US Education': 0,     'US Healthcare': 0,     'US Others': 0 },
    Q4: { 'US Education': 0,     'US Healthcare': 0,     'US Others': 0 },
  },
};

const FY27_VERTICAL_MAP: Record<string, string> = {
  'a.FS':           'BFSI India',
  'c.ROW':          'ME',
  'd.Partner':      'Partner Sales',
  'e.US Education': 'US Education',
  'f.US Healthcare':'US Healthcare',
  'h.US Others':    'US Others',
};

const FILTER_LABELS: Record<string, string> = {
  vertical: 'Vertical',
  stage: 'Deal Stage',
  status: 'Deal Status',
  typeOfSale: 'Type of Sale',
  dealOwner: 'Opp. Sales Owner',
  employeeSize: 'Employee Range',
  industry: 'Industry',
  productAsk: 'Product Ask',
  icpStatus: 'ICP Account Status',
  currentOwner: 'Current Opp. Owner',
  sdInitiator: 'SD Initiator',
  saInitiator: 'SA Initiator',
  pipeAddedRange: 'Pipe Added Date',
  expClosureRange: 'Expected Closure',
  finalClosureRange: 'Final Closure',
};

// ==================== Date helpers ====================
const today = () => new Date();
const getCurrentFY = () => { const d = today(); return d.getMonth() >= 3 ? d.getFullYear() + 1 : d.getFullYear(); };
const getCurrentQuarter = () => {
  const m = today().getMonth();
  if (m >= 3 && m <= 5) return 'Q1';
  if (m >= 6 && m <= 8) return 'Q2';
  if (m >= 9 && m <= 11) return 'Q3';
  return 'Q4';
};
const quarterRange = (fy: number, q: string): [Date, Date] => {
  const sy = fy - 1;
  return ({
    Q1: [new Date(sy, 3, 1), new Date(sy, 5, 30, 23, 59, 59)],
    Q2: [new Date(sy, 6, 1), new Date(sy, 8, 30, 23, 59, 59)],
    Q3: [new Date(sy, 9, 1), new Date(sy, 11, 31, 23, 59, 59)],
    Q4: [new Date(fy, 0, 1), new Date(fy, 2, 31, 23, 59, 59)],
  } as Record<string, [Date, Date]>)[q];
};
const advanceQ = (q: string, fy: number): [string, number] => {
  const next = ({ Q1: 'Q2', Q2: 'Q3', Q3: 'Q4', Q4: 'Q1' } as Record<string, string>)[q];
  const nfy = q === 'Q4' ? fy + 1 : fy;
  return [next, nfy];
};
const parseDate = (v: any): Date | null => {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v as any) ? null : v;
  if (typeof v === 'number' && v > 25000 && v < 60000) {
    const d = new Date((v - 25569) * 86400000);
    return isNaN(d as any) ? null : d;
  }
  const s = String(v).trim();
  if (!s || s.toLowerCase() === 'null' || s === '-') return null;
  const m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (m) {
    const [, d, mo, y] = m;
    const yr = y.length === 2 ? 2000 + +y : +y;
    const dt = new Date(yr, +mo - 1, +d);
    if (!isNaN(dt as any)) return dt;
  }
  const dt = new Date(s);
  return isNaN(dt as any) ? null : dt;
};
const fyLabel = (fy: number) => `FY${String(fy).slice(-2)}`;
const fmtDate = (d: Date | null) => d ? d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: '2-digit' }) : '—';

// ==================== Currency ====================
const fmtINR = (v: any) => {
  if (v == null || isNaN(v)) return '—';
  const n = Number(v);
  if (n === 0) return '₹0';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1e7) return `${sign}₹${(abs / 1e7).toFixed(2)} Cr`;
  if (abs >= 1e5) return `${sign}₹${(abs / 1e5).toFixed(2)} L`;
  if (abs >= 1e3) return `${sign}₹${(abs / 1e3).toFixed(1)}K`;
  return `${sign}₹${abs.toFixed(0)}`;
};
const fmtUSD = (v: any) => {
  if (v == null || isNaN(v)) return '—';
  const n = Number(v);
  if (n === 0) return '$0';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1e6) return `${sign}$${(abs / 1e6).toFixed(2)}M`;
  if (abs >= 1e3) return `${sign}$${(abs / 1e3).toFixed(1)}K`;
  return `${sign}$${abs.toFixed(0)}`;
};
const fmt = (v: any, region: string) => region === 'US' ? fmtUSD(v) : fmtINR(v);

// ==================== Normalization ====================
const normalizeStage = (s: any) => {
  if (!s) return '';
  const x = String(s).toLowerCase().trim();
  if (x.includes('won') || x === 'closed won') return 'Won';
  if (x.includes('lost') || x === 'closed lost') return 'Lost';
  if (x.includes('discovery')) return 'Discovery';
  if (x.includes('deal owner') || x === 'doq' || x.includes('do qualif')) return 'Deal Owner Qualification';
  if (x.includes('pre-sales') || x.includes('presales') || x === 'psq' || x.includes('ps qualif') || x.includes('pre sales')) return 'Pre-Sales Qualification';
  if (x.includes('requirement')) return 'Requirement Gathering';
  if (x.includes('solutioning') || x.includes('solution')) return 'Deal Solutioning';
  if (x.includes('future follow') || x.includes('future-follow')) return 'Future Follow Up';
  if (x.includes('final approval')) return 'Final Approval';
  if (x.includes('negotiation')) return 'Negotiation';
  if (x.includes('proposal')) return 'Proposal';
  return String(s).trim();
};
const normalizeStatus = (s: any) => {
  if (!s) return 'Open';
  const x = String(s).toLowerCase().trim();
  if (x.includes('won')) return 'Won';
  if (x.includes('lost')) return 'Lost';
  if (x.includes('open') || x.includes('active') || x.includes('progress') || x.includes('in-flight')) return 'Open';
  return String(s).trim();
};

const escapeRegex = (s: string) => s.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');

const matchesLabel = (rawValue: any, label: string) => {
  if (!rawValue) return false;
  const v = String(rawValue).toLowerCase().trim();
  if (!v) return false;
  return (VERTICAL_GROUPS[label] || []).some(m => {
    const lm = m.toLowerCase();
    if (v === lm) return true;
    return new RegExp(`\\b${escapeRegex(lm)}\\b`).test(v);
  });
};

const resolveVertical = (rawVertical: any, rawTeams: any): { label: string | null; source: string | null } => {
  if (rawVertical) {
    for (const label of Object.keys(VERTICAL_GROUPS)) {
      if (matchesLabel(rawVertical, label)) return { label, source: 'vertical' };
    }
  }
  if (rawTeams) {
    for (const label of Object.keys(VERTICAL_GROUPS)) {
      if (matchesLabel(rawTeams, label)) return { label, source: 'teams' };
    }
  }
  return { label: null, source: null };
};

const findCol = (headers: string[], candidates: string[]): string | null => {
  const lh = headers.map(h => (h || '').toLowerCase().trim());
  for (const c of candidates) {
    const idx = lh.findIndex(h => h === c.toLowerCase());
    if (idx >= 0) return headers[idx];
  }
  for (const c of candidates) {
    const idx = lh.findIndex(h => h.includes(c.toLowerCase()));
    if (idx >= 0) return headers[idx];
  }
  return null;
};

// Builds the opportunity column mapping from a set of headers.
const buildOppMapping = (hdrs: string[]) => ({
  dealName:       findCol(hdrs, ['Deal Name', 'Opportunity Name', 'Name']),
  account:        findCol(hdrs, ['Account Name', 'Account', 'Company', 'Customer Name']),
  owner:          findCol(hdrs, ['Opportunity Sales Owner Name', 'Opp Sales Owner Name', 'Opportunity Owner', 'Sales Owner Name', 'Sales Owner', 'Deal Owner', 'Owner']),
  stage:          findCol(hdrs, ['Stage', 'Deal Stage', 'Opportunity Stage', 'StageName']),
  status:         findCol(hdrs, ['Status', 'Deal Status', 'Opportunity Status']),
  mrr:            findCol(hdrs, ['Actual Total Deal MRR', 'Total Deal MRR', 'Deal MRR', 'MRR']),
  closureDate:    findCol(hdrs, ['Final Closure Date', 'Closure Date', 'Close Date']),
  mrrStartDate:   findCol(hdrs, ['MRR Start Date', 'Start Date', 'Revenue Start Date']),
  expClosureDate: findCol(hdrs, ['Expected Closure Date', 'Exp Closure Date', 'Exp. Closure Date', 'Expected Close Date']),
  vertical:       findCol(hdrs, ['Vertical', 'Vertical List', 'Sub-Vertical', 'Business Vertical']),
  teams:          findCol(hdrs, ['Teams', 'Team', 'Team Name', 'Sales Team']),
  pipeAdded:      findCol(hdrs, ['Pipe added date', 'Pipe Added Date', 'Pipeline Added Date', 'Date Added to Pipe']),
  usPipeAdded:    findCol(hdrs, ['Date Changes - US Pipeline Entry Date', 'US Pipeline Entry Date', 'US Pipe Entry Date']),
  typeOfSale:     findCol(hdrs, ['Type of Sale', 'Sale Type', 'Deal Type', 'Opportunity Type', 'Type Of Sale']),
  employeeSize:   findCol(hdrs, ['Employee Range', 'Employee Size', 'Headcount Range', 'Headcount Bucket', 'Company Size', 'Number of Employees', 'Headcount', 'Employees', 'Size Range', 'Employee Bucket']),
  industry:       findCol(hdrs, ['Industry', 'Account Industry', 'Sub-Industry', 'Sub Industry', 'Vertical Industry']),
  productAsk:     findCol(hdrs, ['Product Ask', 'Product Interest', 'Product', 'Product Type', 'Solution Ask', 'Product Family']),
  icpStatus:      findCol(hdrs, ['ICP Account Status', 'ICP Status', 'ICP account status', 'Account ICP Status']),
  currentOwner:   findCol(hdrs, ['Current Opportunity Owner', 'Current Opp Owner', 'Current Owner', 'Opportunity Owner (Current)', 'Current Opportunity Owner (User Name)', 'Current Opp Owner Username']),
  sdInitiator:    findCol(hdrs, ['SD Initiator', 'SD Initiator Name', 'Sales Development Initiator']),
  saInitiator:    findCol(hdrs, ['SA Initiator', 'SA Initiator Name', 'Solution Architect Initiator', 'Sales Architect Initiator']),
});

// ==================== Sort helpers ====================
const compareValues = (a: any, b: any) => {
  if (a == null && b == null) return 0;
  if (a == null) return 1;
  if (b == null) return -1;
  if (a instanceof Date && b instanceof Date) return a.getTime() - b.getTime();
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return String(a).localeCompare(String(b));
};

const SORTABLE_COLUMNS: Record<string, { label: string; getValue: (d: any) => any }> = {
  name:           { label: 'Deal',       getValue: d => (d.name || '').toLowerCase() },
  resolvedLabel:  { label: 'Vertical',   getValue: d => (d.resolvedLabel || '').toLowerCase() },
  stage:          { label: 'Stage',      getValue: d => (d.stage || '').toLowerCase() },
  status:         { label: 'Status',     getValue: d => (d.status || '').toLowerCase() },
  mrr:            { label: 'MRR',        getValue: d => d.mrr || 0 },
  pipeAddedDate:  { label: 'Pipe Added', getValue: d => d.pipeAddedDate },
  closureDate:    { label: 'Closure',    getValue: d => d.closureDate },
  mrrStartDate:   { label: 'MRR Start',  getValue: d => d.mrrStartDate },
  expClosureDate: { label: 'Exp. Close', getValue: d => d.expClosureDate },
};

const sortDeals = (deals: any[], sort: any) => {
  if (!sort?.column || !SORTABLE_COLUMNS[sort.column]) return deals;
  const getValue = SORTABLE_COLUMNS[sort.column].getValue;
  const dir = sort.direction === 'asc' ? 1 : -1;
  return [...deals].sort((a, b) => dir * compareValues(getValue(a), getValue(b)));
};

// ==================== XLSX target parser ====================
const targetLabelFromText = (s: any): string | null => {
  const v = String(s || '').toLowerCase().trim();
  if (!v) return null;
  for (const [label, matchers] of Object.entries(TARGET_LABEL_MATCHERS)) {
    if (matchers.some(m => v === m)) return label;
  }
  for (const [label, matchers] of Object.entries(TARGET_LABEL_MATCHERS)) {
    if (matchers.some(m => new RegExp(`\\b${escapeRegex(m)}\\b`).test(v))) return label;
  }
  return null;
};

function parseTargetXlsx(wb: XLSX.WorkBook) {
  const targets: any = { India: { Q1: {}, Q2: {}, Q3: {}, Q4: {} }, US: { Q1: {}, Q2: {}, Q3: {}, Q4: {} } };
  let extractedCount = 0;
  const extractedRows: any[] = [];

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    let headerRowIdx = -1; let qCol: any = {};
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      const cells = row.map((c: any) => String(c || '').toLowerCase().trim());
      const findQ = (qStr: string, monthA: string, monthB: string) => cells.findIndex((c: string) =>
        new RegExp(`\\b${qStr}\\b`).test(c) || (c.includes(monthA) && c.includes(monthB))
      );
      const q1 = findQ('q1', 'apr', 'jun');
      const q2 = findQ('q2', 'jul', 'sep');
      const q3 = findQ('q3', 'oct', 'dec');
      const q4 = findQ('q4', 'jan', 'mar');
      if (q1 >= 0 && q2 >= 0 && q3 >= 0 && q4 >= 0) {
        headerRowIdx = i;
        qCol = { Q1: q1, Q2: q2, Q3: q3, Q4: q4 };
        break;
      }
    }
    if (headerRowIdx < 0) continue;

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const row = rows[i] || [];
      if (row.every((c: any) => c === '' || c == null)) continue;

      let verticalText = '';
      const minQCol = Math.min(...(Object.values(qCol) as number[]));
      for (let j = 0; j < minQCol; j++) {
        const t = String(row[j] || '').trim();
        if (t) { verticalText = t; break; }
      }
      if (!verticalText) continue;

      const lvt = verticalText.toLowerCase();
      if (lvt.includes('sub-total') || lvt.includes('subtotal') || lvt === 'total' ||
          lvt.includes('grand total') || lvt.startsWith('*') || lvt.includes('note') ||
          lvt.includes('full year')) continue;

      const label = targetLabelFromText(verticalText);
      if (!label) continue;

      const region = isUSLabel(label) ? 'US' : 'India';
      const quarterValues: any = {};
      let rowHasValue = false;
      for (const q of ['Q1', 'Q2', 'Q3', 'Q4']) {
        const raw = row[qCol[q]];
        if (raw === '' || raw == null) continue;
        const val = parseFloat(String(raw).replace(/[^0-9.\-]/g, ''));
        if (isNaN(val) || val <= 0) continue;
        targets[region][q][label] = val;
        quarterValues[q] = val;
        rowHasValue = true;
        extractedCount++;
      }
      if (rowHasValue) extractedRows.push({ source: verticalText, label, region, ...quarterValues });
    }
  }
  return { targets, extractedCount, extractedRows };
}

const getPeriodTarget = (targets: any, region: string, period: string, vert: string, fy: number, cq: string) => {
  if (!targets[region]) return 0;
  const [nq] = advanceQ(cq, fy);
  const [hq1] = advanceQ(nq, fy);
  const [hq2] = advanceQ(hq1, fy);
  if (period === 'thisQ') return targets[region][cq]?.[vert] || 0;
  if (period === 'nextQ') return targets[region][nq]?.[vert] || 0;
  if (period === 'nextH') return (targets[region][hq1]?.[vert] || 0) + (targets[region][hq2]?.[vert] || 0);
  return 0;
};

const computeAchievedByVert = (deals: any[], periodStart: Date, periodEnd: Date, verticalList: string[]) => {
  const out: any = {};
  for (const vert of verticalList) {
    const won = deals.filter(d =>
      d.resolvedLabel === vert &&
      d.stage === 'Won' && d.status === 'Won' &&
      d.mrrStartDate && d.mrrStartDate >= periodStart && d.mrrStartDate <= periodEnd
    );
    out[vert] = won.reduce((a, d) => a + d.mrr, 0);
  }
  return out;
};

// ==================== Pills ====================
function StagePill({ stage }: { stage: string }) {
  const cls = stage === 'Won' ? 'bg-emerald-100 text-emerald-800 border-emerald-200' :
              stage === 'Lost' ? 'bg-rose-100 text-rose-800 border-rose-200' :
              PREPIPE_STAGES.includes(stage) ? 'bg-violet-100 text-violet-800 border-violet-200' :
              PIPELINE_STAGES.includes(stage) ? 'bg-sky-100 text-sky-800 border-sky-200' :
              'bg-slate-100 text-slate-700 border-slate-200';
  return <span className={`px-1.5 py-0.5 rounded text-[10px] border ${cls} whitespace-nowrap`}>{stage || '—'}</span>;
}
function StatusPill({ status }: { status: string }) {
  const cls = status === 'Won' ? 'bg-emerald-100 text-emerald-800' :
              status === 'Lost' ? 'bg-rose-100 text-rose-800' :
              'bg-slate-100 text-slate-700';
  return <span className={`px-1.5 py-0.5 rounded text-[10px] ${cls}`}>{status}</span>;
}

// ==================== MultiSelect Filter ====================
function MultiSelectFilter({ label, options, selected, onChange }: any) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState('');
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) { setOpen(false); setSearch(''); } };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  const filtered = options.filter((o: string) => !search || String(o).toLowerCase().includes(search.toLowerCase()));
  const toggle = (opt: string) => {
    if (selected.includes(opt)) onChange(selected.filter((s: string) => s !== opt));
    else onChange([...selected, opt]);
  };

  return (
    <div ref={ref} className="relative">
      <button onClick={() => setOpen(!open)}
        className={`px-2.5 py-1.5 text-sm border rounded transition flex items-center gap-1.5 whitespace-nowrap ${
          selected.length > 0 ? 'border-sky-500 bg-sky-50 text-sky-900 font-medium' : 'border-slate-300 bg-white text-slate-700 hover:border-slate-400'
        }`}>
        <span>{label}</span>
        {selected.length > 0 && <span className="bg-sky-600 text-white text-[10px] font-bold rounded-full px-1.5 min-w-[18px] text-center">{selected.length}</span>}
        <ChevronDown className={`w-3 h-3 transition ${open ? 'rotate-180' : ''}`} />
      </button>
      {open && (
        <div className="absolute top-full left-0 mt-1 bg-white border border-slate-300 rounded shadow-lg z-30 w-64 max-h-80 flex flex-col">
          <div className="p-2 border-b border-slate-200">
            <input type="text" autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder={`Search ${label.toLowerCase()}...`}
              className="w-full px-2 py-1 text-xs border border-slate-300 rounded focus:outline-none focus:border-sky-500" />
          </div>
          <div className="overflow-y-auto flex-1 py-1">
            {filtered.length === 0 ? (
              <div className="px-3 py-2 text-xs text-slate-400 italic">No matches</div>
            ) : filtered.map((opt: string) => (
              <label key={opt} className="flex items-center gap-2 px-3 py-1 hover:bg-slate-50 cursor-pointer text-xs">
                <input type="checkbox" checked={selected.includes(opt)} onChange={() => toggle(opt)}
                  className="w-3.5 h-3.5 text-sky-600 rounded border-slate-300 focus:ring-0 focus:ring-offset-0" />
                <span className="truncate text-slate-900" title={opt}>{opt || '(blank)'}</span>
              </label>
            ))}
          </div>
          {selected.length > 0 && (
            <button onClick={() => onChange([])} className="p-1.5 text-[11px] text-rose-700 hover:bg-rose-50 border-t border-slate-200 font-medium">
              Clear {label}
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function DateRangeFilter({ label, value, onChange }: any) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const from = value?.from || '';
  const to   = value?.to   || '';
  const active = from || to;

  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  const setFrom = (v: string) => onChange({ from: v, to });
  const setTo   = (v: string) => onChange({ from, to: v });
  const clear   = () => { onChange({ from: '', to: '' }); setOpen(false); };

  const displayLabel = active ? `${from || '…'} → ${to || '…'}` : label;

  return (
    <div ref={ref} className="relative">
      <button onClick={() => setOpen(!open)}
        className={`px-2.5 py-1.5 text-sm border rounded transition flex items-center gap-1.5 whitespace-nowrap ${
          active ? 'border-sky-500 bg-sky-50 text-sky-900 font-medium' : 'border-slate-300 bg-white text-slate-700 hover:border-slate-400'
        }`}>
        <span className="max-w-[180px] truncate">{displayLabel}</span>
        <ChevronDown className={`w-3 h-3 flex-shrink-0 transition ${open ? 'rotate-180' : ''}`} />
      </button>
      {open && (
        <div className="absolute top-full left-0 mt-1 bg-white border border-slate-300 rounded shadow-lg z-30 w-56 p-3 space-y-2">
          <div className="text-[10px] uppercase tracking-wider text-slate-500 font-semibold">{label}</div>
          <div>
            <div className="text-[11px] text-slate-500 mb-0.5">From</div>
            <input type="date" value={from} onChange={e => setFrom(e.target.value)}
              className="w-full text-xs px-2 py-1 border border-slate-300 rounded focus:outline-none focus:border-sky-500" />
          </div>
          <div>
            <div className="text-[11px] text-slate-500 mb-0.5">To</div>
            <input type="date" value={to} onChange={e => setTo(e.target.value)}
              className="w-full text-xs px-2 py-1 border border-slate-300 rounded focus:outline-none focus:border-sky-500" />
          </div>
          {active && (
            <button onClick={clear} className="w-full text-[11px] text-rose-700 hover:bg-rose-50 rounded py-1 font-medium border-t border-slate-200">
              Clear {label}
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function EditableTarget({ value, onChange, region }: any) {
  const [editing, setEditing] = useState(false);
  const [local, setLocal] = useState(value || 0);
  useEffect(() => { if (!editing) setLocal(value || 0); }, [value, editing]);

  if (editing) {
    return (
      <input type="text" inputMode="numeric" autoFocus value={local}
        onChange={(e) => { const v = e.target.value.replace(/[^0-9]/g, ''); setLocal(v ? +v : 0); }}
        onBlur={() => { onChange(local); setEditing(false); }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { onChange(local); setEditing(false); }
          if (e.key === 'Escape') { setLocal(value || 0); setEditing(false); }
        }}
        className="w-28 bg-white border border-sky-500 ring-2 ring-sky-100 rounded px-1.5 py-0.5 text-right text-sm tabular-nums focus:outline-none"
      />
    );
  }
  return (
    <button onClick={() => setEditing(true)} title="Click to edit"
      className="group flex items-center gap-1 hover:bg-slate-100 rounded px-1.5 py-0.5 transition w-full justify-end tabular-nums">
      <span className={value ? 'text-slate-900' : 'text-slate-400 italic'}>{value ? fmt(value, region) : 'Set…'}</span>
      <Edit3 className="w-3 h-3 text-slate-400 opacity-0 group-hover:opacity-100 transition flex-shrink-0" />
    </button>
  );
}

function MetricCell({ value, deals, tone, isCount, onClick, region }: any) {
  const display = isCount ? (value || 0).toLocaleString('en-IN') : fmt(value, region);
  const toneCls = !value ? 'text-slate-400' : (({
    positive: 'text-emerald-700',
    negative: 'text-rose-700',
    pipeline: 'text-sky-700',
    lapsed: 'text-amber-700',
    available: 'text-cyan-700',
    prepipe: 'text-violet-700',
    carry: 'text-orange-700',
    neutral: 'text-slate-900',
  } as Record<string, string>)[tone] || 'text-slate-900');
  const clickable = deals && deals.length > 0;
  if (!clickable) return <span className={`tabular-nums ${toneCls}`}>{display}</span>;
  return (
    <button onClick={onClick} className={`tabular-nums ${toneCls} hover:underline decoration-dotted underline-offset-2 font-medium`}>
      {display}
    </button>
  );
}

// ==================== Sortable Header Cell ====================
function SortHeader({ column, label, sort, onSort, align = 'left' }: any) {
  const active = sort?.column === column;
  const Icon = !active ? ArrowUpDown : (sort.direction === 'asc' ? ArrowUp : ArrowDown);
  return (
    <button onClick={() => onSort(column)}
      className={`flex items-center gap-1 hover:text-slate-900 transition ${align === 'right' ? 'ml-auto' : ''} ${active ? 'text-sky-700' : 'text-slate-600'}`}>
      <span>{label}</span>
      <Icon className={`w-3 h-3 ${active ? 'opacity-100' : 'opacity-40'}`} />
    </button>
  );
}

function DrillSortHeader({ col, label, sort, onSort, align = 'left' }: any) {
  const active = sort.col === col;
  const Icon = !active ? ArrowUpDown : (sort.dir === 'asc' ? ArrowUp : ArrowDown);
  return (
    <button
      onClick={() => onSort((s: any) => s.col === col ? { col, dir: s.dir === 'asc' ? 'desc' : 'asc' } : { col, dir: 'desc' })}
      className={`flex items-center gap-0.5 hover:text-slate-900 transition whitespace-nowrap ${align === 'right' ? 'ml-auto' : ''} ${active ? 'text-sky-700' : 'text-slate-600'}`}>
      <span>{label}</span><Icon className={`w-3 h-3 ${active ? 'opacity-100' : 'opacity-40'}`} />
    </button>
  );
}

function sortRows(rows: any[], sort: any) {
  if (!rows || !rows.length || !sort?.col) return rows || [];
  const { col, dir } = sort;
  const mult = dir === 'asc' ? 1 : -1;
  const arr = [...rows];
  arr.sort((a, b) => {
    const av = a[col]; const bv = b[col];
    if (av == null && bv == null) return 0;
    if (av == null) return 1;
    if (bv == null) return -1;
    if (av instanceof Date && bv instanceof Date) return mult * (av.getTime() - bv.getTime());
    if (typeof av === 'number' && typeof bv === 'number') return mult * (av - bv);
    return mult * String(av).toLowerCase().localeCompare(String(bv).toLowerCase());
  });
  return arr;
}

// ==================== Sortable Deal Table ====================
function SortableDealTable({ deals, onRowClick, region, sort, onSort }: any) {
  const colW = ['9%','9%','7%','7%','7%','7%','7%','6%','6%','7%','7%','7%','7%','7%'];
  return (
    <table className="w-full text-xs table-fixed">
      <colgroup>{colW.map((w, i) => <col key={i} style={{ width: w }} />)}</colgroup>
      <thead className="bg-slate-50 text-[10px] uppercase tracking-wider sticky top-0 z-10">
        <tr>
          <th className="px-2 py-1.5 text-left font-semibold">Account</th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="name" label="Deal" sort={sort} onSort={onSort} /></th>
          <th className="px-2 py-1.5 text-left font-semibold">Current Owner</th>
          <th className="px-2 py-1.5 text-left font-semibold">SD Initiator</th>
          <th className="px-2 py-1.5 text-left font-semibold">SA Initiator</th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="resolvedLabel" label="Vertical" sort={sort} onSort={onSort} /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="stage" label="Stage" sort={sort} onSort={onSort} /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="status" label="Status" sort={sort} onSort={onSort} /></th>
          <th className="px-2 py-1.5 text-left font-semibold">Emp. Range</th>
          <th className="px-2 py-1.5 text-right font-semibold"><SortHeader column="mrr" label="MRR" sort={sort} onSort={onSort} align="right" /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="pipeAddedDate" label="Pipe Added" sort={sort} onSort={onSort} /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="expClosureDate" label="Exp. Close" sort={sort} onSort={onSort} /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="closureDate" label="Final Close" sort={sort} onSort={onSort} /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SortHeader column="mrrStartDate" label="MRR Start" sort={sort} onSort={onSort} /></th>
        </tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
        {deals.map((d: any) => {
          const dealRegion = region || (US_VERTICALS.includes(d.resolvedLabel) ? 'US' : 'India');
          return (
            <tr key={d._id} onClick={() => onRowClick && onRowClick(d)} className="hover:bg-slate-50 cursor-pointer">
              <td className="px-2 py-1.5 overflow-hidden"><div className="text-[10px] text-slate-700 font-medium truncate" title={d.account}>{d.account || '—'}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="font-medium text-slate-900 truncate text-xs" title={d.name}>{d.name || '—'}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="text-[10px] text-slate-700 truncate" title={d.currentOwner}>{d.currentOwner || '—'}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="text-[10px] text-slate-700 truncate" title={d.sdInitiator}>{d.sdInitiator || '—'}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="text-[10px] text-slate-700 truncate" title={d.saInitiator}>{d.saInitiator || '—'}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="text-[10px] font-semibold text-slate-900 truncate">{d.resolvedLabel || <span className="text-amber-700">unmapped</span>}</div></td>
              <td className="px-2 py-1.5"><StagePill stage={d.stage} /></td>
              <td className="px-2 py-1.5"><StatusPill status={d.status} /></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="text-[10px] text-slate-600 truncate" title={d.employeeSize}>{d.employeeSize || '—'}</div></td>
              <td className="px-2 py-1.5 text-right font-medium tabular-nums text-slate-900 whitespace-nowrap overflow-hidden">{fmt(d.mrr, dealRegion)}</td>
              <td className="px-2 py-1.5 text-slate-600 tabular-nums whitespace-nowrap text-[10px]">{fmtDate(d.pipeAddedDate)}</td>
              <td className="px-2 py-1.5 text-slate-600 tabular-nums whitespace-nowrap text-[10px]">{fmtDate(d.expClosureDate)}</td>
              <td className="px-2 py-1.5 text-slate-600 tabular-nums whitespace-nowrap text-[10px]">{fmtDate(d.closureDate)}</td>
              <td className="px-2 py-1.5 text-slate-600 tabular-nums whitespace-nowrap text-[10px]">{fmtDate(d.mrrStartDate)}</td>
            </tr>
          );
        })}
        {deals.length === 0 && (
          <tr><td colSpan={14} className="px-4 py-8 text-center text-slate-400 text-xs">No deals match.</td></tr>
        )}
      </tbody>
    </table>
  );
}

// ==================== Pagination Controls ====================
function PaginationControls({ page, pageSize, total, onPageChange, onPageSizeChange }: any) {
  const pageCount = Math.max(1, Math.ceil(total / pageSize));
  const safePage = Math.min(Math.max(0, page), pageCount - 1);
  const start = total === 0 ? 0 : safePage * pageSize + 1;
  const end = Math.min(total, (safePage + 1) * pageSize);

  const btn = (props: any, children: any) => (
    <button {...props}
      className="px-2 py-1 text-xs rounded bg-slate-100 hover:bg-slate-200 disabled:opacity-40 disabled:cursor-not-allowed text-slate-700 flex items-center">
      {children}
    </button>
  );

  return (
    <div className="px-4 py-3 border-t border-slate-200 flex flex-wrap items-center gap-3 justify-between text-sm">
      <div className="text-xs text-slate-600 tabular-nums">
        Showing <span className="font-semibold text-slate-900">{start.toLocaleString()}–{end.toLocaleString()}</span> of <span className="font-semibold text-slate-900">{total.toLocaleString()}</span>
      </div>
      <div className="flex items-center gap-2">
        <span className="text-xs text-slate-500">Per page:</span>
        {[20, 50, 100].map(s => (
          <button key={s} onClick={() => onPageSizeChange(s)}
            className={`px-2 py-1 text-xs rounded font-medium ${pageSize === s ? 'bg-sky-600 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'}`}>{s}</button>
        ))}
      </div>
      <div className="flex items-center gap-1">
        {btn({ disabled: safePage === 0, onClick: () => onPageChange(0), title: 'First page' }, <ChevronsLeft className="w-3.5 h-3.5" />)}
        {btn({ disabled: safePage === 0, onClick: () => onPageChange(safePage - 1), title: 'Previous' }, <ChevronLeft className="w-3.5 h-3.5" />)}
        <span className="px-3 text-xs text-slate-700 tabular-nums">Page <span className="font-semibold text-slate-900">{safePage + 1}</span> of <span className="font-semibold text-slate-900">{pageCount}</span></span>
        {btn({ disabled: safePage >= pageCount - 1, onClick: () => onPageChange(safePage + 1), title: 'Next' }, <ChevronRight className="w-3.5 h-3.5" />)}
        {btn({ disabled: safePage >= pageCount - 1, onClick: () => onPageChange(pageCount - 1), title: 'Last page' }, <ChevronsRight className="w-3.5 h-3.5" />)}
      </div>
    </div>
  );
}

// ==================== Active Filter Chips ====================
function ActiveFilterChips({ filters, onRemove }: any) {
  const chips: any[] = [];
  for (const [key, values] of Object.entries(filters)) {
    if (Array.isArray(values)) {
      if (!values.length) continue;
      for (const value of values) chips.push({ key, value, label: FILTER_LABELS[key] || key, isDate: false });
    } else if (typeof values === 'object' && values !== null && ((values as any).from || (values as any).to)) {
      chips.push({ key, value: `${(values as any).from || '…'} → ${(values as any).to || '…'}`, label: FILTER_LABELS[key] || key, isDate: true });
    }
  }
  if (chips.length === 0) return null;
  return (
    <div className="flex flex-wrap items-center gap-1.5 px-4 py-2 bg-sky-50/60 border-b border-slate-200">
      <span className="text-[11px] uppercase tracking-wider text-slate-600 font-semibold mr-1">Filters:</span>
      {chips.map((c, i) => (
        <span key={i} className="inline-flex items-center gap-1 text-[11px] bg-white border border-sky-300 text-sky-900 rounded px-1.5 py-0.5">
          <span className="text-slate-500">{c.label}:</span>
          <span className="font-medium">{c.value || '(blank)'}</span>
          {onRemove && (
            <button onClick={() => onRemove(c.key, c.isDate ? null : c.value)} className="text-slate-400 hover:text-rose-600 ml-0.5">
              <X className="w-3 h-3" />
            </button>
          )}
        </span>
      ))}
    </div>
  );
}

// ==================== Drill Deal Table ====================
function DrillDealTable({ deals, region, onRowClick }: any) {
  const [sort, setSort] = useState({ column: 'mrr', direction: 'desc' });
  const handleSort = (col: string) => setSort(s => s.column === col ? { column: col, direction: s.direction === 'asc' ? 'desc' : 'asc' } : { column: col, direction: 'desc' });
  const sorted = useMemo(() => sortDeals(deals, sort), [deals, sort]);

  const SH = ({ col, label, align = 'left' }: any) => {
    const active = sort.column === col;
    const Icon = !active ? ArrowUpDown : (sort.direction === 'asc' ? ArrowUp : ArrowDown);
    return (
      <button onClick={() => handleSort(col)} className={`flex items-center gap-0.5 hover:text-slate-900 transition whitespace-nowrap ${align === 'right' ? 'ml-auto' : ''} ${active ? 'text-sky-700' : 'text-slate-600'}`}>
        <span>{label}</span><Icon className={`w-3 h-3 ${active ? 'opacity-100' : 'opacity-40'}`} />
      </button>
    );
  };

  const colW = ['9%','9%','7%','7%','7%','7%','7%','6%','6%','7%','7%','7%','7%','7%'];

  return (
    <table className="w-full text-xs table-fixed">
      <colgroup>{colW.map((w, i) => <col key={i} style={{ width: w }} />)}</colgroup>
      <thead className="bg-slate-50 text-[10px] uppercase tracking-wider sticky top-0 z-10">
        <tr>
          <th className="px-2 py-1.5 text-left font-semibold">Account</th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="name" label="Deal" /></th>
          <th className="px-2 py-1.5 text-left font-semibold">Current Owner</th>
          <th className="px-2 py-1.5 text-left font-semibold">SD Initiator</th>
          <th className="px-2 py-1.5 text-left font-semibold">SA Initiator</th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="resolvedLabel" label="Vertical" /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="stage" label="Stage" /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="status" label="Status" /></th>
          <th className="px-2 py-1.5 text-left font-semibold">Emp. Range</th>
          <th className="px-2 py-1.5 text-right font-semibold"><SH col="mrr" label="MRR" align="right" /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="pipeAddedDate" label="Pipe Added" /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="expClosureDate" label="Exp. Close" /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="closureDate" label="Final Close" /></th>
          <th className="px-2 py-1.5 text-left font-semibold"><SH col="mrrStartDate" label="MRR Start" /></th>
        </tr>
      </thead>
      <tbody className="divide-y divide-slate-100">
        {sorted.map((d: any) => {
          const dr = region || (US_VERTICALS.includes(d.resolvedLabel) ? 'US' : 'India');
          const isMrr = !!d._isMrrEntry;
          const rowCls = isMrr
            ? 'bg-emerald-50/60 hover:bg-emerald-50 cursor-pointer'
            : 'hover:bg-slate-50 cursor-pointer';
          return (
            <tr key={d._id} onClick={() => onRowClick && onRowClick(d)} className={rowCls}>
              <td className="px-2 py-1.5 overflow-hidden">
                <div className="truncate text-[10px] text-slate-700 font-medium" title={d.account}>{d.account || '—'}</div>
                {isMrr && <div className="text-[9px] text-emerald-700 font-semibold">MRR File</div>}
              </td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="truncate font-medium text-slate-900 text-xs" title={d.name}>{d.name || '—'}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] text-slate-700" title={d.currentOwner}>{d.currentOwner || '—'}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] text-slate-500 italic">{isMrr ? 'n/a' : (d.sdInitiator || '—')}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] text-slate-500 italic">{isMrr ? 'n/a' : (d.saInitiator || '—')}</div></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] font-semibold text-slate-900">{d.resolvedLabel || <span className="text-amber-700">unmapped</span>}</div></td>
              <td className="px-2 py-1.5"><StagePill stage={d.stage} /></td>
              <td className="px-2 py-1.5"><StatusPill status={d.status} /></td>
              <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] text-slate-600">{isMrr ? '—' : (d.employeeSize || '—')}</div></td>
              <td className="px-2 py-1.5 text-right font-medium tabular-nums text-emerald-700 whitespace-nowrap text-xs">{fmt(d.mrr, dr)}</td>
              <td className="px-2 py-1.5 tabular-nums text-[10px] text-slate-500 whitespace-nowrap">{isMrr ? '—' : fmtDate(d.pipeAddedDate)}</td>
              <td className="px-2 py-1.5 tabular-nums text-[10px] text-slate-500 whitespace-nowrap">{isMrr ? '—' : fmtDate(d.expClosureDate)}</td>
              <td className="px-2 py-1.5 tabular-nums text-[10px] text-slate-600 whitespace-nowrap">{fmtDate(d.closureDate)}</td>
              <td className="px-2 py-1.5 tabular-nums text-[10px] text-slate-600 whitespace-nowrap">{fmtDate(d.mrrStartDate)}</td>
            </tr>
          );
        })}
        {sorted.length === 0 && (
          <tr><td colSpan={14} className="px-4 py-8 text-center text-slate-400 text-xs">No deals match.</td></tr>
        )}
      </tbody>
    </table>
  );
}

// ==================== Modals ====================
function DrillModal({ drill, region, activeFilters, onClose, onRowClick, onRemoveFilter }: any) {
  const total = drill.deals.reduce((a: number, d: any) => a + (d.mrr || 0), 0);

  return (
    <div className="fixed inset-0 bg-slate-900/50 z-50 p-4 flex items-center justify-center" onClick={onClose}>
      <div className="bg-white border border-slate-200 rounded-lg shadow-2xl flex flex-col w-full max-w-[95vw]" style={{ maxHeight: 'calc(100vh - 2rem)' }} onClick={(e) => e.stopPropagation()}>
        <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between flex-shrink-0">
          <div>
            <h3 className="font-semibold text-slate-900 text-sm">{drill.title}</h3>
            <p className="text-xs text-slate-600 mt-0.5">
              {drill.deals.length.toLocaleString()} {drill.deals.length === 1 ? 'deal' : 'deals'} • Total MRR: <span className="text-slate-900 font-medium">{fmt(total, region)}</span>
              {drill.deals.some((d: any) => d._isMrrEntry) && (
                <span className="ml-3 inline-flex items-center gap-1 text-emerald-700 font-medium">
                  <span className="w-2.5 h-2.5 rounded-sm bg-emerald-200 inline-block" /> green rows = MRR Transaction file entries
                </span>
              )}
            </p>
          </div>
          <button onClick={onClose} className="text-slate-500 hover:text-slate-900 p-1 rounded hover:bg-slate-100"><X className="w-4 h-4" /></button>
        </div>
        <ActiveFilterChips filters={activeFilters} onRemove={onRemoveFilter} />
        <div className="flex-1 overflow-auto min-h-0">
          <DrillDealTable deals={drill.deals} region={region} onRowClick={onRowClick} />
        </div>
      </div>
    </div>
  );
}

function DealDetailModal({ deal, onClose }: any) {
  if (!deal) return null;
  const dealRegion = US_VERTICALS.includes(deal.resolvedLabel) ? 'US' : 'India';
  const fields = Object.entries(deal._raw || {}).filter(([, v]) => v != null && String(v).trim() !== '');
  return (
    <div className="fixed inset-0 bg-slate-900/60 z-[60] p-4 flex items-center justify-center" onClick={onClose}>
      <div className="bg-white border border-slate-200 rounded-lg shadow-2xl flex flex-col w-full max-w-4xl" style={{ maxHeight: 'calc(100vh - 2rem)' }} onClick={(e) => e.stopPropagation()}>
        <div className="px-4 py-3 border-b border-slate-200 flex items-start justify-between gap-3 flex-shrink-0">
          <div className="min-w-0 flex-1">
            <h3 className="font-semibold text-slate-900 truncate text-sm">{deal.name || 'Deal Details'}</h3>
            <div className="flex flex-wrap items-center gap-x-2 gap-y-1 mt-1">
              {deal.account && <span className="text-[11px] text-slate-600">{deal.account}</span>}
              <StagePill stage={deal.stage} />
              <StatusPill status={deal.status} />
              <span className="text-[11px] font-medium text-emerald-700">{fmt(deal.mrr, dealRegion)}</span>
              <span className="text-[11px] text-slate-600">{deal.resolvedLabel || 'unmapped'}</span>
            </div>
          </div>
          <button onClick={onClose} className="text-slate-500 hover:text-slate-900 p-1 rounded hover:bg-slate-100 flex-shrink-0"><X className="w-4 h-4" /></button>
        </div>
        <div className="flex-1 overflow-auto min-h-0 p-3">
          <div className="text-[11px] text-slate-600 mb-2 flex items-center gap-1">
            <Info className="w-3 h-3" /> All fields from source ({fields.length})
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
            {fields.map(([k, v]) => (
              <div key={k} className="border border-slate-200 rounded p-2 bg-slate-50/50">
                <div className="text-[10px] uppercase tracking-wider text-slate-500 mb-0.5">{k}</div>
                <div className="text-xs text-slate-900 break-words">{String(v)}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ==================== This Quarter Table ====================
function ThisQuarterTable({ subtitle, accent, deals, achievedDeals, fy, q, region, verticalList, targets, onTargetChange, onDrill }: any) {
  const [qs, qe] = quarterRange(fy, q);
  const now = today();

  const rows = verticalList.map((vert: string) => {
    const vd = deals.filter((d: any) => d.resolvedLabel === vert);
    const avd = (achievedDeals || []).filter((d: any) => d.resolvedLabel === vert);
    const wonDeals = avd.filter((d: any) => d._isMrrEntry && d.mrrStartDate && d.mrrStartDate >= qs && d.mrrStartDate <= qe);

    const openDeals = vd.filter((d: any) =>
      d.status !== 'Won' && d.status !== 'Lost' &&
      d.stage !== 'Won' && d.stage !== 'Lost' &&
      d.pipeAddedDate &&
      d.expClosureDate && d.expClosureDate >= qs && d.expClosureDate <= qe
    );

    const lostDeals = vd.filter((d: any) =>
      d.stage === 'Lost' && d.status === 'Lost' &&
      d.expClosureDate && d.expClosureDate >= qs && d.expClosureDate <= qe &&
      d.closureDate && d.closureDate >= qs && d.closureDate <= qe
    );

    const lapsedDeals = openDeals.filter((d: any) => d.expClosureDate && d.expClosureDate < now);
    const availableDeals = openDeals;

    const target = targets[vert] || 0;
    const achieved = wonDeals.reduce((a: number, d: any) => a + d.mrr, 0);
    return {
      vert, target, achieved, wonDeals,
      gap: target - achieved,
      open: openDeals.reduce((a: number, d: any) => a + d.mrr, 0), openDeals,
      lost: lostDeals.reduce((a: number, d: any) => a + d.mrr, 0), lostDeals,
      lapsed: lapsedDeals.reduce((a: number, d: any) => a + d.mrr, 0), lapsedDeals,
      available: availableDeals.reduce((a: number, d: any) => a + d.mrr, 0), availableDeals,
    };
  });

  const totals = rows.reduce((a: any, r: any) => ({
    target: a.target + r.target, achieved: a.achieved + r.achieved, gap: a.gap + r.gap,
    open: a.open + r.open, lost: a.lost + r.lost, lapsed: a.lapsed + r.lapsed, available: a.available + r.available,
  }), { target: 0, achieved: 0, gap: 0, open: 0, lost: 0, lapsed: 0, available: 0 });

  const totalAllDeals = (key: string) => rows.flatMap((r: any) => r[key + 'Deals'] || []);

  return (
    <div className="bg-white border border-slate-200 rounded-lg overflow-hidden shadow-sm">
      <div className={`px-4 py-3 border-b border-slate-200 bg-gradient-to-r ${accent}`}>
        <div className="text-xs uppercase tracking-wider text-slate-600 font-semibold">This Quarter — {region}</div>
        <div className="text-base font-bold text-slate-900 mt-0.5">{subtitle}</div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm table-fixed">
          <colgroup>
            <col style={{ width: '18%' }} /><col style={{ width: '14%' }} /><col style={{ width: '14%' }} />
            <col style={{ width: '13%' }} /><col style={{ width: '14%' }} /><col style={{ width: '13%' }} /><col style={{ width: '14%' }} />
          </colgroup>
          <thead className="bg-slate-50 text-[11px] uppercase tracking-wider text-slate-600">
            <tr>
              <th className="px-3 py-2.5 text-left font-semibold">Vertical</th>
              <th className="px-3 py-2.5 text-right font-semibold">MRR Target</th>
              <th className="px-3 py-2.5 text-right font-semibold leading-tight">MRR Achieved<br/><span className="text-slate-400 normal-case font-normal text-[10px]">(Start MRR)</span></th>
              <th className="px-3 py-2.5 text-right font-semibold">MRR Gap</th>
              <th className="px-3 py-2.5 text-right font-semibold">Open Pipe</th>
              <th className="px-3 py-2.5 text-right font-semibold">Lost Pipe</th>
              <th className="px-3 py-2.5 text-right font-semibold">Lapsed Pipe</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {rows.map((r: any) => (
              <tr key={r.vert} className="hover:bg-slate-50">
                <td className="px-3 py-2 font-semibold text-slate-900">{r.vert}</td>
                <td className="px-3 py-2 text-right"><EditableTarget value={r.target} onChange={(v: number) => onTargetChange(r.vert, v)} region={region} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.achieved} deals={r.wonDeals} tone="positive" region={region} onClick={() => onDrill(`MRR Achieved — ${r.vert} — ${q} ${fyLabel(fy)}`, r.wonDeals, region)} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.gap} tone={r.gap > 0 ? 'negative' : 'positive'} region={region} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.open} deals={r.openDeals} tone="pipeline" region={region} onClick={() => onDrill(`Open Pipe — ${r.vert} — Exp. Closure in ${q} ${fyLabel(fy)}`, r.openDeals, region)} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.lost} deals={r.lostDeals} tone="negative" region={region} onClick={() => onDrill(`Lost Pipe — ${r.vert} — Exp. & Final Closure in ${q} ${fyLabel(fy)}`, r.lostDeals, region)} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.lapsed} deals={r.lapsedDeals} tone="lapsed" region={region} onClick={() => onDrill(`Lapsed Pipe — ${r.vert} — Exp. closure this Q, past today`, r.lapsedDeals, region)} /></td>
              </tr>
            ))}
            <tr className="bg-slate-100 font-semibold border-t-2 border-slate-300">
              <td className="px-3 py-2.5">Total</td>
              <td className="px-3 py-2.5 text-right tabular-nums">{fmt(totals.target, region)}</td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.achieved} deals={totalAllDeals('won')} tone="positive" region={region} onClick={() => onDrill(`MRR Achieved — All ${region}`, totalAllDeals('won'), region)} /></td>
              <td className={`px-3 py-2.5 text-right tabular-nums ${totals.gap > 0 ? 'text-rose-700' : 'text-emerald-700'}`}>{fmt(totals.gap, region)}</td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.open} deals={totalAllDeals('open')} tone="pipeline" region={region} onClick={() => onDrill(`Open Pipe — All ${region}`, totalAllDeals('open'), region)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.lost} deals={totalAllDeals('lost')} tone="negative" region={region} onClick={() => onDrill(`Lost Pipe — All ${region}`, totalAllDeals('lost'), region)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.lapsed} deals={totalAllDeals('lapsed')} tone="lapsed" region={region} onClick={() => onDrill(`Lapsed Pipe — All ${region}`, totalAllDeals('lapsed'), region)} /></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ==================== Forecast Table ====================
function ForecastTable({ title, subtitle, accent, deals, achievedDeals, periodStart, periodEnd, qLabel, region, verticalList, targets, carryForward, prevPeriodLabel, onTargetChange, onDrill }: any) {
  const rows = verticalList.map((vert: string) => {
    const vd = deals.filter((d: any) => d.resolvedLabel === vert);
    const avd = (achievedDeals || []).filter((d: any) => d.resolvedLabel === vert);
    const wonDeals = avd.filter((d: any) => d._isMrrEntry && d.mrrStartDate && d.mrrStartDate >= periodStart && d.mrrStartDate <= periodEnd);
    const dueDeals = vd.filter((d: any) =>
      d.status !== 'Won' && d.status !== 'Lost' &&
      d.stage !== 'Won' && d.stage !== 'Lost' &&
      d.pipeAddedDate &&
      d.expClosureDate && d.expClosureDate >= periodStart && d.expClosureDate <= periodEnd
    );
    const lostDeals = vd.filter((d: any) =>
      (d.stage === 'Lost' || d.status === 'Lost') &&
      d.closureDate && d.closureDate >= periodStart && d.closureDate <= periodEnd
    );
    const prepipeDeals = vd.filter((d: any) => PREPIPE_STAGES.includes(d.stage) && d.status !== 'Won' && d.status !== 'Lost');

    const target = targets[vert] || 0;
    const carry = (carryForward && carryForward[vert]) || 0;
    const achieved = wonDeals.reduce((a: number, d: any) => a + d.mrr, 0);
    const gap = target + carry - achieved;
    return {
      vert, carry, target, achieved, wonDeals, gap,
      due: dueDeals.reduce((a: number, d: any) => a + d.mrr, 0), dueDeals,
      lost: lostDeals.reduce((a: number, d: any) => a + d.mrr, 0), lostDeals,
      prepipeCount: prepipeDeals.length, prepipeDeals,
      prepipeMrr: prepipeDeals.reduce((a: number, d: any) => a + d.mrr, 0),
    };
  });

  const totals = rows.reduce((a: any, r: any) => ({
    carry: a.carry + r.carry,
    target: a.target + r.target, achieved: a.achieved + r.achieved, gap: a.gap + r.gap,
    due: a.due + r.due, lost: a.lost + r.lost, prepipeCount: a.prepipeCount + r.prepipeCount, prepipeMrr: a.prepipeMrr + r.prepipeMrr,
  }), { carry: 0, target: 0, achieved: 0, gap: 0, due: 0, lost: 0, prepipeCount: 0, prepipeMrr: 0 });

  const totalAllDeals = (key: string) => rows.flatMap((r: any) => r[key + 'Deals'] || []);
  const colW = ['13%', '11%', '11%', '11%', '11%', '11%', '11%', '10%', '11%'];

  return (
    <div className="bg-white border border-slate-200 rounded-lg overflow-hidden shadow-sm">
      <div className={`px-4 py-3 border-b border-slate-200 bg-gradient-to-r ${accent}`}>
        <div className="text-xs uppercase tracking-wider text-slate-600 font-semibold">{title} — {region}</div>
        <div className="text-base font-bold text-slate-900 mt-0.5">{subtitle}</div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm table-fixed">
          <colgroup>{colW.map((w, i) => <col key={i} style={{ width: w }} />)}</colgroup>
          <thead className="bg-slate-50 text-[11px] uppercase tracking-wider text-slate-600">
            <tr>
              <th className="px-3 py-2.5 text-left font-semibold">Vertical</th>
              <th className="px-3 py-2.5 text-right font-semibold">MRR Target</th>
              <th className="px-3 py-2.5 text-right font-semibold leading-tight">Gap Carried Fwd<br/><span className="text-slate-400 normal-case font-normal text-[10px]">(from {prevPeriodLabel})</span></th>
              <th className="px-3 py-2.5 text-right font-semibold leading-tight">MRR Achieved<br/><span className="text-slate-400 normal-case font-normal text-[10px]">(Start MRR)</span></th>
              <th className="px-3 py-2.5 text-right font-semibold leading-tight">MRR Gap<br/><span className="text-slate-400 normal-case font-normal text-[10px]">(Target + Carry − Achieved)</span></th>
              <th className="px-3 py-2.5 text-right font-semibold">Pipe Due for Closure</th>
              <th className="px-3 py-2.5 text-right font-semibold">Lost Pipe <span className="text-slate-400 normal-case font-normal text-[10px]">(Final Close in period)</span></th>
              <th className="px-3 py-2.5 text-right font-semibold">Pre-pipe (Deal Count)</th>
              <th className="px-3 py-2.5 text-right font-semibold">Pre-Pipe MRR</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {rows.map((r: any) => (
              <tr key={r.vert} className="hover:bg-slate-50">
                <td className="px-3 py-2 font-semibold text-slate-900 truncate">{r.vert}</td>
                <td className="px-3 py-2 text-right"><EditableTarget value={r.target} onChange={(v: number) => onTargetChange(r.vert, v)} region={region} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.carry} tone={r.carry > 0 ? 'carry' : (r.carry < 0 ? 'positive' : 'neutral')} region={region} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.achieved} deals={r.wonDeals} tone="positive" region={region} onClick={() => onDrill(`MRR Achieved — ${r.vert} — ${qLabel}`, r.wonDeals, region)} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.gap} tone={r.gap > 0 ? 'negative' : 'positive'} region={region} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.due} deals={r.dueDeals} tone="pipeline" region={region} onClick={() => onDrill(`Pipe Due for Closure — ${r.vert} — ${qLabel}`, r.dueDeals, region)} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.lost} deals={r.lostDeals} tone="negative" region={region} onClick={() => onDrill(`Lost Pipe — ${r.vert} — Final Close in ${qLabel}`, r.lostDeals, region)} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.prepipeCount} deals={r.prepipeDeals} tone="prepipe" region={region} isCount onClick={() => onDrill(`Pre-Pipe Deals — ${r.vert}`, r.prepipeDeals, region)} /></td>
                <td className="px-3 py-2 text-right"><MetricCell value={r.prepipeMrr} deals={r.prepipeDeals} tone="prepipe" region={region} onClick={() => onDrill(`Pre-Pipe MRR — ${r.vert}`, r.prepipeDeals, region)} /></td>
              </tr>
            ))}
            <tr className="bg-slate-100 font-semibold border-t-2 border-slate-300">
              <td className="px-3 py-2.5">Total</td>
              <td className="px-3 py-2.5 text-right tabular-nums">{fmt(totals.target, region)}</td>
              <td className={`px-3 py-2.5 text-right tabular-nums ${totals.carry > 0 ? 'text-orange-700' : totals.carry < 0 ? 'text-emerald-700' : 'text-slate-900'}`}>{fmt(totals.carry, region)}</td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.achieved} deals={totalAllDeals('won')} tone="positive" region={region} onClick={() => onDrill(`MRR Achieved — All ${region} — ${qLabel}`, totalAllDeals('won'), region)} /></td>
              <td className={`px-3 py-2.5 text-right tabular-nums ${totals.gap > 0 ? 'text-rose-700' : 'text-emerald-700'}`}>{fmt(totals.gap, region)}</td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.due} deals={totalAllDeals('due')} tone="pipeline" region={region} onClick={() => onDrill(`Pipe Due — All ${region} — ${qLabel}`, totalAllDeals('due'), region)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.lost} deals={totalAllDeals('lost')} tone="negative" region={region} onClick={() => onDrill(`Lost Pipe — All ${region} — Final Close in ${qLabel}`, totalAllDeals('lost'), region)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.prepipeCount} deals={totalAllDeals('prepipe')} tone="prepipe" region={region} isCount onClick={() => onDrill(`Pre-Pipe Deals — All ${region}`, totalAllDeals('prepipe'), region)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={totals.prepipeMrr} deals={totalAllDeals('prepipe')} tone="prepipe" region={region} onClick={() => onDrill(`Pre-Pipe MRR — All ${region}`, totalAllDeals('prepipe'), region)} /></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ==================== Defaults ====================
const DEFAULT_TARGETS: any = {
  India: {
    Q1: { 'BFSI India': 350000,  'EV MM': 1153104, 'EV ENT': 2646897, 'ME': 290000, 'Partner Sales': 483606 },
    Q2: { 'BFSI India': 2600000, 'EV MM': 1155111, 'EV ENT': 2651508, 'ME': 460000, 'Partner Sales': 483606 },
    Q3: { 'BFSI India': 2700000, 'EV MM': 999084,  'EV ENT': 2293351, 'ME': 565000, 'Partner Sales': 500673 },
    Q4: { 'BFSI India': 2400000, 'EV MM': 1092700, 'EV ENT': 2508245, 'ME': 685000, 'Partner Sales': 732116 },
  },
  US: {
    Q1: { 'US Healthcare': 39586, 'US Education': 43750, 'US Others': 0 },
    Q2: { 'US Healthcare': 39595, 'US Education': 43750, 'US Others': 0 },
    Q3: { 'US Healthcare': 39604, 'US Education': 43750, 'US Others': 0 },
    Q4: { 'US Healthcare': 39612, 'US Education': 43750, 'US Others': 0 },
  },
};
const EMPTY_FILTERS: any = { vertical: [], stage: [], status: [], typeOfSale: [], dealOwner: [], employeeSize: [], industry: [], productAsk: [], icpStatus: [], currentOwner: [], sdInitiator: [], saInitiator: [], pipeAddedRange: { from: '', to: '' }, expClosureRange: { from: '', to: '' }, finalClosureRange: { from: '', to: '' } };

// ==================== MRR Performance Tab ====================
function MrrPerformanceTab({ mrrAllRaw, region, fy, q }: any) {
  const [drillModal, setDrillModal] = useState<any>(null);
  const [drillSort, setDrillSort] = useState({ col: 'mrr', dir: 'desc' });
  const verticals = region === 'India' ? MRR_INDIA_VERTICALS : MRR_US_VERTICALS;
  const MRR_TYPES = ['NewBusiness', 'Expansion', 'Contraction', 'Churn', 'Reactivation'];

  const allParsed = useMemo(() => mrrAllRaw.map((r: any, i: number) => {
    const fy27vert = (r['FY27 Vertical'] || '').trim();
    let vert = FY27_VERTICAL_MAP[fy27vert];
    if (!vert && fy27vert === 'b.EV') vert = 'EV';
    const mrr = parseFloat(String(r[' Mrr'] || r['Mrr'] || '0').replace(/[^0-9.\-]/g, '')) || 0;
    const effDate  = r['Mrr Effective Date'] ? new Date(r['Mrr Effective Date'].trim()) : null;
    const declDate = r['Mrr Declared Date']  ? new Date(r['Mrr Declared Date'].trim())  : null;
    let achievementDate: Date | null = null;
    if (effDate && declDate) achievementDate = effDate >= declDate ? effDate : declDate;
    else achievementDate = effDate || declDate;
    return { _id: i, vert, mrr, type: (r['Mrr_type'] || '').trim(),
      effDate, declDate, achievementDate,
      company: (r['Companyname'] || '').trim(), owner: (r['Current Sales Owner'] || '').trim(),
      achDate: r['Achievement Month'] ? new Date(r['Achievement Month'].trim()) : null };
  }).filter((r: any) => r.vert && r.achievementDate), [mrrAllRaw]);

  const regionParsed = useMemo(() =>
    allParsed.filter((r: any) => (region === 'India' ? MRR_INDIA_VERTICALS : MRR_US_VERTICALS).includes(r.vert)),
    [allParsed, region]);

  const buildTable = (periodStart: Date, periodEnd: Date, qKey: string) => {
    const rows = regionParsed.filter((r: any) => r.achievementDate >= periodStart && r.achievementDate <= periodEnd);
    const targets = MRR_NET_TARGETS[region]?.[qKey] || {};
    const byVert: any = {};
    for (const v of verticals) {
      const vr = rows.filter((x: any) => x.vert === v);
      const byType: any = {};
      for (const t of MRR_TYPES) byType[t] = vr.filter((x: any) => x.type === t);
      const newB = byType.NewBusiness.reduce((a: number, x: any) => a + x.mrr, 0);
      const exp  = byType.Expansion.reduce((a: number, x: any) => a + x.mrr, 0);
      const con  = byType.Contraction.reduce((a: number, x: any) => a + x.mrr, 0);
      const chu  = byType.Churn.reduce((a: number, x: any) => a + x.mrr, 0);
      const rea  = byType.Reactivation.reduce((a: number, x: any) => a + x.mrr, 0);
      const net = newB + exp + con + chu + rea;
      const target = targets[v] || 0;
      const variance = target - net;
      byVert[v] = { byType, newB, exp, con, chu, rea, net, target, variance };
    }
    const totals = verticals.reduce((a: any, v: string) => {
      const r = byVert[v] || {};
      return {
        target: a.target + (r.target || 0),
        newB: a.newB+(r.newB||0), exp: a.exp+(r.exp||0), con: a.con+(r.con||0),
        chu: a.chu+(r.chu||0), rea: a.rea+(r.rea||0), net: a.net+(r.net||0),
        variance: a.variance + (r.variance || 0),
      };
    }, { target:0, newB:0, exp:0, con:0, chu:0, rea:0, net:0, variance:0 });
    const allByType: any = {};
    for (const t of MRR_TYPES) allByType[t] = rows.filter((x: any) => x.type === t);
    return { byVert, totals, allByType };
  };

  const [currStart, currEnd] = useMemo(() => quarterRange(fy, q), [q, fy]);
  const [nextQ, nextFy] = useMemo(() => {
    const order = ['Q1','Q2','Q3','Q4'];
    const idx = order.indexOf(q);
    return idx === 3 ? ['Q1', fy + 1] : [order[idx + 1], fy];
  }, [q, fy]) as [string, number];
  const [nextStart, nextEnd] = useMemo(() => quarterRange(nextFy, nextQ), [nextQ, nextFy]);

  const currTable = useMemo(() => buildTable(currStart, currEnd, q),     [regionParsed, currStart, currEnd, verticals, q, region]);
  const nextTable = useMemo(() => buildTable(nextStart, nextEnd, nextQ), [regionParsed, nextStart, nextEnd, verticals, nextQ, region]);

  if (!mrrAllRaw.length) return (
    <div className="bg-white border border-slate-200 rounded-xl p-12 text-center shadow-sm">
      <FileSpreadsheet className="w-10 h-10 mx-auto text-slate-400 mb-3" />
      <p className="text-base font-semibold text-slate-900">No MRR Transaction data loaded</p>
      <p className="text-xs text-slate-500 mt-2">The FY27 Transaction Detailed View is fetched automatically. Use the upload card as a manual fallback.</p>
    </div>
  );

  const drill = (title: string, deals: any[]) => setDrillModal({ title, deals: deals || [] });
  const colW = ['14%','11%','11%','11%','11%','10%','11%','11%','10%'];

  const renderTable = (title: string, subtitle: string, accent: string, t: any) => (
    <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm">
      <div className={`px-4 py-3 border-b border-slate-200 bg-gradient-to-r ${accent}`}>
        <div className="text-xs uppercase tracking-wider text-slate-600 font-semibold">{title} — {region}</div>
        <div className="text-base font-bold text-slate-900 mt-0.5">{subtitle}</div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm table-fixed">
          <colgroup>{colW.map((w,i)=><col key={i} style={{width:w}}/>)}</colgroup>
          <thead className="bg-slate-50 text-[11px] uppercase tracking-wider text-slate-600">
            <tr>
              <th className="px-3 py-2.5 text-left font-semibold">Vertical</th>
              <th className="px-3 py-2.5 text-right font-semibold">Target</th>
              <th className="px-3 py-2.5 text-right font-semibold">New MRR</th>
              <th className="px-3 py-2.5 text-right font-semibold">Expansion</th>
              <th className="px-3 py-2.5 text-right font-semibold">Contraction</th>
              <th className="px-3 py-2.5 text-right font-semibold">Churn</th>
              <th className="px-3 py-2.5 text-right font-semibold">Reactivation</th>
              <th className="px-3 py-2.5 text-right font-semibold">Net MRR</th>
              <th className="px-3 py-2.5 text-right font-semibold leading-tight">Net Variance<br/><span className="text-slate-400 normal-case font-normal text-[10px]">(Target − Net)</span></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {verticals.map((vert: string) => {
              const r = t.byVert[vert] || {};
              return (
                <tr key={vert} className="hover:bg-slate-50">
                  <td className="px-3 py-2 font-semibold text-slate-900 truncate">{vert}</td>
                  <td className="px-3 py-2 text-right"><span className={`tabular-nums font-medium ${(r.target||0) >= 0 ? 'text-slate-900' : 'text-rose-700'}`}>{fmt(r.target||0, region)}</span></td>
                  <td className="px-3 py-2 text-right"><MetricCell value={r.newB||0} deals={r.byType?.NewBusiness} tone="positive" region={region} onClick={()=>drill(`${title} — New MRR — ${vert}`, r.byType?.NewBusiness)} /></td>
                  <td className="px-3 py-2 text-right"><MetricCell value={r.exp||0}  deals={r.byType?.Expansion}   tone="positive" region={region} onClick={()=>drill(`${title} — Expansion — ${vert}`, r.byType?.Expansion)} /></td>
                  <td className="px-3 py-2 text-right"><MetricCell value={r.con||0}  deals={r.byType?.Contraction} tone="negative" region={region} onClick={()=>drill(`${title} — Contraction — ${vert}`, r.byType?.Contraction)} /></td>
                  <td className="px-3 py-2 text-right"><MetricCell value={r.chu||0}  deals={r.byType?.Churn}       tone="negative" region={region} onClick={()=>drill(`${title} — Churn — ${vert}`, r.byType?.Churn)} /></td>
                  <td className="px-3 py-2 text-right"><MetricCell value={r.rea||0}  deals={r.byType?.Reactivation}tone="pipeline" region={region} onClick={()=>drill(`${title} — Reactivation — ${vert}`, r.byType?.Reactivation)} /></td>
                  <td className="px-3 py-2 text-right font-semibold"><span className={`tabular-nums ${(r.net||0) >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>{fmt(r.net||0, region)}</span></td>
                  <td className="px-3 py-2 text-right font-semibold"><span className={`tabular-nums ${(r.variance||0) <= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>{fmt(r.variance||0, region)}</span></td>
                </tr>
              );
            })}
            <tr className="bg-slate-100 font-semibold border-t-2 border-slate-300">
              <td className="px-3 py-2.5">Total</td>
              <td className="px-3 py-2.5 text-right tabular-nums font-bold text-slate-900">{fmt(t.totals.target, region)}</td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={t.totals.newB} deals={t.allByType.NewBusiness}  tone="positive" region={region} onClick={()=>drill(`${title} — New MRR — All ${region}`, t.allByType.NewBusiness)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={t.totals.exp}  deals={t.allByType.Expansion}    tone="positive" region={region} onClick={()=>drill(`${title} — Expansion — All ${region}`, t.allByType.Expansion)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={t.totals.con}  deals={t.allByType.Contraction}  tone="negative" region={region} onClick={()=>drill(`${title} — Contraction — All ${region}`, t.allByType.Contraction)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={t.totals.chu}  deals={t.allByType.Churn}        tone="negative" region={region} onClick={()=>drill(`${title} — Churn — All ${region}`, t.allByType.Churn)} /></td>
              <td className="px-3 py-2.5 text-right"><MetricCell value={t.totals.rea}  deals={t.allByType.Reactivation} tone="pipeline" region={region} onClick={()=>drill(`${title} — Reactivation — All ${region}`, t.allByType.Reactivation)} /></td>
              <td className="px-3 py-2.5 text-right"><span className={`tabular-nums font-bold ${t.totals.net >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>{fmt(t.totals.net, region)}</span></td>
              <td className="px-3 py-2.5 text-right"><span className={`tabular-nums font-bold ${t.totals.variance <= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>{fmt(t.totals.variance, region)}</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );

  return (
    <>
      <div className="space-y-4">
        {renderTable('Current Quarter', `${q} FY${fy} • ${fmtDate(currStart)} – ${fmtDate(currEnd)}`, 'from-emerald-50 to-white', currTable)}
        {renderTable('Next Quarter',    `${nextQ} FY${nextFy} • ${fmtDate(nextStart)} – ${fmtDate(nextEnd)}`, 'from-violet-50 to-white', nextTable)}
      </div>

      {drillModal && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 p-4 flex items-center justify-center" onClick={()=>setDrillModal(null)}>
          <div className="bg-white border border-slate-200 rounded-lg shadow-2xl flex flex-col w-full max-w-[90vw]" style={{maxHeight:'calc(100vh - 2rem)'}} onClick={e=>e.stopPropagation()}>
            <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between flex-shrink-0">
              <div>
                <h3 className="font-semibold text-slate-900 text-sm">{drillModal.title}</h3>
                <p className="text-xs text-slate-600 mt-0.5">{drillModal.deals.length} entries • {fmt(drillModal.deals.reduce((a: number, r: any)=>a+r.mrr,0), region)}</p>
              </div>
              <button onClick={()=>setDrillModal(null)} className="text-slate-500 hover:text-slate-900 p-1 rounded hover:bg-slate-100"><X className="w-4 h-4"/></button>
            </div>
            <div className="flex-1 overflow-auto min-h-0">
              <table className="w-full text-xs table-fixed">
                <colgroup><col style={{width:'22%'}}/><col style={{width:'12%'}}/><col style={{width:'11%'}}/><col style={{width:'11%'}}/><col style={{width:'12%'}}/><col style={{width:'11%'}}/><col style={{width:'11%'}}/><col style={{width:'10%'}}/></colgroup>
                <thead className="bg-slate-50 text-[10px] uppercase tracking-wider sticky top-0 z-10">
                  <tr>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="company" label="Company" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="vert" label="Vertical" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="type" label="MRR Type" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-right font-semibold"><DrillSortHeader col="mrr" label="MRR" sort={drillSort} onSort={setDrillSort} align="right" /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="owner" label="Owner" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="effDate" label="Effective Date" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="declDate" label="Declared Date" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="achievementDate" label="Achievement Date" sort={drillSort} onSort={setDrillSort} /></th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortRows(drillModal.deals, drillSort).map((d: any,i: number)=>{
                    const declWon = d.effDate && d.declDate && d.declDate > d.effDate;
                    const effWon  = d.effDate && d.declDate && d.effDate > d.declDate;
                    return (
                    <tr key={i} className="hover:bg-slate-50">
                      <td className="px-2 py-1.5 overflow-hidden"><div className="truncate font-medium text-slate-900" title={d.company}>{d.company||'—'}</div></td>
                      <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] font-semibold text-slate-800">{d.vert||'—'}</div></td>
                      <td className="px-2 py-1.5"><span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded ${d.type==='Churn'||d.type==='Contraction'?'bg-rose-100 text-rose-800':d.type==='NewBusiness'?'bg-emerald-100 text-emerald-800':d.type==='Expansion'?'bg-sky-100 text-sky-800':'bg-violet-100 text-violet-800'}`}>{d.type}</span></td>
                      <td className={`px-2 py-1.5 text-right tabular-nums font-medium ${d.mrr < 0 ? 'text-rose-700' : 'text-slate-900'}`}>{fmt(d.mrr, region)}</td>
                      <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] text-slate-600">{d.owner||'—'}</div></td>
                      <td className={`px-2 py-1.5 text-[10px] ${effWon ? 'text-sky-700 font-semibold' : 'text-slate-600'}`}>{d.effDate ? d.effDate.toISOString().slice(0,10) : '—'}</td>
                      <td className={`px-2 py-1.5 text-[10px] ${declWon ? 'text-sky-700 font-semibold' : 'text-slate-600'}`}>{d.declDate ? d.declDate.toISOString().slice(0,10) : '—'}</td>
                      <td className="px-2 py-1.5 text-[10px] text-slate-900 font-semibold">{d.achievementDate ? d.achievementDate.toISOString().slice(0,10) : '—'}</td>
                    </tr>
                  );})}
                  {drillModal.deals.length===0 && <tr><td colSpan={8} className="px-4 py-6 text-center text-slate-400 text-xs">No entries</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ==================== Renewal Risks Tab ====================
function RenewalRisksTab({ renewalRaw, region, fy, q }: any) {
  const [drillModal, setDrillModal] = useState<any>(null);
  const [drillSort, setDrillSort] = useState({ col: 'mrr', dir: 'desc' });
  const todayD = useMemo(() => new Date(), []);
  const [, qe] = useMemo(() => quarterRange(fy, q), [fy, q]);

  const verticals = region === 'India' ? RENEWAL_INDIA_VERTICALS : RENEWAL_US_VERTICALS;
  const INR_TO_USD = 92;

  const renewalRows = useMemo(() => renewalRaw.map((r: any, i: number) => {
    const bu = (r['New Final Bu'] || '').trim();
    const vert = RENEWAL_BU_MAP[bu];
    const endDate = r['Contract End Date Date'] ? new Date(r['Contract End Date Date'].trim()) : null;
    const mrrInr = parseFloat(String(r['MRR'] || '0').replace(/[^0-9.\-]/g, '')) || 0;
    const isUsAccount = RENEWAL_US_VERTICALS.includes(vert);
    const mrr = isUsAccount ? (mrrInr / INR_TO_USD) : mrrInr;
    return { _id: i, bu, vert, endDate, mrr,
      company: (r['Companyname'] || r['Tenant Name'] || '').trim(),
      csm: (r['CSM Owner'] || '').trim(),
      cp: (r['Acc Client Partner Owner Name'] || '').trim(),
      startDate: r['Contract Start Date Date'] ? new Date(r['Contract Start Date Date'].trim()) : null,
    };
  }).filter((r: any) => r.vert && r.endDate), [renewalRaw]);

  const regionRows = useMemo(() =>
    renewalRows.filter((r: any) => verticals.includes(r.vert)),
    [renewalRows, verticals]);

  const byVert = useMemo(() => {
    const out: any = {};
    for (const v of verticals) {
      const vr = regionRows.filter((r: any) => r.vert === v);
      const overdue = vr.filter((r: any) => r.endDate < todayD);
      const upcoming = vr.filter((r: any) => r.endDate >= todayD && r.endDate <= qe);
      const total = [...overdue, ...upcoming];
      out[v] = {
        overdue, overdueCount: overdue.length, overdueMrr: overdue.reduce((a: number, r: any)=>a+r.mrr,0),
        upcoming, upcomingCount: upcoming.length, upcomingMrr: upcoming.reduce((a: number, r: any)=>a+r.mrr,0),
        total, totalCount: total.length, totalMrr: total.reduce((a: number, r: any)=>a+r.mrr,0),
      };
    }
    return out;
  }, [regionRows, verticals, todayD, qe]);

  const totals = useMemo(() => verticals.reduce((a: any, v: string) => {
    const r = byVert[v]||{};
    return {
      overdueCount:  a.overdueCount  + (r.overdueCount  || 0), overdueMrr:  a.overdueMrr  + (r.overdueMrr  || 0),
      upcomingCount: a.upcomingCount + (r.upcomingCount || 0), upcomingMrr: a.upcomingMrr + (r.upcomingMrr || 0),
      totalCount:    a.totalCount    + (r.totalCount    || 0), totalMrr:    a.totalMrr    + (r.totalMrr    || 0),
    };
  }, {overdueCount:0,overdueMrr:0,upcomingCount:0,upcomingMrr:0,totalCount:0,totalMrr:0}), [byVert, verticals]);

  const drill = (title: string, rows: any[]) => setDrillModal({ title, rows: rows||[] });

  if (!renewalRaw.length) return (
    <div className="bg-white border border-slate-200 rounded-xl p-12 text-center shadow-sm">
      <p className="text-slate-500 text-sm">No <strong>Renewals</strong> data loaded. It is fetched automatically; use the upload card as a manual fallback.</p>
    </div>
  );

  const colW = ['18%','12%','14%','12%','14%','15%','15%'];

  return (
    <>
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm">
        <div className="px-4 py-3 border-b border-slate-200 bg-gradient-to-r from-rose-50 to-white">
          <div className="text-xs uppercase tracking-wider text-slate-600 font-semibold">Renewal Risks — {region}</div>
          <div className="text-base font-bold text-slate-900 mt-0.5">{q} FY{fy} • Contract End Date analysis as of {todayD.toISOString().slice(0,10)}</div>
          {region === 'US' && <div className="text-[11px] text-slate-500 mt-1">USD values converted from INR at 1 USD = ₹92</div>}
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm table-fixed">
            <colgroup>{colW.map((w,i)=><col key={i} style={{width:w}}/>)}</colgroup>
            <thead className="bg-slate-50 text-[11px] uppercase tracking-wider text-slate-600">
              <tr>
                <th className="px-3 py-2.5 text-left font-semibold">Vertical</th>
                <th className="px-3 py-2.5 text-right font-semibold leading-tight">Overdue<br/><span className="text-slate-400 normal-case font-normal text-[10px]">end date past today</span></th>
                <th className="px-3 py-2.5 text-right font-semibold leading-tight">Overdue MRR<br/><span className="text-slate-400 normal-case font-normal text-[10px]">at risk now</span></th>
                <th className="px-3 py-2.5 text-right font-semibold leading-tight">Upcoming in {q}<br/><span className="text-slate-400 normal-case font-normal text-[10px]">today → Q end</span></th>
                <th className="px-3 py-2.5 text-right font-semibold leading-tight">Upcoming MRR</th>
                <th className="px-3 py-2.5 text-right font-semibold leading-tight">Total in {q}<br/><span className="text-slate-400 normal-case font-normal text-[10px]">unique accounts</span></th>
                <th className="px-3 py-2.5 text-right font-semibold leading-tight">Total MRR</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {verticals.map((vert: string) => {
                const r = byVert[vert]||{};
                return (
                  <tr key={vert} className="hover:bg-slate-50">
                    <td className="px-3 py-2 font-semibold text-slate-900 truncate">{vert}</td>
                    <td className="px-3 py-2 text-right">{(r.overdueCount||0) > 0 ? <button onClick={()=>drill(`Overdue Renewals — ${vert}`, r.overdue)} className="tabular-nums text-rose-700 font-bold hover:underline decoration-dotted underline-offset-2">{r.overdueCount}</button> : <span className="text-slate-400 tabular-nums">0</span>}</td>
                    <td className="px-3 py-2 text-right">{(r.overdueMrr||0) > 0 ? <button onClick={()=>drill(`Overdue Renewals — ${vert}`, r.overdue)} className="tabular-nums text-rose-700 font-semibold hover:underline decoration-dotted underline-offset-2">{fmt(r.overdueMrr, region)}</button> : <span className="text-slate-400 tabular-nums">{fmt(0, region)}</span>}</td>
                    <td className="px-3 py-2 text-right">{(r.upcomingCount||0) > 0 ? <button onClick={()=>drill(`Upcoming in ${q} — ${vert}`, r.upcoming)} className="tabular-nums text-amber-700 font-semibold hover:underline decoration-dotted underline-offset-2">{r.upcomingCount}</button> : <span className="text-slate-400 tabular-nums">0</span>}</td>
                    <td className="px-3 py-2 text-right">{(r.upcomingMrr||0) > 0 ? <button onClick={()=>drill(`Upcoming in ${q} — ${vert}`, r.upcoming)} className="tabular-nums text-amber-700 font-semibold hover:underline decoration-dotted underline-offset-2">{fmt(r.upcomingMrr, region)}</button> : <span className="text-slate-400 tabular-nums">{fmt(0, region)}</span>}</td>
                    <td className="px-3 py-2 text-right">{(r.totalCount||0) > 0 ? <button onClick={()=>drill(`Total Renewal Risk in ${q} — ${vert}`, r.total)} className="tabular-nums text-slate-900 font-bold hover:underline decoration-dotted underline-offset-2">{r.totalCount}</button> : <span className="text-slate-400 tabular-nums">0</span>}</td>
                    <td className="px-3 py-2 text-right tabular-nums font-bold text-slate-900">{fmt(r.totalMrr||0, region)}</td>
                  </tr>
                );
              })}
              <tr className="bg-slate-100 font-semibold border-t-2 border-slate-300">
                <td className="px-3 py-2.5">Total</td>
                <td className="px-3 py-2.5 text-right tabular-nums text-rose-700 font-bold">{totals.overdueCount}</td>
                <td className="px-3 py-2.5 text-right tabular-nums text-rose-700 font-bold">{fmt(totals.overdueMrr, region)}</td>
                <td className="px-3 py-2.5 text-right tabular-nums text-amber-700 font-semibold">{totals.upcomingCount}</td>
                <td className="px-3 py-2.5 text-right tabular-nums text-amber-700 font-semibold">{fmt(totals.upcomingMrr, region)}</td>
                <td className="px-3 py-2.5 text-right tabular-nums text-slate-900 font-bold">{totals.totalCount}</td>
                <td className="px-3 py-2.5 text-right tabular-nums text-slate-900 font-bold">{fmt(totals.totalMrr, region)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {drillModal && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 p-4 flex items-center justify-center" onClick={()=>setDrillModal(null)}>
          <div className="bg-white border border-slate-200 rounded-lg shadow-2xl flex flex-col w-full max-w-[90vw]" style={{maxHeight:'calc(100vh - 2rem)'}} onClick={e=>e.stopPropagation()}>
            <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between flex-shrink-0">
              <div>
                <h3 className="font-semibold text-slate-900 text-sm">{drillModal.title}</h3>
                <p className="text-xs text-slate-600 mt-0.5">{drillModal.rows.length} accounts • {fmt(drillModal.rows.reduce((a: number, r: any)=>a+r.mrr,0), region)}</p>
              </div>
              <button onClick={()=>setDrillModal(null)} className="text-slate-500 hover:text-slate-900 p-1 rounded hover:bg-slate-100"><X className="w-4 h-4"/></button>
            </div>
            <div className="flex-1 overflow-auto min-h-0">
              <table className="w-full text-xs table-fixed">
                <colgroup><col style={{width:'22%'}}/><col style={{width:'12%'}}/><col style={{width:'12%'}}/><col style={{width:'14%'}}/><col style={{width:'12%'}}/><col style={{width:'14%'}}/><col style={{width:'14%'}}/></colgroup>
                <thead className="bg-slate-50 text-[10px] uppercase tracking-wider sticky top-0 z-10">
                  <tr>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="company" label="Account" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="vert" label="Vertical" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-right font-semibold"><DrillSortHeader col="mrr" label="MRR" sort={drillSort} onSort={setDrillSort} align="right" /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="startDate" label="Contract Start" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="endDate" label="Contract End" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="csm" label="CSM" sort={drillSort} onSort={setDrillSort} /></th>
                    <th className="px-2 py-1.5 text-left font-semibold"><DrillSortHeader col="cp" label="Client Partner" sort={drillSort} onSort={setDrillSort} /></th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortRows(drillModal.rows, drillSort).map((r: any,i: number) => (
                    <tr key={i} className={`hover:bg-slate-50 ${r.endDate < todayD ? 'bg-rose-50/50' : ''}`}>
                      <td className="px-2 py-1.5 overflow-hidden"><div className="truncate font-medium text-slate-900 text-xs" title={r.company}>{r.company||'—'}</div></td>
                      <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] font-semibold text-slate-800">{r.vert||'—'}</div></td>
                      <td className="px-2 py-1.5 text-right tabular-nums font-medium text-slate-900">{fmt(r.mrr, region)}</td>
                      <td className="px-2 py-1.5 text-[10px] text-slate-600 whitespace-nowrap">{r.startDate ? r.startDate.toISOString().slice(0,10) : '—'}</td>
                      <td className="px-2 py-1.5 whitespace-nowrap">
                        <span className={`text-[10px] font-semibold ${r.endDate < todayD ? 'text-rose-700' : 'text-amber-700'}`}>{r.endDate ? r.endDate.toISOString().slice(0,10) : '—'}</span>
                        {r.endDate < todayD && <span className="ml-1 bg-rose-100 text-rose-700 text-[9px] font-bold px-1 rounded">OVERDUE</span>}
                      </td>
                      <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] text-slate-600">{r.csm||'—'}</div></td>
                      <td className="px-2 py-1.5 overflow-hidden"><div className="truncate text-[10px] text-slate-600">{r.cp||'—'}</div></td>
                    </tr>
                  ))}
                  {drillModal.rows.length===0 && <tr><td colSpan={7} className="px-4 py-6 text-center text-slate-400 text-xs">No accounts</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ==================== Main App ====================
export default function App() {
  const [raw, setRaw] = useState<any[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [colMap, setColMap] = useState<any>(null);
  const [showMap, setShowMap] = useState(false);
  const [targets, setTargets] = useState<any>(DEFAULT_TARGETS);
  const [targetsSource, setTargetsSource] = useState('');
  const [extractedRows, setExtractedRows] = useState<any[]>([]);
  const [activeRegion, setActiveRegion] = useState('India');
  const [drill, setDrill] = useState<any>(null);
  const [dealDetail, setDealDetail] = useState<any>(null);
  const [globalFilters, setGlobalFilters] = useState<any>(EMPTY_FILTERS);
  const [tableSearch, setTableSearch] = useState('');
  const [tableSort, setTableSort] = useState({ column: 'mrr', direction: 'desc' });
  const [tablePage, setTablePage] = useState(0);
  const [tablePageSize, setTablePageSize] = useState(20);
  const [showDiagnostics, setShowDiagnostics] = useState(false);
  const [filtersCollapsed, setFiltersCollapsed] = useState(false);
  const [mrrRaw, setMrrRaw] = useState<any[]>([]);
  const [mrrAllRaw, setMrrAllRaw] = useState<any[]>([]);
  const [mrrFileName, setMrrFileName] = useState('');
  const [renewalRaw, setRenewalRaw] = useState<any[]>([]);
  const [renewalFileName, setRenewalFileName] = useState('');
  const [activeTab, setActiveTab] = useState('pipeline');
  const [oppUploadedAt, setOppUploadedAt] = useState('');
  const [mrrUploadedAt, setMrrUploadedAt] = useState('');
  const [renewalUploadedAt, setRenewalUploadedAt] = useState('');
  const [autoSync, setAutoSync] = useState<{ status: string; at: string }>({ status: 'idle', at: '' });
  const fileRef = useRef<HTMLInputElement>(null);
  const xlsxRef = useRef<HTMLInputElement>(null);
  const mrrFileRef = useRef<HTMLInputElement>(null);
  const renewalFileRef = useRef<HTMLInputElement>(null);
  const saveTimer = useRef<any>(null);

  // ---- Ingest helpers: take an array of row objects and load into state ----
  const ingestOpportunityRows = useCallback((rows: any[], name: string, uploadedAt: string, persist: boolean) => {
    if (!rows || !rows.length) return;
    const hdrs = Object.keys(rows[0]);
    const mapping = buildOppMapping(hdrs);
    setHeaders(hdrs); setRaw(rows); setFileName(name); setColMap(mapping); setOppUploadedAt(uploadedAt); setError('');
    if (persist) storage.set('pipeline_opportunity_v1', JSON.stringify({ rows, headers: hdrs, fileName: name, uploadedAt }));
  }, []);

  const ingestMrrRows = useCallback((rows: any[], name: string, uploadedAt: string, persist: boolean) => {
    if (!rows || !rows.length) return;
    const nb = rows.filter(x => (x['Mrr_type'] || '').trim() === 'NewBusiness');
    setMrrRaw(nb); setMrrAllRaw(rows); setMrrFileName(name); setMrrUploadedAt(uploadedAt); setError('');
    if (persist) storage.set('pipeline_mrr_v1', JSON.stringify({ rows, fileName: name, uploadedAt }));
  }, []);

  const ingestRenewalRows = useCallback((rows: any[], name: string, uploadedAt: string, persist: boolean) => {
    if (!rows || !rows.length) return;
    setRenewalRaw(rows); setRenewalFileName(name); setRenewalUploadedAt(uploadedAt); setError('');
    if (persist) storage.set('pipeline_renewal_v1', JSON.stringify({ rows, fileName: name, uploadedAt }));
  }, []);

  // ---- Auto-fetch the daily-updated JSON files (populated by GitHub Actions) ----
  const fetchAutoData = useCallback(async () => {
    setAutoSync({ status: 'loading', at: '' });
    let any = false;
    const tryFetch = async (file: string) => {
      try {
        const res = await fetch(dataUrl(file), { cache: 'no-store' });
        if (!res.ok) return null;
        return await res.json();
      } catch { return null; }
    };
    const [opp, mrr, ren] = await Promise.all([
      tryFetch('opportunities.json'),
      tryFetch('mrr_transactions.json'),
      tryFetch('renewals.json'),
    ]);
    if (opp?.rows?.length) { ingestOpportunityRows(opp.rows, opp.fileName || 'opportunities.json (auto)', opp.generatedAt || new Date().toISOString(), false); any = true; }
    if (mrr?.rows?.length) { ingestMrrRows(mrr.rows, mrr.fileName || 'mrr_transactions.json (auto)', mrr.generatedAt || new Date().toISOString(), false); any = true; }
    if (ren?.rows?.length) { ingestRenewalRows(ren.rows, ren.fileName || 'renewals.json (auto)', ren.generatedAt || new Date().toISOString(), false); any = true; }
    setAutoSync({ status: any ? 'ok' : 'empty', at: new Date().toISOString() });
  }, [ingestOpportunityRows, ingestMrrRows, ingestRenewalRows]);

  // Load targets (storage), then restore any locally-persisted uploads, then auto-fetch.
  useEffect(() => {
    (async () => {
      try {
        const r = await storage.get('pipeline_targets_v4');
        if (r?.value) {
          const parsed = JSON.parse(r.value);
          const merge = (region: string, q: string) => ({ ...DEFAULT_TARGETS[region][q], ...(parsed[region]?.[q] || {}) });
          setTargets({
            India: { Q1: merge('India','Q1'), Q2: merge('India','Q2'), Q3: merge('India','Q3'), Q4: merge('India','Q4') },
            US:    { Q1: merge('US','Q1'),    Q2: merge('US','Q2'),    Q3: merge('US','Q3'),    Q4: merge('US','Q4') },
          });
          setTargetsSource('saved (storage)');
        }
      } catch (e) {}

      // Restore locally-persisted manual uploads (fallback when no auto data yet)
      try { const r = await storage.get('pipeline_opportunity_v1'); if (r?.value) { const d = JSON.parse(r.value); ingestOpportunityRows(d.rows, d.fileName, d.uploadedAt || '', false); } } catch (e) {}
      try { const r = await storage.get('pipeline_mrr_v1');         if (r?.value) { const d = JSON.parse(r.value); ingestMrrRows(d.rows, d.fileName, d.uploadedAt || '', false); } } catch (e) {}
      try { const r = await storage.get('pipeline_renewal_v1');     if (r?.value) { const d = JSON.parse(r.value); ingestRenewalRows(d.rows, d.fileName, d.uploadedAt || '', false); } } catch (e) {}

      // Then overlay fresh auto-fetched data (takes precedence if present)
      fetchAutoData();
    })();
  }, []);

  const persistTargets = useCallback((t: any) => {
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => { storage.set('pipeline_targets_v4', JSON.stringify(t)); }, 500);
  }, []);

  const fy = getCurrentFY();
  const cq = getCurrentQuarter();
  const [nq, nfy] = advanceQ(cq, fy);
  const [hq1, hfy1] = advanceQ(nq, nfy);
  const [hq2, hfy2] = advanceQ(hq1, hfy1);
  const halfStart = quarterRange(hfy1, hq1)[0];
  const halfEnd = quarterRange(hfy2, hq2)[1];
  const halfLabel = hfy1 === hfy2 ? `${hq1}–${hq2} ${fyLabel(hfy1)}` : `${hq1} ${fyLabel(hfy1)} – ${hq2} ${fyLabel(hfy2)}`;
  const nextQRange = quarterRange(nfy, nq);
  const thisQRange = quarterRange(fy, cq);

  const updateTarget = (region: string, period: string, vert: string, value: number) => {
    let next;
    if (period === 'nextH') {
      const half = Math.round(value / 2);
      next = { ...targets, [region]: { ...targets[region], [hq1]: { ...targets[region][hq1], [vert]: half }, [hq2]: { ...targets[region][hq2], [vert]: value - half } }};
    } else {
      const q = period === 'thisQ' ? cq : nq;
      next = { ...targets, [region]: { ...targets[region], [q]: { ...targets[region][q], [vert]: value } }};
    }
    setTargets(next); persistTargets(next); setTargetsSource('saved');
  };

  // ---- Manual upload (CSV) parsers — fallback path ----
  const handleFile = useCallback((file?: File) => {
    if (!file) return;
    setError('');
    const reader = new FileReader();
    reader.onload = (e) => Papa.parse(e.target!.result as string, {
      header: true, skipEmptyLines: true, dynamicTyping: false, transformHeader: (h) => (h || '').trim(),
      complete: (res: any) => {
        if (!res.data?.length) { setError('No rows found in CSV.'); return; }
        ingestOpportunityRows(res.data, file.name, new Date().toISOString(), true);
      },
      error: (err: any) => setError('CSV parse error: ' + err.message),
    });
    reader.onerror = () => setError('Failed to read file.');
    reader.readAsText(file);
  }, [ingestOpportunityRows]);

  const handleXlsx = useCallback((file?: File) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buf = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(buf, { type: 'array' });
        const { targets: parsed, extractedCount, extractedRows: rows } = parseTargetXlsx(wb);
        if (extractedCount === 0) { setError('No MRR targets found.'); return; }
        setTargets(parsed); setExtractedRows(rows); setTargetsSource(`xlsx (${extractedCount} targets, ${file.name})`);
        storage.set('pipeline_targets_v4', JSON.stringify(parsed)); setError('');
      } catch (e: any) { setError('Failed to parse XLSX: ' + e.message); }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleMrrFile = useCallback((file?: File) => {
    if (!file) return;
    setError('');
    const reader = new FileReader();
    reader.onload = (e) => Papa.parse(e.target!.result as string, {
      header: true, skipEmptyLines: true, dynamicTyping: false, transformHeader: (h) => (h || '').trim(),
      complete: (res: any) => {
        if (!res.data?.length) { setError('No rows found in MRR file.'); return; }
        ingestMrrRows(res.data, file.name, new Date().toISOString(), true);
      },
      error: (err: any) => setError('MRR CSV parse error: ' + err.message),
    });
    reader.onerror = () => setError('Failed to read MRR file.');
    reader.readAsText(file);
  }, [ingestMrrRows]);

  const handleRenewalFile = useCallback((file?: File) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => Papa.parse(e.target!.result as string, {
      header: true, skipEmptyLines: true, dynamicTyping: false, transformHeader: (h) => (h || '').trim(),
      complete: (res: any) => {
        if (!res.data?.length) { setError('No rows found in Renewal file.'); return; }
        ingestRenewalRows(res.data, file.name, new Date().toISOString(), true);
      },
      error: (err: any) => setError('Renewal CSV parse error: ' + err.message),
    });
    reader.onerror = () => setError('Failed to read Renewal file.');
    reader.readAsText(file);
  }, [ingestRenewalRows]);

  const mrrDeals = useMemo(() => {
    if (!mrrRaw.length) return [];
    return mrrRaw.map((row, i) => {
      const fy27vert = (row['FY27 Vertical'] || '').trim();
      const salesBu  = (row['Sales Bu'] || '').trim().toLowerCase();
      let resolvedLabel = FY27_VERTICAL_MAP[fy27vert];
      if (!resolvedLabel && fy27vert === 'b.EV') resolvedLabel = (salesBu.includes('enterprise') ? 'EV ENT' : 'EV MM');
      const mrrVal = parseFloat(String(row[' Mrr'] || row['Mrr'] || '0').replace(/[^0-9.\-]/g, '')) || 0;
      const effDate = row['Mrr Effective Date'] ? new Date(row['Mrr Effective Date'].trim()) : null;
      const achMonth = row['Achievement Month'] ? new Date(row['Achievement Month'].trim()) : null;
      const mrrStartDate = achMonth || effDate;
      return {
        _id: `mrr_${i}`, _raw: row, _isMrrEntry: true,
        name: row['Description'] || row['Companyname'] || `MRR Entry ${i + 1}`,
        account: (row['Companyname'] || '').trim(),
        owner: (row['Current Sales Owner'] || '').trim(),
        currentOwner: (row['Current Sales Owner'] || '').trim(),
        sdInitiator: '', saInitiator: '',
        resolvedLabel: resolvedLabel || null, resolutionSource: 'mrr-file',
        stage: 'Won', status: 'Won', mrr: mrrVal, mrrStartDate,
        closureDate: effDate, expClosureDate: null, pipeAddedDate: null,
        typeOfSale: 'New Business', employeeSize: '',
        industry: (row['Industry'] || '').trim(), productAsk: '', icpStatus: '',
      };
    }).filter(d => d.resolvedLabel && d.mrr > 0);
  }, [mrrRaw]);

  const deals = useMemo(() => {
    if (!raw.length || !colMap) return [];
    return raw.map((row, i) => {
      const adDate = parseDate(colMap.pipeAdded ? row[colMap.pipeAdded] : null);
      const gfDate = parseDate(colMap.usPipeAdded ? row[colMap.usPipeAdded] : null);
      const rawVertical = colMap.vertical ? (row[colMap.vertical] || '').trim() : '';
      const rawTeams    = colMap.teams ? (row[colMap.teams] || '').trim() : '';
      const { label, source } = resolveVertical(rawVertical, rawTeams);
      return {
        _id: i, _raw: row,
        name: colMap.dealName ? row[colMap.dealName] : `Deal ${i + 1}`,
        account: colMap.account ? row[colMap.account] : '',
        owner: colMap.owner ? (row[colMap.owner] || '').trim() : '',
        rawVertical, rawTeams, resolvedLabel: label, resolutionSource: source,
        stage: normalizeStage(colMap.stage ? row[colMap.stage] : ''),
        status: normalizeStatus(colMap.status ? row[colMap.status] : ''),
        mrr: colMap.mrr ? (parseFloat(String(row[colMap.mrr] || '0').replace(/[^0-9.\-]/g, '')) || 0) : 0,
        closureDate: parseDate(colMap.closureDate ? row[colMap.closureDate] : null),
        mrrStartDate: parseDate(colMap.mrrStartDate ? row[colMap.mrrStartDate] : null),
        expClosureDate: parseDate(colMap.expClosureDate ? row[colMap.expClosureDate] : null),
        pipeAddedDate: gfDate || adDate,
        typeOfSale:   colMap.typeOfSale ? (row[colMap.typeOfSale] || '').trim() : '',
        employeeSize: colMap.employeeSize ? (row[colMap.employeeSize] || '').trim() : '',
        industry:     colMap.industry ? (row[colMap.industry] || '').trim() : '',
        productAsk:   colMap.productAsk ? (row[colMap.productAsk] || '').trim() : '',
        icpStatus:    colMap.icpStatus ? (row[colMap.icpStatus] || '').trim() : '',
        currentOwner: colMap.currentOwner ? (row[colMap.currentOwner] || '').trim() : '',
        sdInitiator:  colMap.sdInitiator ? (row[colMap.sdInitiator] || '').trim() : '',
        saInitiator:  colMap.saInitiator ? (row[colMap.saInitiator] || '').trim() : '',
      };
    });
  }, [raw, colMap]);

  const filteredDeals = useMemo(() => deals.filter(d => {
    if (globalFilters.vertical.length && !globalFilters.vertical.includes(d.resolvedLabel || '(unmapped)')) return false;
    if (globalFilters.stage.length && !globalFilters.stage.includes(d.stage)) return false;
    if (globalFilters.status.length && !globalFilters.status.includes(d.status)) return false;
    if (globalFilters.typeOfSale.length && !globalFilters.typeOfSale.includes(d.typeOfSale || '(blank)')) return false;
    if (globalFilters.dealOwner.length && !globalFilters.dealOwner.includes(d.owner || '(blank)')) return false;
    if (globalFilters.employeeSize.length && !globalFilters.employeeSize.includes(d.employeeSize || '(blank)')) return false;
    if (globalFilters.industry.length && !globalFilters.industry.includes(d.industry || '(blank)')) return false;
    if (globalFilters.productAsk.length && !globalFilters.productAsk.includes(d.productAsk || '(blank)')) return false;
    if (globalFilters.icpStatus.length && !globalFilters.icpStatus.includes(d.icpStatus || '(blank)')) return false;
    if (globalFilters.currentOwner.length && !globalFilters.currentOwner.includes(d.currentOwner || '(blank)')) return false;
    if (globalFilters.sdInitiator.length && !globalFilters.sdInitiator.includes(d.sdInitiator || '(blank)')) return false;
    if (globalFilters.saInitiator.length && !globalFilters.saInitiator.includes(d.saInitiator || '(blank)')) return false;
    const toISO = (dt: any) => dt instanceof Date && !isNaN(dt as any) ? dt.toISOString().slice(0, 10) : '';
    if (globalFilters.pipeAddedRange?.from || globalFilters.pipeAddedRange?.to) {
      const v = toISO(d.pipeAddedDate); if (!v) return false;
      if (globalFilters.pipeAddedRange.from && v < globalFilters.pipeAddedRange.from) return false;
      if (globalFilters.pipeAddedRange.to   && v > globalFilters.pipeAddedRange.to)   return false;
    }
    if (globalFilters.expClosureRange?.from || globalFilters.expClosureRange?.to) {
      const v = toISO(d.expClosureDate); if (!v) return false;
      if (globalFilters.expClosureRange.from && v < globalFilters.expClosureRange.from) return false;
      if (globalFilters.expClosureRange.to   && v > globalFilters.expClosureRange.to)   return false;
    }
    if (globalFilters.finalClosureRange?.from || globalFilters.finalClosureRange?.to) {
      const v = toISO(d.closureDate); if (!v) return false;
      if (globalFilters.finalClosureRange.from && v < globalFilters.finalClosureRange.from) return false;
      if (globalFilters.finalClosureRange.to   && v > globalFilters.finalClosureRange.to)   return false;
    }
    return true;
  }), [deals, globalFilters]);

  const regionDeals = useMemo(() => {
    const labels = activeRegion === 'India' ? INDIA_VERTICALS : US_VERTICALS;
    return filteredDeals.filter(d => labels.includes(d.resolvedLabel as string));
  }, [filteredDeals, activeRegion]);

  const mrrRegionDeals = useMemo(() => {
    const labels = activeRegion === 'India' ? INDIA_VERTICALS : US_VERTICALS;
    return mrrDeals.filter(d => labels.includes(d.resolvedLabel as string));
  }, [mrrDeals, activeRegion]);

  const mrrOnlyAchieved = mrrRegionDeals;

  const verticalBreakdown = useMemo(() => {
    const counts = new Map();
    deals.forEach(d => {
      const key = `${d.rawVertical || '(blank)'}||${d.rawTeams || '(blank)'}`;
      const entry = counts.get(key) || { rawVertical: d.rawVertical || '(blank)', rawTeams: d.rawTeams || '(blank)', resolvedLabel: d.resolvedLabel, source: d.resolutionSource, count: 0 };
      entry.count++; counts.set(key, entry);
    });
    return Array.from(counts.values()).sort((a, b) => b.count - a.count);
  }, [deals]);

  const openDrill = (title: string, deals: any[], dealRegion: string) => setDrill({ title, deals, region: dealRegion });
  const verticalList = activeRegion === 'India' ? INDIA_VERTICALS : US_VERTICALS;

  const buildPeriodTargets = (region: string, period: string) => {
    const out: any = {};
    for (const vert of verticalList) out[vert] = getPeriodTarget(targets, region, period, vert, fy, cq);
    return out;
  };
  const thisQTargets = buildPeriodTargets(activeRegion, 'thisQ');
  const nextQTargets = buildPeriodTargets(activeRegion, 'nextQ');
  const nextHTargets = buildPeriodTargets(activeRegion, 'nextH');

  const thisQAchievedByVert = useMemo(() => computeAchievedByVert(mrrOnlyAchieved, thisQRange[0], thisQRange[1], verticalList), [mrrOnlyAchieved, thisQRange, verticalList]);
  const nextQAchievedByVert = useMemo(() => computeAchievedByVert(mrrOnlyAchieved, nextQRange[0], nextQRange[1], verticalList), [mrrOnlyAchieved, nextQRange, verticalList]);

  const thisQGapByVert = useMemo(() => {
    const out: any = {};
    for (const vert of verticalList) out[vert] = (thisQTargets[vert] || 0) - (thisQAchievedByVert[vert] || 0);
    return out;
  }, [verticalList, thisQTargets, thisQAchievedByVert]);

  const nextQGapByVert = useMemo(() => {
    const out: any = {};
    for (const vert of verticalList) {
      const carry = thisQGapByVert[vert] || 0;
      out[vert] = (nextQTargets[vert] || 0) + carry - (nextQAchievedByVert[vert] || 0);
    }
    return out;
  }, [verticalList, nextQTargets, nextQAchievedByVert, thisQGapByVert]);

  const filterOptions = useMemo(() => {
    const collect = (extractor: (d: any) => any) => {
      const s = new Set<string>();
      deals.forEach(d => { const v = extractor(d); if (v) s.add(v); });
      return Array.from(s).sort();
    };
    return {
      vertical:     [...INDIA_VERTICALS, ...US_VERTICALS].filter(l => deals.some(d => d.resolvedLabel === l)),
      stage:        collect(d => d.stage),
      status:       collect(d => d.status),
      typeOfSale:   collect(d => d.typeOfSale || '(blank)'),
      dealOwner:    collect(d => d.owner || '(blank)'),
      employeeSize: collect(d => d.employeeSize || '(blank)'),
      industry:     collect(d => d.industry || '(blank)'),
      productAsk:   collect(d => d.productAsk || '(blank)'),
      icpStatus:    collect(d => d.icpStatus || '(blank)'),
      currentOwner: collect(d => d.currentOwner || '(blank)'),
      sdInitiator:  collect(d => d.sdInitiator || '(blank)'),
      saInitiator:  collect(d => d.saInitiator || '(blank)'),
    };
  }, [deals]);

  const tableDealsFiltered = useMemo(() => {
    if (!tableSearch) return regionDeals;
    const q = tableSearch.toLowerCase();
    return regionDeals.filter(d => `${d.name || ''} ${d.account || ''} ${d.owner || ''} ${d.currentOwner || ''} ${d.sdInitiator || ''} ${d.saInitiator || ''} ${d.rawVertical} ${d.rawTeams} ${d.industry || ''} ${d.productAsk || ''}`.toLowerCase().includes(q));
  }, [regionDeals, tableSearch]);

  const tableDealsSorted = useMemo(() => sortDeals(tableDealsFiltered, tableSort), [tableDealsFiltered, tableSort]);

  useEffect(() => { setTablePage(0); }, [activeRegion, globalFilters, tableSearch, tablePageSize]);

  const pageCount = Math.max(1, Math.ceil(tableDealsSorted.length / tablePageSize));
  const safePage = Math.min(Math.max(0, tablePage), pageCount - 1);
  const pagedDeals = tableDealsSorted.slice(safePage * tablePageSize, (safePage + 1) * tablePageSize);

  const handleTableSort = (col: string) => setTableSort(s => s.column === col ? { column: col, direction: s.direction === 'asc' ? 'desc' : 'asc' } : { column: col, direction: 'desc' });

  const exportCSV = () => {
    const rows = tableDealsSorted.map(d => d._raw);
    const csv = Papa.unparse(rows);
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `pipeline_${activeRegion}_${new Date().toISOString().slice(0,10)}.csv`; a.click();
    URL.revokeObjectURL(url);
  };

  const missingCritical = colMap ? Object.entries({
    'Vertical': colMap.vertical, 'Stage': colMap.stage, 'Status': colMap.status,
    'MRR (Actual Total Deal MRR)': colMap.mrr, 'Pipe added date': colMap.pipeAdded,
  }).filter(([, v]) => !v).map(([k]) => k) : [];

  const setFilter = (key: string, value: any) => setGlobalFilters((prev: any) => ({ ...prev, [key]: value }));
  const removeFilterValue = (key: string, value: any) => setGlobalFilters((prev: any) => {
    const current = prev[key];
    if (Array.isArray(current)) return { ...prev, [key]: current.filter((v: any) => v !== value) };
    return { ...prev, [key]: { from: '', to: '' } };
  });
  const clearAllFilters = () => setGlobalFilters(EMPTY_FILTERS);
  const totalActiveFilters = Object.entries(globalFilters).reduce((a, [, v]) => {
    if (Array.isArray(v)) return a + v.length;
    if (typeof v === 'object' && v !== null) return a + ((v as any).from || (v as any).to ? 1 : 0);
    return a;
  }, 0);

  const unmappedCount = verticalBreakdown.filter(v => !v.resolvedLabel && !(v.rawVertical === '(blank)' && v.rawTeams === '(blank)')).reduce((a, v) => a + v.count, 0);
  const blankAllCount = verticalBreakdown.filter(v => v.rawVertical === '(blank)' && v.rawTeams === '(blank)').reduce((a, v) => a + v.count, 0);
  const viaTeamsCount = deals.filter(d => d.resolutionSource === 'teams').length;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="max-w-[1500px] mx-auto p-4 md:p-6 space-y-5">

        <div className="flex flex-wrap items-end justify-between gap-4">
          <div>
            <h1 className="text-2xl md:text-3xl font-bold tracking-tight text-slate-900">LeadSquared FY27 Dashboard</h1>
            <p className="text-sm text-slate-600 mt-1">
              {fyLabel(fy)} • Indian FY (Apr–Mar) • Current: <span className="text-slate-900 font-semibold">{cq} {fyLabel(fy)}</span>
            </p>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={fetchAutoData} title="Re-fetch the latest auto-synced data"
              className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-100 shadow-sm">
              <RefreshCw className={`w-3.5 h-3.5 ${autoSync.status === 'loading' ? 'animate-spin' : ''}`} />
              {autoSync.status === 'loading' ? 'Syncing…' : autoSync.status === 'ok' ? 'Synced' : autoSync.status === 'empty' ? 'No auto data' : 'Sync'}
            </button>
            <div className="flex items-center gap-1 bg-white border border-slate-200 rounded-lg p-1 shadow-sm">
              {['India', 'US'].map(r => (
                <button key={r} onClick={() => setActiveRegion(r)}
                  className={`px-4 py-1.5 rounded text-sm font-semibold transition ${activeRegion === r ? 'bg-sky-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'}`}>
                  {r}
                </button>
              ))}
            </div>
          </div>
        </div>

        {autoSync.status === 'ok' && autoSync.at && (
          <div className="text-[11px] text-emerald-700 bg-emerald-50 border border-emerald-200 rounded px-3 py-1.5 inline-flex items-center gap-1.5">
            <CheckCircle2 className="w-3.5 h-3.5" /> Auto-synced from LeadSquared • last checked {new Date(autoSync.at).toLocaleString()}
          </div>
        )}

        {/* Tab Bar */}
        <div className="flex gap-0 border-b border-slate-200 bg-white rounded-t-lg shadow-sm overflow-hidden">
          {[
            { id: 'pipeline', label: '📊 Pipeline Number View' },
            { id: 'mrr',      label: '📈 MRR Performance' },
            { id: 'renewal',  label: '🔄 Renewal Risks' },
          ].map(tab => (
            <button key={tab.id} onClick={() => setActiveTab(tab.id)}
              className={`px-5 py-3 text-sm font-semibold border-b-2 transition whitespace-nowrap ${
                activeTab === tab.id ? 'border-sky-600 text-sky-700 bg-sky-50' : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              }`}>
              {tab.label}
            </button>
          ))}
        </div>

        {/* File upload cards — manual fallback */}
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-3">
            <div className="bg-white border border-slate-200 rounded-lg p-3 flex flex-wrap items-center justify-between gap-2 shadow-sm">
              <div className="flex items-center gap-3 min-w-0">
                <FileSpreadsheet className="w-5 h-5 text-sky-600 flex-shrink-0" />
                <div className="min-w-0">
                  <div className="text-sm font-semibold text-slate-900 truncate">Opportunity Data</div>
                  <div className="text-xs text-slate-600 flex items-center gap-1">
                    {raw.length ? <><CheckCircle2 className="w-3 h-3 text-emerald-600 flex-shrink-0" /> <span className="truncate">{fileName} — {raw.length.toLocaleString()} deals</span></> : 'Auto-synced — upload to override'}
                  </div>
                  {oppUploadedAt && <div className="text-[10px] text-slate-500">Loaded {new Date(oppUploadedAt).toLocaleString()}</div>}
                </div>
              </div>
              <div className="flex items-center gap-2">
                {raw.length > 0 && <button onClick={() => setShowMap(!showMap)} className="text-xs text-slate-600 hover:text-slate-900 px-2 py-1">{showMap ? 'Hide' : 'Map'}</button>}
                <button onClick={() => fileRef.current?.click()} className="px-3 py-1.5 bg-sky-600 hover:bg-sky-700 text-white rounded text-sm font-medium shadow-sm">{raw.length ? 'Replace' : 'Upload'}</button>
                <input ref={fileRef} type="file" accept=".csv" className="hidden" onChange={(e) => handleFile(e.target.files?.[0])} />
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-lg p-3 flex flex-wrap items-center justify-between gap-2 shadow-sm">
              <div className="flex items-center gap-3 min-w-0">
                <FileText className="w-5 h-5 text-violet-600 flex-shrink-0" />
                <div className="min-w-0">
                  <div className="text-sm font-semibold text-slate-900">Target File (XLSX)</div>
                  <div className="text-xs text-slate-600 flex items-center gap-1">
                    {targetsSource ? <><CheckCircle2 className="w-3 h-3 text-emerald-600" /> {targetsSource}</> : 'Using built-in defaults'}
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={() => xlsxRef.current?.click()} className="px-3 py-1.5 bg-violet-600 hover:bg-violet-700 text-white rounded text-sm font-medium shadow-sm">{targetsSource ? 'Replace' : 'Upload'}</button>
                <input ref={xlsxRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleXlsx(e.target.files?.[0])} />
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-lg p-3 flex flex-wrap items-center justify-between gap-2 shadow-sm">
              <div className="flex items-center gap-3 min-w-0">
                <FileSpreadsheet className="w-5 h-5 text-emerald-700 flex-shrink-0" />
                <div className="min-w-0">
                  <div className="text-sm font-semibold text-slate-900">MRR Transaction Data</div>
                  <div className="text-xs text-slate-600 flex items-center gap-1">
                    {mrrFileName ? <><CheckCircle2 className="w-3 h-3 text-emerald-600" /> {mrrFileName} — {mrrRaw.length} NB / {mrrAllRaw.length} total</> : 'Auto-synced — upload to override'}
                  </div>
                  {mrrUploadedAt && <div className="text-[10px] text-slate-500">Loaded {new Date(mrrUploadedAt).toLocaleString()}</div>}
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={() => mrrFileRef.current?.click()} className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded text-sm font-medium shadow-sm">{mrrFileName ? 'Replace' : 'Upload'}</button>
                <input ref={mrrFileRef} type="file" accept=".csv" className="hidden" onChange={(e) => handleMrrFile(e.target.files?.[0])} />
              </div>
            </div>
            <div className="bg-white border border-slate-200 rounded-lg p-3 flex flex-wrap items-center justify-between gap-2 shadow-sm">
              <div className="flex items-center gap-3 min-w-0">
                <FileText className="w-5 h-5 text-rose-600 flex-shrink-0" />
                <div className="min-w-0">
                  <div className="text-sm font-semibold text-slate-900">Renewals Data</div>
                  <div className="text-xs text-slate-600 flex items-center gap-1">
                    {renewalFileName ? <><CheckCircle2 className="w-3 h-3 text-emerald-600" /> {renewalFileName} — {renewalRaw.length} accounts</> : 'Auto-synced — upload to override'}
                  </div>
                  {renewalUploadedAt && <div className="text-[10px] text-slate-500">Loaded {new Date(renewalUploadedAt).toLocaleString()}</div>}
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={() => renewalFileRef.current?.click()} className="px-3 py-1.5 bg-rose-600 hover:bg-rose-700 text-white rounded text-sm font-medium shadow-sm">{renewalFileName ? 'Replace' : 'Upload'}</button>
                <input ref={renewalFileRef} type="file" accept=".csv" className="hidden" onChange={(e) => handleRenewalFile(e.target.files?.[0])} />
              </div>
            </div>
        </div>

        {error && !raw.length && !mrrAllRaw.length && !renewalRaw.length && (
          <div className="bg-rose-50 border border-rose-200 rounded-lg p-3 text-sm text-rose-700">{error}</div>
        )}

        {activeTab === 'pipeline' && !raw.length && (
          <div className="bg-white border border-slate-200 rounded-xl p-12 text-center shadow-sm">
            <FileSpreadsheet className="w-10 h-10 mx-auto text-slate-400 mb-3" />
            <p className="text-base font-semibold text-slate-900">No Opportunity data loaded yet</p>
            <p className="text-xs text-slate-500 mt-2 max-w-md mx-auto">Data is fetched automatically from LeadSquared via the daily sync. If empty, click "Sync" above or upload a CSV manually.</p>
          </div>
        )}

        {missingCritical.length > 0 && raw.length > 0 && (
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 flex items-start gap-2">
            <AlertCircle className="w-4 h-4 text-amber-600 mt-0.5 flex-shrink-0" />
            <div className="text-sm">
              <span className="text-amber-900 font-semibold">Critical columns not auto-detected:</span>{' '}
              <span className="text-amber-800">{missingCritical.join(', ')}</span>
              <span className="text-slate-600"> — open "Map" to verify.</span>
            </div>
          </div>
        )}

        {raw.length > 0 && activeTab === 'pipeline' && (
          <div className="bg-white border border-slate-200 rounded-lg shadow-sm">
            <div className="px-4 py-3 flex items-center justify-between gap-3 border-b border-slate-100">
              <button onClick={() => setFiltersCollapsed(!filtersCollapsed)} className="flex items-center gap-2 text-sm font-semibold text-slate-900 hover:text-sky-700 transition">
                <Filter className="w-4 h-4 text-slate-500" /> Global Filters
                {totalActiveFilters > 0 && <span className="bg-sky-600 text-white text-[10px] font-bold rounded-full px-1.5 min-w-[20px] text-center">{totalActiveFilters}</span>}
                {filtersCollapsed ? <ChevronDown className="w-4 h-4 text-slate-500" /> : <ChevronUp className="w-4 h-4 text-slate-500" />}
              </button>
              <div className="flex items-center gap-3">
                <div className="text-xs text-slate-600">Filtered: <span className="font-semibold text-slate-900">{filteredDeals.length.toLocaleString()}</span> / {deals.length.toLocaleString()} deals</div>
                {totalActiveFilters > 0 && (
                  <button onClick={clearAllFilters} className="px-2.5 py-1 text-xs font-medium text-rose-700 hover:bg-rose-50 rounded flex items-center gap-1 transition">
                    <XCircle className="w-3.5 h-3.5" /> Clear all
                  </button>
                )}
              </div>
            </div>
            {!filtersCollapsed && (
              <div className="p-4 flex flex-wrap items-center gap-2">
                <MultiSelectFilter label={FILTER_LABELS.vertical}     options={filterOptions.vertical}     selected={globalFilters.vertical}     onChange={(v: any) => setFilter('vertical', v)} />
                <MultiSelectFilter label={FILTER_LABELS.stage}        options={filterOptions.stage}        selected={globalFilters.stage}        onChange={(v: any) => setFilter('stage', v)} />
                <MultiSelectFilter label={FILTER_LABELS.status}       options={filterOptions.status}       selected={globalFilters.status}       onChange={(v: any) => setFilter('status', v)} />
                <MultiSelectFilter label={FILTER_LABELS.typeOfSale}   options={filterOptions.typeOfSale}   selected={globalFilters.typeOfSale}   onChange={(v: any) => setFilter('typeOfSale', v)} />
                <MultiSelectFilter label={FILTER_LABELS.productAsk}   options={filterOptions.productAsk}   selected={globalFilters.productAsk}   onChange={(v: any) => setFilter('productAsk', v)} />
                <MultiSelectFilter label={FILTER_LABELS.dealOwner}    options={filterOptions.dealOwner}    selected={globalFilters.dealOwner}    onChange={(v: any) => setFilter('dealOwner', v)} />
                <MultiSelectFilter label={FILTER_LABELS.employeeSize} options={filterOptions.employeeSize} selected={globalFilters.employeeSize} onChange={(v: any) => setFilter('employeeSize', v)} />
                <MultiSelectFilter label={FILTER_LABELS.industry}     options={filterOptions.industry}     selected={globalFilters.industry}     onChange={(v: any) => setFilter('industry', v)} />
                <MultiSelectFilter label={FILTER_LABELS.icpStatus}    options={filterOptions.icpStatus}    selected={globalFilters.icpStatus}    onChange={(v: any) => setFilter('icpStatus', v)} />
                <MultiSelectFilter label={FILTER_LABELS.currentOwner} options={filterOptions.currentOwner} selected={globalFilters.currentOwner} onChange={(v: any) => setFilter('currentOwner', v)} />
                <MultiSelectFilter label={FILTER_LABELS.sdInitiator}  options={filterOptions.sdInitiator}  selected={globalFilters.sdInitiator}  onChange={(v: any) => setFilter('sdInitiator', v)} />
                <MultiSelectFilter label={FILTER_LABELS.saInitiator}  options={filterOptions.saInitiator}  selected={globalFilters.saInitiator}  onChange={(v: any) => setFilter('saInitiator', v)} />
                <DateRangeFilter label={FILTER_LABELS.pipeAddedRange}    value={globalFilters.pipeAddedRange}    onChange={(v: any) => setFilter('pipeAddedRange', v)} />
                <DateRangeFilter label={FILTER_LABELS.expClosureRange}   value={globalFilters.expClosureRange}   onChange={(v: any) => setFilter('expClosureRange', v)} />
                <DateRangeFilter label={FILTER_LABELS.finalClosureRange} value={globalFilters.finalClosureRange} onChange={(v: any) => setFilter('finalClosureRange', v)} />
              </div>
            )}
          </div>
        )}

        {raw.length > 0 && activeTab === 'pipeline' && (
          <div className="bg-white border border-slate-200 rounded-lg shadow-sm">
            <button onClick={() => setShowDiagnostics(!showDiagnostics)} className="w-full px-4 py-3 flex items-center justify-between hover:bg-slate-50 transition">
              <div className="flex items-center gap-2 flex-wrap">
                <Info className="w-4 h-4 text-slate-500" />
                <span className="text-sm font-semibold text-slate-900">Diagnostics</span>
                <span className="text-xs text-slate-500">
                  {verticalBreakdown.length} vertical/team combos
                  {viaTeamsCount > 0 && <span className="text-violet-700"> • {viaTeamsCount} via Teams</span>}
                  {unmappedCount > 0 && <span className="text-amber-700"> • {unmappedCount} unmapped</span>}
                  {blankAllCount > 0 && <span className="text-rose-700"> • {blankAllCount} fully blank</span>}
                  {unmappedCount === 0 && blankAllCount === 0 && <span className="text-emerald-700"> • all classified</span>}
                </span>
              </div>
              {showDiagnostics ? <ChevronUp className="w-4 h-4 text-slate-500" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
            </button>
            {showDiagnostics && (
              <div className="p-4 border-t border-slate-200 space-y-4">
                <div>
                  <div className="text-xs font-semibold text-slate-700 mb-2">Vertical Resolution (Vertical column → Teams column fallback)</div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
                    {verticalBreakdown.map((b, i) => {
                      const fullyBlank = b.rawVertical === '(blank)' && b.rawTeams === '(blank)';
                      const cls = b.resolvedLabel ? (b.source === 'teams' ? 'bg-violet-50 border-violet-200' : 'bg-emerald-50 border-emerald-200') : fullyBlank ? 'bg-rose-50 border-rose-200' : 'bg-amber-50 border-amber-200';
                      return (
                        <div key={i} className={`border rounded p-2 text-xs ${cls}`}>
                          <div className="space-y-0.5">
                            <div className="flex gap-1.5"><span className="text-slate-500 min-w-[60px]">Vertical:</span><span className="font-medium text-slate-900 truncate" title={b.rawVertical}>{b.rawVertical}</span></div>
                            <div className="flex gap-1.5"><span className="text-slate-500 min-w-[60px]">Teams:</span><span className="font-medium text-slate-900 truncate" title={b.rawTeams}>{b.rawTeams}</span></div>
                          </div>
                          <div className="mt-1.5 pt-1.5 border-t border-slate-200/60">
                            <span className="text-slate-700">{b.count} deals → </span>
                            {b.resolvedLabel ? <span className={b.source === 'teams' ? 'text-violet-700 font-semibold' : 'text-emerald-700 font-semibold'}>{b.resolvedLabel} <span className="font-normal text-[10px]">(via {b.source})</span></span> : <span className={fullyBlank ? 'text-rose-700 font-semibold' : 'text-amber-700 font-semibold'}>{fullyBlank ? 'fully blank' : 'unmapped'}</span>}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {extractedRows.length > 0 && (
                  <div>
                    <div className="text-xs font-semibold text-slate-700 mb-2">Target File Extraction</div>
                    <div className="border border-slate-200 rounded overflow-hidden">
                      <table className="w-full text-xs">
                        <thead className="bg-slate-50 text-[10px] uppercase text-slate-600">
                          <tr><th className="px-2 py-1.5 text-left font-semibold">Source Label</th><th className="px-2 py-1.5 text-left font-semibold">Mapped To</th><th className="px-2 py-1.5 text-left font-semibold">Region</th><th className="px-2 py-1.5 text-right font-semibold">Q1</th><th className="px-2 py-1.5 text-right font-semibold">Q2</th><th className="px-2 py-1.5 text-right font-semibold">Q3</th><th className="px-2 py-1.5 text-right font-semibold">Q4</th></tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {extractedRows.map((r, i) => (
                            <tr key={i}>
                              <td className="px-2 py-1.5 text-slate-700">{r.source}</td>
                              <td className="px-2 py-1.5 font-medium text-slate-900">{r.label}</td>
                              <td className="px-2 py-1.5 text-slate-600">{r.region}</td>
                              <td className="px-2 py-1.5 text-right tabular-nums">{r.Q1 ? fmt(r.Q1, r.region) : '—'}</td>
                              <td className="px-2 py-1.5 text-right tabular-nums">{r.Q2 ? fmt(r.Q2, r.region) : '—'}</td>
                              <td className="px-2 py-1.5 text-right tabular-nums">{r.Q3 ? fmt(r.Q3, r.region) : '—'}</td>
                              <td className="px-2 py-1.5 text-right tabular-nums">{r.Q4 ? fmt(r.Q4, r.region) : '—'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {showMap && colMap && (
                  <div>
                    <div className="text-xs font-semibold text-slate-700 mb-2">Column Mapping</div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 text-xs">
                      {Object.entries(colMap).map(([k, v]) => (
                        <div key={k} className="border border-slate-200 rounded p-2 bg-slate-50">
                          <div className="text-slate-500 mb-0.5 capitalize">{k.replace(/([A-Z])/g, ' $1').trim()}</div>
                          <div className={v ? 'text-slate-900 font-medium' : 'text-rose-700'}>{(v as string) || '— not found —'}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {raw.length > 0 && activeTab === 'pipeline' && (
          <div className="space-y-5">
            <ThisQuarterTable
              subtitle={`${cq} ${fyLabel(fy)} • ${fmtDate(thisQRange[0])} – ${fmtDate(thisQRange[1])}`}
              accent="from-sky-50 to-white" deals={regionDeals} achievedDeals={mrrOnlyAchieved}
              fy={fy} q={cq} region={activeRegion} verticalList={verticalList} targets={thisQTargets}
              onTargetChange={(v: string, val: number) => updateTarget(activeRegion, 'thisQ', v, val)} onDrill={openDrill}
            />
            <ForecastTable
              title="Next Quarter" subtitle={`${nq} ${fyLabel(nfy)} • ${fmtDate(nextQRange[0])} – ${fmtDate(nextQRange[1])}`}
              accent="from-violet-50 to-white" deals={regionDeals} achievedDeals={mrrOnlyAchieved}
              periodStart={nextQRange[0]} periodEnd={nextQRange[1]} qLabel={`${nq} ${fyLabel(nfy)}`}
              region={activeRegion} verticalList={verticalList} targets={nextQTargets} carryForward={thisQGapByVert}
              prevPeriodLabel={`${cq} ${fyLabel(fy)}`} onTargetChange={(v: string, val: number) => updateTarget(activeRegion, 'nextQ', v, val)} onDrill={openDrill}
            />
            <ForecastTable
              title="H2 / Next Half" subtitle={`${halfLabel} • ${fmtDate(halfStart)} – ${fmtDate(halfEnd)}`}
              accent="from-emerald-50 to-white" deals={regionDeals} achievedDeals={mrrOnlyAchieved}
              periodStart={halfStart} periodEnd={halfEnd} qLabel={halfLabel}
              region={activeRegion} verticalList={verticalList} targets={nextHTargets} carryForward={nextQGapByVert}
              prevPeriodLabel={`${nq} ${fyLabel(nfy)}`} onTargetChange={(v: string, val: number) => updateTarget(activeRegion, 'nextH', v, val)} onDrill={openDrill}
            />
          </div>
        )}

        {raw.length > 0 && activeTab === 'pipeline' && (
          <div className="bg-white border border-slate-200 rounded-lg shadow-sm">
            <div className="p-4 border-b border-slate-200">
              <div className="flex flex-wrap items-center gap-3">
                <div>
                  <h2 className="font-semibold text-slate-900">All {activeRegion} Deals</h2>
                  <p className="text-xs text-slate-600 mt-0.5">{tableDealsSorted.length.toLocaleString()} of {regionDeals.length.toLocaleString()} • Click any row for details • Click column headers to sort {totalActiveFilters > 0 && <span className="text-sky-700 font-medium">• {totalActiveFilters} global filter{totalActiveFilters > 1 ? 's' : ''} active</span>}</p>
                </div>
                <div className="flex-1" />
                <div className="flex items-center gap-2 flex-wrap">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input type="text" placeholder="Search deal / account / owner / industry / product..." value={tableSearch} onChange={(e) => setTableSearch(e.target.value)}
                      className="bg-white border border-slate-300 rounded pl-7 pr-2 py-1.5 text-sm w-72 focus:outline-none focus:border-sky-500" />
                  </div>
                  <button onClick={exportCSV} className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-900 rounded text-sm flex items-center gap-1 font-medium">
                    <Download className="w-3.5 h-3.5" /> Export
                  </button>
                </div>
              </div>
            </div>
            <ActiveFilterChips filters={globalFilters} onRemove={removeFilterValue} />
            <div className="overflow-auto" style={{ maxHeight: '70vh' }}>
              <SortableDealTable deals={pagedDeals} onRowClick={setDealDetail} region={activeRegion} sort={tableSort} onSort={handleTableSort} />
            </div>
            <PaginationControls page={safePage} pageSize={tablePageSize} total={tableDealsSorted.length} onPageChange={setTablePage} onPageSizeChange={setTablePageSize} />
          </div>
        )}

        {raw.length > 0 && activeTab === 'pipeline' && (
          <details className="bg-white border border-slate-200 rounded-lg shadow-sm">
            <summary className="p-3 text-sm font-semibold cursor-pointer flex items-center gap-2 hover:bg-slate-50 text-slate-900">
              <Info className="w-4 h-4 text-slate-500" /> Metric definitions & logic
            </summary>
            <div className="p-4 border-t border-slate-200 text-xs text-slate-700 space-y-2">
              <div><span className="text-emerald-700 font-semibold">MRR Achieved:</span> Stage = Won AND Status = Won AND MRR Start Date in the period.</div>
              <div><span className="text-sky-700 font-semibold">Open Pipe (This Q):</span> Open pipeline deals with Pipe Added Date AND Expected Closure Date in this quarter.</div>
              <div><span className="text-rose-700 font-semibold">Lost Pipe:</span> Stage = Lost AND Status = Lost AND <strong>both</strong> Expected Closure Date AND Final Closure Date fall in this quarter.</div>
              <div><span className="text-amber-700 font-semibold">Lapsed Pipe:</span> Subset of Open Pipe where Expected Closure Date is already past today.</div>
              <div><span className="text-sky-700 font-semibold">Pipe Due for Closure (Next Q / H2):</span> Open pipeline deals with Expected Closure Date in that future period.</div>
              <div><span className="text-violet-700 font-semibold">Pre-Pipe:</span> Current snapshot of deals in Discovery / DOQ / PSQ (not Won/Lost).</div>
              <div><span className="text-orange-700 font-semibold">Gap Carried Forward:</span> Previous period's MRR Gap (Target − Achieved). Positive = under-achievement carried into the next period.</div>
              <div><span className="text-rose-700 font-semibold">MRR Gap (Next Q / H2):</span> MRR Target + Gap Carried Forward − MRR Achieved.</div>
              <div className="pt-2 border-t border-slate-200 mt-2"><span className="text-slate-900 font-semibold">Data source:</span> Auto-synced daily from LeadSquared. Manual upload overrides the synced data for the current session.</div>
              <div><span className="text-slate-900 font-semibold">Target source:</span> {targetsSource || 'built-in defaults'}.</div>
            </div>
          </details>
        )}
      </div>

      {drill && <DrillModal drill={drill} region={drill.region || activeRegion} activeFilters={globalFilters} onRemoveFilter={removeFilterValue} onClose={() => setDrill(null)} onRowClick={setDealDetail} />}
      {dealDetail && <DealDetailModal deal={dealDetail} onClose={() => setDealDetail(null)} />}

      {activeTab === 'mrr' && <div className="max-w-[1500px] mx-auto px-4 md:px-6 pb-6 space-y-4"><MrrPerformanceTab mrrAllRaw={mrrAllRaw} region={activeRegion} fy={fy} q={cq} /></div>}
      {activeTab === 'renewal' && <div className="max-w-[1500px] mx-auto px-4 md:px-6 pb-6 space-y-4"><RenewalRisksTab renewalRaw={renewalRaw} region={activeRegion} fy={fy} q={cq} /></div>}
    </div>
  );
}
