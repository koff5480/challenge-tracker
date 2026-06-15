import { useState, useEffect, useCallback } from 'react';
import {
  RefreshCw, Settings, Plus, Minus,
  ChevronDown, ChevronUp, Copy, Check, X, Upload
} from 'lucide-react';

// ─── Apps Script Template ────────────────────────────────────
const APPS_SCRIPT = `// ================================================
// 挑戰追蹤器 · Google Apps Script 後端
// ================================================
// 設定步驟：
// 1. 開啟 Google Sheets，建立新試算表
// 2. 複製試算表 URL 中段的 ID（/d/ 和 /edit 之間那串）
// 3. 點「擴充功能 → Apps Script」，貼上此程式碼
// 4. 將下方 YOUR_GOOGLE_SHEET_ID 替換成你的試算表 ID
// 5. 執行一次 initSheets() 函式（點▶執行）
// 6. 部署 → 新增部署 → 網頁應用程式
//    「以何身分執行」= 我自己；「誰可以存取」= 所有人
// 7. 複製部署網址，貼到追蹤器設定頁
// ================================================

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // ← 填入你的試算表 ID

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'getData';
    if (action === 'getData') return getData();
    if (action === 'update') {
      const data = JSON.parse(decodeURIComponent(e.parameter.data));
      return updateData(data);
    }
    if (action === 'init') return initSheets();
    return resp({ error: 'Unknown action' });
  } catch (err) {
    return resp({ error: err.toString() });
  }
}

function resp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  function parseSheet(name) {
    const s = ss.getSheetByName(name);
    if (!s) return [];
    const rows = s.getDataRange().getValues();
    const headers = rows[0];
    return rows.slice(1).map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) {
        let v = row[i];
        if (v === 'TRUE') v = true;
        else if (v === 'FALSE') v = false;
        else if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
        obj[h] = v;
      });
      return obj;
    });
  }
  const cfgSheet = ss.getSheetByName('config');
  const config = {};
  if (cfgSheet) cfgSheet.getDataRange().getValues().forEach(function(r) { config[r[0]] = r[1]; });
  return resp({
    goal1Weeks: parseSheet('goal1'),
    goal2Weeks: parseSheet('goal2'),
    goal2EndDate: config['goal2EndDate'] || '2026-12-20',
    lastUpdated: new Date().toISOString()
  });
}

function updateData(p) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  function upSheet(name, wk, fields) {
    const s = ss.getSheetByName(name);
    const vals = s.getDataRange().getValues();
    const headers = vals[0];
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] == wk) {
        fields.forEach(function(f) {
          const col = headers.indexOf(f) + 1;
          const v = typeof p[f] === 'boolean' ? (p[f] ? 'TRUE' : 'FALSE') : (p[f] === undefined ? '' : p[f]);
          s.getRange(i + 1, col).setValue(v);
        });
        break;
      }
    }
  }
  if (p.type === 'goal1') upSheet('goal1', p.week, ['sweets','junk','isExempt','notes']);
  if (p.type === 'goal2') upSheet('goal2', p.week, ['outdoor','home','isGrace','notes']);
  if (p.type === 'config') {
    const s = ss.getSheetByName('config');
    const vals = s.getDataRange().getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === p.key) { s.getRange(i+1, 2).setValue(p.value); break; }
    }
  }
  const cfg = ss.getSheetByName('config');
  if (cfg) {
    const cv = cfg.getDataRange().getValues();
    for (let i = 0; i < cv.length; i++) {
      if (cv[i][0] === 'lastUpdated') { cfg.getRange(i+1,2).setValue(new Date().toISOString()); break; }
    }
  }
  return resp({ success: true });
}

function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const fmt = function(d) { return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd'); };
  let g1 = ss.getSheetByName('goal1');
  if (!g1) g1 = ss.insertSheet('goal1'); else g1.clearContents();
  g1.getRange(1,1,1,6).setValues([['week','startDate','sweets','junk','isExempt','notes']]);
  const g1S = new Date('2026-06-15');
  for (let i = 0; i < 10; i++) {
    const d = new Date(g1S.getTime() + i*7*86400000);
    g1.getRange(i+2,1,1,6).setValues([[i+1,fmt(d),0,0,'FALSE','']]);
  }
  let g2 = ss.getSheetByName('goal2');
  if (!g2) g2 = ss.insertSheet('goal2'); else g2.clearContents();
  g2.getRange(1,1,1,6).setValues([['week','startDate','outdoor','home','isGrace','notes']]);
  const g2S = new Date('2026-06-15'), g2E = new Date('2026-12-20');
  const nW = Math.ceil((g2E-g2S)/(7*86400000))+1;
  for (let i = 0; i < nW; i++) {
    const d = new Date(g2S.getTime() + i*7*86400000);
    g2.getRange(i+2,1,1,6).setValues([[i+1,fmt(d),0,0,'FALSE','']]);
  }
  let cfg = ss.getSheetByName('config');
  if (!cfg) cfg = ss.insertSheet('config'); else cfg.clearContents();
  cfg.getRange(1,1,4,2).setValues([
    ['goal1StartDate','2026-06-15'],
    ['goal2StartDate','2026-06-15'],
    ['goal2EndDate','2026-12-20'],
    ['lastUpdated',new Date().toISOString()]
  ]);
  return resp({ success: true, message: '初始化完成！' });
}`;

// ─── Types ────────────────────────────────────────────────────
interface G1Week { week: number; startDate: string; sweets: number; junk: number; isExempt: boolean; notes: string; }
interface G2Week { week: number; startDate: string; outdoor: number; home: number; isGrace: boolean; notes: string; }
interface AppData { g1: G1Week[]; g2: G2Week[]; g2End: string; }

// ─── Constants & Utils ────────────────────────────────────────
const G1_START = '2026-06-15';
const G2_START = '2026-06-15';
const MSW = 604800000; // ms per week

const dateOffset = (iso: string, weeks: number): string => {
  const [y, m, d] = iso.slice(0, 10).split('-').map(Number);
  const dt = new Date(y, m - 1, d + weeks * 7);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
};

const curWeek = (start: string, max: number): number =>
  Math.max(1, Math.min(max, Math.floor((Date.now() - new Date(start).getTime()) / MSW) + 1));

const eff = (o: number, h: number): number => o + Math.floor(h / 3);

// 居家運動跨週累積：前幾週未湊成整次的居家次數，帶入本週
const homeCarryIn = (weeks: G2Week[], weekNum: number): number =>
  weeks.filter(w => w.week < weekNum).reduce((sum, w) => sum + w.home, 0) % 3;

// 含跨週帶入的本週有效次數
const weekEff = (weeks: G2Week[], weekNum: number): number => {
  const w = weeks.find(x => x.week === weekNum);
  if (!w) return 0;
  return w.outdoor + Math.floor((homeCarryIn(weeks, weekNum) + w.home) / 3);
};
const monthOf = (d: string): string => d.slice(0, 7);
const fmtDate = (d: string): string => d.slice(5, 10).replace('-', '/');

// 從挑戰起始日 + 週序號計算本地時間範圍（避免時區問題）
const weekRange = (challengeStart: string, weekNum: number): string => {
  const [y, m, d] = challengeStart.split('-').map(Number);
  const mon = new Date(y, m - 1, d + (weekNum - 1) * 7);
  const sun = new Date(y, m - 1, d + (weekNum - 1) * 7 + 6);
  const f = (dt: Date) => `${dt.getMonth() + 1}/${dt.getDate()}`;
  return `${f(mon)} 至 ${f(sun)}`;
};
const weekMon = (challengeStart: string, weekNum: number): string => {
  const [y, m, d] = challengeStart.split('-').map(Number);
  const mon = new Date(y, m - 1, d + (weekNum - 1) * 7);
  return `${mon.getMonth() + 1}/${mon.getDate()}`;
};

type G1S = 'future' | 'current' | 'failed' | 'success' | 'exempt';
type G2S = 'future' | 'current' | 'success' | 'warning' | 'grace';

const g1Status = (w: G1Week, cur: number): G1S =>
  w.week > cur ? 'future'
  : w.isExempt ? 'exempt'
  : (w.sweets > 2 || w.junk > 2) ? 'failed'
  : w.week === cur ? 'current'
  : 'success';

const g2Status = (weeks: G2Week[], w: G2Week, cur: number): G2S =>
  w.week > cur ? 'future'
  : w.isGrace ? 'grace'
  : w.week === cur ? 'current'
  : weekEff(weeks, w.week) >= 2 ? 'success'
  : 'warning';

const calcStreak = (g1: G1Week[], cur: number): number => {
  let s = 0;
  for (const w of [...g1].filter(x => x.week < cur).sort((a, b) => a.week - b.week)) {
    const st = g1Status(w, cur);
    if (st === 'success' || st === 'exempt') s++;
    else if (st === 'failed') s = 0;
  }
  return s;
};

const initData = (g2End = '2026-12-20'): AppData => {
  const n2 = Math.ceil((new Date(g2End).getTime() - new Date(G2_START).getTime()) / MSW) + 1;
  return {
    g1: Array.from({ length: 10 }, (_, i) => ({ week: i+1, startDate: dateOffset(G1_START, i), sweets: 0, junk: 0, isExempt: false, notes: '' })),
    g2: Array.from({ length: n2 }, (_, i) => ({ week: i+1, startDate: dateOffset(G2_START, i), outdoor: 0, home: 0, isGrace: false, notes: '' })),
    g2End
  };
};

// 修正 localStorage 可能殘留的完整 ISO 日期字串（時區問題）
const cleanDate = (d: string): string => {
  if (!d || d.length <= 10) return d || '';
  const dt = new Date(d); // 當作 UTC 解析後，輸出本地日期
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
};

const loadData = (): AppData => {
  try {
    const s = localStorage.getItem('ct_v1');
    if (s) {
      const p = JSON.parse(s) as AppData;
      return {
        ...p,
        g1: (p.g1 || []).map(w => ({ ...w, startDate: cleanDate(w.startDate) })),
        g2: (p.g2 || []).map(w => ({ ...w, startDate: cleanDate(w.startDate) })),
        g2End: cleanDate(p.g2End || '2026-12-20'),
      };
    }
  } catch {}
  return initData();
};
const saveData = (d: AppData) => localStorage.setItem('ct_v1', JSON.stringify(d));

const apiGet = async (url: string): Promise<AppData> => {
  const r = await fetch(`${url}?action=getData`);
  const j = await r.json();
  return { g1: j.goal1Weeks || [], g2: j.goal2Weeks || [], g2End: j.goal2EndDate || '2026-12-20' };
};
const apiPush = (url: string, p: object) =>
  fetch(`${url}?action=update&data=${encodeURIComponent(JSON.stringify(p))}`).catch(() => {});

// ─── Status style maps ────────────────────────────────────────
const DOT: Record<string, string> = {
  future:  'bg-slate-200 text-slate-400',
  current: 'bg-indigo-500 text-white shadow-md shadow-indigo-200',
  success: 'bg-emerald-400 text-white',
  failed:  'bg-rose-400 text-white',
  exempt:  'bg-violet-400 text-white',
  warning: 'bg-amber-400 text-white',
  grace:   'bg-sky-400 text-white',
};
const BADGE: Record<string, [string, string]> = {
  future:  ['bg-slate-100 text-slate-400',           '—'],
  current: ['bg-indigo-100 text-indigo-600',         '⏳ 進行中'],
  success: ['bg-emerald-100 text-emerald-700',       '✓ 成功'],
  failed:  ['bg-rose-100 text-rose-600',             '✗ 失敗'],
  exempt:  ['bg-violet-100 text-violet-700',         '🛡 豁免'],
  warning: ['bg-amber-100 text-amber-700',           '⚠️ 未達標'],
  grace:   ['bg-sky-100 text-sky-700',               '💙 容錯'],
};

// ─── Sub-components ───────────────────────────────────────────
function Badge({ status }: { status: string }) {
  const [cls, label] = BADGE[status] || BADGE.future;
  return <span className={`text-xs font-semibold px-2.5 py-1 rounded-full whitespace-nowrap ${cls}`}>{label}</span>;
}

function Ctr({ icon, label, value, onChange, warn = false }: {
  icon: string; label: string; value: number; warn?: boolean;
  onChange: (v: number) => void;
}) {
  return (
    <div className={`flex items-center gap-3 p-3 rounded-2xl border-2 transition-all ${warn ? 'border-rose-300 bg-rose-50' : 'border-slate-100 bg-slate-50/80'}`}>
      <span className="text-2xl flex-shrink-0">{icon}</span>
      <div className="flex-1 min-w-0">
        <div className="text-sm font-medium text-slate-700 truncate">{label}</div>
        {warn && <div className="text-xs text-rose-500 font-bold">⚠️ 已超過上限！</div>}
      </div>
      <div className="flex items-center gap-2 flex-shrink-0">
        <button onClick={() => onChange(Math.max(0, value - 1))}
          style={{ touchAction: 'manipulation' }}
          className="w-11 h-11 rounded-full bg-white border border-slate-200 hover:bg-slate-100 flex items-center justify-center shadow-sm transition-colors active:scale-95">
          <Minus size={18} className="text-slate-600" />
        </button>
        <span className={`w-10 text-center font-bold text-2xl tabular-nums ${warn ? 'text-rose-500' : 'text-slate-800'}`}>{value}</span>
        <button onClick={() => onChange(value + 1)}
          style={{ touchAction: 'manipulation' }}
          className="w-11 h-11 rounded-full bg-white border border-slate-200 hover:bg-slate-100 flex items-center justify-center shadow-sm transition-colors active:scale-95">
          <Plus size={18} className="text-slate-600" />
        </button>
      </div>
    </div>
  );
}

// ─── Goal 1 Card ──────────────────────────────────────────────
function G1Card({ weeks, curW, activeW, onSelectW, onUpdate, showHist, onToggleHist }: {
  weeks: G1Week[]; curW: number; activeW: number;
  onSelectW: (w: number | null) => void;
  onUpdate: (week: number, p: Partial<G1Week>) => void;
  showHist: boolean; onToggleHist: () => void;
}) {
  // 成功週數：已結束的週，未超標且非豁免
  const successCount = weeks.filter(w => g1Status(w, curW) === 'success').length;
  // 目標週數：基礎 10 週 + 豁免次數
  const targetWeeks = weeks.length;
  const done = successCount >= 10;
  const w = weeks.find(x => x.week === activeW);
  const st = w ? g1Status(w, curW) : 'future';
  const isFuture = activeW > curW;

  return (
    <div className="bg-white rounded-3xl shadow-sm border border-pink-100 overflow-hidden flex flex-col">
      {/* Header */}
      <div className={`p-5 text-white ${done ? 'bg-gradient-to-br from-yellow-400 to-amber-500' : 'bg-gradient-to-br from-rose-400 to-pink-500'}`}>
        <div className="flex items-start justify-between">
          <div>
            <div className="text-sm font-medium opacity-75">目標一</div>
            <div className="text-xl font-bold leading-tight">許芸瑄 的飲食挑戰</div>
          </div>
          <div className="text-right flex-shrink-0 ml-2">
            <div className="text-4xl">{done ? '🏆' : '🇰🇷'}</div>
            <div className="text-xs opacity-75 mt-0.5">韓國旅遊</div>
          </div>
        </div>
        <div className="mt-4 flex gap-2">
          <div className="bg-white/20 rounded-2xl px-4 py-3 flex-1 text-center">
            <div className="text-3xl font-bold tabular-nums">{Math.min(successCount, 10)}/{targetWeeks}</div>
            <div className="text-xs opacity-80 mt-0.5">週次進度（集滿 10 次）</div>
          </div>
          {done && <div className="bg-white/30 rounded-2xl px-4 py-3 flex-1 text-center font-bold text-sm flex items-center justify-center">🎉 完成！</div>}
        </div>
        <div className="mt-3">
          <div className="h-2 bg-white/25 rounded-full overflow-hidden">
            <div className="h-full bg-white/90 rounded-full transition-all duration-700" style={{ width: `${Math.min(100, (successCount / 10) * 100)}%` }} />
          </div>
        </div>
      </div>

      {/* Week bubbles */}
      <div className="px-5 pt-4 pb-2">
        <div className="text-xs text-slate-400 font-semibold mb-2 uppercase tracking-wide">週次進度 — 點擊切換編輯</div>
        <div className="flex gap-2 flex-wrap">
          {weeks.map(wk => {
            const s = g1Status(wk, curW);
            const isAct = wk.week === activeW;
            return (
              <button key={wk.week}
                onClick={() => onSelectW(wk.week === curW ? null : wk.week)}
                title={`第${wk.week}週 ${weekRange(G1_START, wk.week)}`}
                style={{ touchAction: 'manipulation' }}
                className={`w-11 h-11 rounded-full text-sm font-bold transition-all ${DOT[s]} ${isAct ? 'ring-2 ring-offset-2 ring-current scale-110' : 'hover:scale-105'}`}>
                {wk.week}
              </button>
            );
          })}
        </div>
      </div>

      {/* Editor */}
      <div className="p-5 space-y-3 flex-1">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <span className="font-bold text-slate-800">第 {activeW} 週</span>
            <span className="text-slate-400 text-sm">{weekRange(G1_START, activeW)}</span>
            {activeW === curW && <span className="text-xs bg-indigo-100 text-indigo-600 px-2 py-0.5 rounded-full font-medium">本週</span>}
          </div>
          <Badge status={st} />
        </div>

        {isFuture ? (
          <div className="text-center py-6 text-slate-300 text-sm">此週尚未開始</div>
        ) : (
          <>
            <Ctr icon="🍰" label="甜點次數（上限 2 次）" value={w?.sweets ?? 0}
              warn={(w?.sweets ?? 0) > 2}
              onChange={v => w && onUpdate(activeW, { sweets: v })} />
            <Ctr icon="🍟" label="垃圾食物次數（上限 2 次）" value={w?.junk ?? 0}
              warn={(w?.junk ?? 0) > 2}
              onChange={v => w && onUpdate(activeW, { junk: v })} />
            <button
              onClick={() => w && onUpdate(activeW, { isExempt: !w.isExempt })}
              style={{ touchAction: 'manipulation' }}
              className={`w-full py-4 rounded-2xl border-2 font-semibold text-sm transition-all ${
                w?.isExempt
                  ? 'border-violet-400 bg-violet-50 text-violet-700'
                  : 'border-dashed border-slate-200 text-slate-400 hover:border-violet-300 hover:bg-violet-50 hover:text-violet-600'
              }`}>
              🛡&nbsp;{w?.isExempt ? '本週已標記豁免（點擊取消）' : '標記本週豁免'}
            </button>
            <p className="text-xs text-slate-400 text-center">豁免週不視為失敗，週數往後遞延（第 11 週起自動新增）</p>
            {/* Notes */}
            <div className="space-y-1.5 pt-1">
              <div className="text-xs font-semibold text-slate-400 uppercase tracking-wide">📝 本週備註</div>
              <textarea
                value={w?.notes ?? ''}
                onChange={e => w && onUpdate(activeW, { notes: e.target.value })}
                placeholder="記錄吃了什麼甜點、垃圾食物…"
                rows={2}
                className="w-full border border-slate-200 rounded-2xl px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-rose-200 focus:border-transparent resize-none placeholder-slate-300 text-slate-700"
              />
            </div>
          </>
        )}
      </div>

      {/* History */}
      <div className="border-t border-slate-100">
        <button onClick={onToggleHist}
          className="w-full px-5 py-3 flex items-center justify-between text-sm text-slate-400 hover:bg-slate-50 transition-colors">
          <span>📅 歷史紀錄</span>
          {showHist ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
        </button>
        {showHist && (
          <div className="px-5 pb-4">
            <div className="space-y-0.5 max-h-64 overflow-y-auto">
              {weeks.filter(wk => wk.week <= curW).reverse().map(wk => (
                <div key={wk.week} className="py-2 border-b border-slate-50 last:border-0">
                  <div className="flex items-center gap-2 text-xs">
                    <span className="font-semibold text-slate-600 flex-shrink-0 w-10">第{wk.week}週</span>
                    <span className="text-slate-400 flex-shrink-0">{weekRange(G1_START, wk.week)}</span>
                    <span className="flex-1 text-right text-slate-500">🍰{wk.sweets} 🍟{wk.junk}</span>
                    <Badge status={g1Status(wk, curW)} />
                  </div>
                  {wk.notes && (
                    <p className="text-xs text-slate-400 mt-0.5 pl-10 italic truncate">📝 {wk.notes}</p>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Goal 2 Card ──────────────────────────────────────────────
function G2Card({ weeks, curW, activeW, onSelectW, onUpdate, g2End, showHist, onToggleHist, onEditEnd }: {
  weeks: G2Week[]; curW: number; activeW: number;
  onSelectW: (w: number | null) => void;
  onUpdate: (week: number, p: Partial<G2Week>) => void;
  g2End: string; showHist: boolean;
  onToggleHist: () => void; onEditEnd: () => void;
}) {
  const daysLeft = Math.max(0, Math.ceil((new Date(g2End).getTime() - Date.now()) / 86400000));
  const done = daysLeft <= 0;
  // 有效達標週數：週次 ≤ 當前週，且有效次數 ≥ 2（當週達標即計；容錯週不計入，但遞延週數）
  const g2SuccessCount = weeks.filter(wk => wk.week <= curW && weekEff(weeks, wk.week) >= 2).length;
  const w = weeks.find(x => x.week === activeW);
  const st = w ? g2Status(weeks, w, curW) : 'future';
  const isFuture = activeW > curW;
  const carry = homeCarryIn(weeks, activeW);
  const totalHome = carry + (w?.home ?? 0);
  const effHome = Math.floor(totalHome / 3);
  const carryOut = totalHome % 3;
  const effNow = (w?.outdoor ?? 0) + effHome;
  const graceThisMonth = weeks.some(wk => wk.week !== activeW && wk.isGrace && monthOf(wk.startDate) === monthOf(w?.startDate ?? ''));

  // Month grace summary
  const currentMonth = w ? monthOf(w.startDate) : '';
  const graceCountThisMonth = weeks.filter(wk => wk.isGrace && monthOf(wk.startDate) === currentMonth).length;

  // View: weeks surrounding active
  const viewStart = Math.max(0, activeW - 4);
  const viewEnd = Math.min(weeks.length, activeW + 8);
  const viewWeeks = weeks.slice(viewStart, viewEnd);

  return (
    <div className="bg-white rounded-3xl shadow-sm border border-teal-100 overflow-hidden flex flex-col">
      {/* Header */}
      <div className={`p-5 text-white ${done ? 'bg-gradient-to-br from-yellow-400 to-amber-500' : 'bg-gradient-to-br from-emerald-400 to-teal-500'}`}>
        <div className="flex items-start justify-between">
          <div>
            <div className="text-sm font-medium opacity-75">目標二</div>
            <div className="text-xl font-bold leading-tight">黃冠鈞 的運動挑戰</div>
          </div>
          <div className="text-right flex-shrink-0 ml-2">
            <div className="text-4xl">{done ? '🏆' : '🏀'}</div>
            <div className="text-xs opacity-75 mt-0.5">美國 NBA 之旅</div>
          </div>
        </div>
        <div className="mt-4 flex gap-2">
          <div className="bg-white/20 rounded-2xl px-4 py-3 flex-1 text-center">
            <div className="text-3xl font-bold tabular-nums">{g2SuccessCount}/{weeks.length}</div>
            <div className="text-xs opacity-80 mt-0.5">週次進度</div>
          </div>
          <button onClick={onEditEnd}
            className="bg-white/20 hover:bg-white/30 rounded-2xl px-3 py-3 flex-1 text-center transition-colors">
            <div className="text-base font-bold">{g2End.slice(5).replace('-', '/')}</div>
            <div className="text-xs opacity-75 mt-0.5">截止日（可調整）</div>
          </button>
          {done && <div className="bg-white/30 rounded-2xl px-4 py-3 flex-1 text-center font-bold text-sm flex items-center justify-center">🏆 完成！</div>}
        </div>
        <div className="mt-3">
          <div className="h-2 bg-white/25 rounded-full overflow-hidden">
            <div className="h-full bg-white/90 rounded-full transition-all duration-700" style={{ width: `${Math.min(100, (g2SuccessCount / weeks.length) * 100)}%` }} />
          </div>
        </div>
      </div>

      {/* Week mini-dots */}
      <div className="px-5 pt-4 pb-2">
        <div className="text-xs text-slate-400 font-semibold mb-2 uppercase tracking-wide">近期週次 — 點擊切換編輯</div>
        <div className="flex gap-1.5 flex-wrap">
          {viewWeeks.map(wk => {
            const s = g2Status(weeks, wk, curW);
            const isAct = wk.week === activeW;
            return (
              <button key={wk.week}
                onClick={() => onSelectW(wk.week === curW ? null : wk.week)}
                title={`第${wk.week}週 ${weekRange(G2_START, wk.week)}`}
                style={{ touchAction: 'manipulation' }}
                className={`w-10 h-10 rounded-full text-xs font-bold transition-all ${DOT[s]} ${isAct ? 'ring-2 ring-offset-2 ring-current scale-110' : 'hover:scale-105'}`}>
                {wk.week}
              </button>
            );
          })}
        </div>
      </div>

      {/* Editor */}
      <div className="p-5 space-y-3 flex-1">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <span className="font-bold text-slate-800">第 {activeW} 週</span>
            <span className="text-slate-400 text-sm">{weekRange(G2_START, activeW)}</span>
            {activeW === curW && <span className="text-xs bg-indigo-100 text-indigo-600 px-2 py-0.5 rounded-full font-medium">本週</span>}
          </div>
          <Badge status={st} />
        </div>

        {isFuture ? (
          <div className="text-center py-6 text-slate-300 text-sm">此週尚未開始</div>
        ) : (
          <>
            {/* Effective count display */}
            <div className={`p-4 rounded-2xl text-center border-2 transition-all ${
              effNow >= 2 ? 'border-emerald-200 bg-emerald-50'
              : effNow > 0 ? 'border-amber-100 bg-amber-50'
              : 'border-slate-100 bg-slate-50'
            }`}>
              <div className={`text-5xl font-bold tabular-nums ${
                effNow >= 2 ? 'text-emerald-500' : 'text-slate-300'
              }`}>{effNow}</div>
              <div className="text-sm text-slate-500 mt-1">有效次數 / 目標 2 次</div>
              <div className="mt-2 text-xs text-slate-400 space-y-1 text-left px-1">
                <div className="flex items-center gap-1">
                  <span>🏃</span>
                  <span>外出 <strong className="text-slate-600">{w?.outdoor ?? 0}</strong> 次（各計 1 有效）</span>
                </div>
                <div className="flex items-center gap-1 flex-wrap">
                  <span>🏠</span>
                  <span>居家：
                    {carry > 0
                      ? <><span className="text-sky-500 font-semibold">結餘 {carry}</span> + 本週 <strong className="text-slate-600">{w?.home ?? 0}</strong></>
                      : <>本週 <strong className="text-slate-600">{w?.home ?? 0}</strong></>
                    }
                    {' '}= {totalHome} 次 → <strong className="text-slate-600">{effHome}</strong> 有效
                  </span>
                </div>
                {carryOut > 0 && (
                  <div className="flex items-center gap-1 text-sky-500 font-semibold">
                    <span>↪</span><span>本週結餘 {carryOut} 次居家，帶入下週</span>
                  </div>
                )}
              </div>
              {effNow >= 2 && <div className="text-emerald-600 font-semibold text-sm mt-2">✓ 本週已達標！</div>}
              {effNow < 2 && activeW < curW && !w?.isGrace && (
                <div className="text-amber-600 text-sm mt-2">⚠️ 未達標（警告，進度不重置）</div>
              )}
            </div>

            <Ctr icon="🏃" label="外出運動（1 次 = 1 有效次）" value={w?.outdoor ?? 0}
              onChange={v => w && onUpdate(activeW, { outdoor: v })} />
            <Ctr icon="🏠"
              label={carry > 0
                ? `居家運動（帶入結餘 ${carry} 次，再加 ${3 - carry} 次可湊 1 有效）`
                : '居家運動（累計 3 次 = 1 有效，可跨週累積）'}
              value={w?.home ?? 0}
              onChange={v => w && onUpdate(activeW, { home: v })} />

            {/* Grace button */}
            <button
              disabled={!w?.isGrace && graceThisMonth}
              onClick={() => w && onUpdate(activeW, { isGrace: !w.isGrace })}
              style={{ touchAction: 'manipulation' }}
              className={`w-full py-4 rounded-2xl border-2 font-semibold text-sm transition-all ${
                w?.isGrace ? 'border-sky-400 bg-sky-50 text-sky-700'
                : graceThisMonth ? 'border-slate-100 bg-slate-50 text-slate-300 cursor-not-allowed'
                : 'border-dashed border-slate-200 text-slate-400 hover:border-sky-300 hover:bg-sky-50 hover:text-sky-600'
              }`}>
              💙&nbsp;{w?.isGrace
                ? '本週容錯中（點擊取消）'
                : graceThisMonth ? `本月容錯已使用（${graceCountThisMonth}/1）`
                : '標記本週容錯（每月限用 1 次）'}
            </button>
            <p className="text-xs text-slate-400 text-center">容錯週不計進度，但挑戰週數往後遞延一週</p>
            {/* Notes */}
            <div className="space-y-1.5 pt-1">
              <div className="text-xs font-semibold text-slate-400 uppercase tracking-wide">📝 本週備註</div>
              <textarea
                value={w?.notes ?? ''}
                onChange={e => w && onUpdate(activeW, { notes: e.target.value })}
                placeholder="記錄做了什麼運動、去哪裡運動…"
                rows={2}
                className="w-full border border-slate-200 rounded-2xl px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-teal-200 focus:border-transparent resize-none placeholder-slate-300 text-slate-700"
              />
            </div>
          </>
        )}
      </div>

      {/* History */}
      <div className="border-t border-slate-100">
        <button onClick={onToggleHist}
          className="w-full px-5 py-3 flex items-center justify-between text-sm text-slate-400 hover:bg-slate-50 transition-colors">
          <span>📅 歷史紀錄</span>
          {showHist ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
        </button>
        {showHist && (
          <div className="px-5 pb-4">
            <div className="space-y-0.5 max-h-64 overflow-y-auto">
              {weeks.filter(wk => wk.week <= curW).reverse().map(wk => (
                <div key={wk.week} className="py-2 border-b border-slate-50 last:border-0">
                  <div className="flex items-center gap-2 text-xs">
                    <span className="font-semibold text-slate-600 flex-shrink-0 w-10">第{wk.week}週</span>
                    <span className="text-slate-400 flex-shrink-0">{weekRange(G2_START, wk.week)}</span>
                    <span className="flex-1 text-right text-slate-500">
                      🏃{wk.outdoor} 🏠{wk.home}
                      {homeCarryIn(weeks, wk.week) > 0 && <span className="text-sky-400 ml-0.5">+{homeCarryIn(weeks, wk.week)}</span>}
                      <span className="ml-1">→{weekEff(weeks, wk.week)}效</span>
                    </span>
                    <Badge status={g2Status(weeks, wk, curW)} />
                  </div>
                  {wk.notes && (
                    <p className="text-xs text-slate-400 mt-0.5 pl-10 italic truncate">📝 {wk.notes}</p>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Setup Modal ──────────────────────────────────────────────
function SetupModal({ tmpUrl, setTmpUrl, onSave, onClose, copied, onCopy }: {
  tmpUrl: string; setTmpUrl: (s: string) => void;
  onSave: () => void; onClose: () => void;
  copied: boolean; onCopy: () => void;
}) {
  return (
    <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
        <div className="sticky top-0 bg-white border-b border-slate-100 px-6 py-4 flex items-center justify-between rounded-t-3xl">
          <h2 className="text-lg font-bold text-slate-800">⚙️ 設定 Google Sheets 雲端同步</h2>
          <button onClick={onClose} className="p-2 hover:bg-slate-100 rounded-xl transition-colors"><X size={16} className="text-slate-500" /></button>
        </div>
        <div className="p-6 space-y-6">
          {/* URL input */}
          <div>
            <label className="text-sm font-semibold text-slate-700 block mb-2">Apps Script 部署網址</label>
            <div className="flex gap-2">
              <input value={tmpUrl} onChange={e => setTmpUrl(e.target.value)}
                placeholder="https://script.google.com/macros/s/.../exec"
                className="flex-1 border border-slate-200 rounded-2xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-300 focus:border-transparent" />
              <button onClick={onSave}
                className="px-5 py-2.5 bg-indigo-500 hover:bg-indigo-600 text-white rounded-2xl text-sm font-semibold transition-colors">
                連接
              </button>
            </div>
          </div>

          {/* Steps */}
          <div>
            <h3 className="text-sm font-bold text-slate-700 mb-3">📋 設定步驟</h3>
            <ol className="space-y-3">
              {[
                <>開啟 <a href="https://sheets.google.com" target="_blank" rel="noreferrer" className="text-indigo-500 underline font-medium">Google Sheets</a>，建立新試算表，從 URL 複製試算表 ID（<code className="bg-slate-100 px-1 py-0.5 rounded text-slate-600">/d/</code> 和 <code className="bg-slate-100 px-1 py-0.5 rounded text-slate-600">/edit</code> 之間那串）</>,
                <>在試算表中點選 <strong>擴充功能 → Apps Script</strong>，把下方程式碼全選貼入</>,
                <>將程式碼中的 <code className="bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded font-mono">YOUR_GOOGLE_SHEET_ID</code> 替換成你的試算表 ID</>,
                <>點左上角執行 <code className="bg-slate-100 px-1 py-0.5 rounded">initSheets</code> 函式（第一次需授權）</>,
                <>點「部署 → 新增部署」→ 類型選<strong>網頁應用程式</strong>，以何身分執行選<strong>我自己</strong>，誰可以存取選<strong>所有人</strong>，完成後複製部署網址</>,
                <>兩人各自開啟此追蹤器，在上方欄位貼入相同的網址，點「連接」即可共用同步！</>
              ].map((step, i) => (
                <li key={i} className="flex gap-3 text-sm text-slate-600">
                  <span className="bg-indigo-100 text-indigo-600 rounded-full w-6 h-6 flex items-center justify-center text-xs font-bold flex-shrink-0 mt-0.5">{i+1}</span>
                  <span>{step}</span>
                </li>
              ))}
            </ol>
          </div>

          {/* Code block */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-sm font-bold text-slate-700">📄 Apps Script 程式碼</h3>
              <button onClick={onCopy}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-semibold transition-colors ${
                  copied ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                }`}>
                {copied ? <><Check size={12} /> 已複製！</> : <><Copy size={12} /> 複製程式碼</>}
              </button>
            </div>
            <pre className="bg-slate-50 border border-slate-200 rounded-2xl p-4 text-xs text-slate-700 overflow-x-auto max-h-60 overflow-y-auto font-mono whitespace-pre leading-relaxed">
              {APPS_SCRIPT}
            </pre>
          </div>

          <div className="bg-amber-50 border border-amber-200 rounded-2xl p-4 text-sm text-amber-700">
            💡 <strong>不想設定雲端同步？</strong>沒問題！資料會自動存在瀏覽器本機，完全可以使用。若需要兩人共享，再照上面步驟設定即可。
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── End Date Modal ───────────────────────────────────────────
function EndDateModal({ current, onSave, onClose }: {
  current: string;
  onSave: (d: string) => void;
  onClose: () => void;
}) {
  const [sel, setSel] = useState(current);
  const opts = [
    { val: '2026-12-20', label: '2026/12/20（週日）', note: '原始截止日' },
    { val: '2027-01-10', label: '2027/1/10（週日）', note: '延長版截止日' },
  ];
  return (
    <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-sm p-6">
        <h3 className="font-bold text-slate-800 text-lg mb-4">📅 調整目標截止日</h3>
        <div className="space-y-2 mb-5">
          {opts.map(o => (
            <button key={o.val} onClick={() => setSel(o.val)}
              className={`w-full p-4 rounded-2xl border-2 text-left transition-all ${
                sel === o.val ? 'border-emerald-400 bg-emerald-50' : 'border-slate-200 hover:border-slate-300'
              }`}>
              <div className="font-bold text-slate-800">{o.label}</div>
              <div className="text-sm text-slate-500 mt-0.5">{o.note}</div>
            </button>
          ))}
        </div>
        <div className="flex gap-2">
          <button onClick={onClose} className="flex-1 py-3 rounded-2xl border border-slate-200 text-slate-600 text-sm font-medium hover:bg-slate-50 transition-colors">取消</button>
          <button onClick={() => onSave(sel)} className="flex-1 py-3 rounded-2xl bg-emerald-500 hover:bg-emerald-600 text-white font-bold text-sm transition-colors">確認</button>
        </div>
      </div>
    </div>
  );
}

// ─── App ──────────────────────────────────────────────────────
export default function App() {
  const [data, setData] = useState<AppData>(loadData);
  const [url, setUrl] = useState(() => localStorage.getItem('scriptUrl') || '');
  const [tmpUrl, setTmpUrl] = useState(url);
  const [loading, setLoading] = useState(false);
  const [lastSync, setLastSync] = useState<Date | null>(null);
  const [syncErr, setSyncErr] = useState<string | null>(null);
  const [showSetup, setShowSetup] = useState(false);
  const [showH1, setShowH1] = useState(false);
  const [showH2, setShowH2] = useState(false);
  const [sw1, setSw1] = useState<number | null>(null);
  const [sw2, setSw2] = useState<number | null>(null);
  const [copied, setCopied] = useState(false);
  const [showEndEdit, setShowEndEdit] = useState(false);

  const cw1 = curWeek(G1_START, data.g1.length);
  const cw2 = curWeek(G2_START, data.g2.length);
  const aw1 = sw1 ?? cw1;
  const aw2 = sw2 ?? cw2;

  const sync = useCallback(async () => {
    if (!url) return;
    setLoading(true);
    try {
      const remote = await apiGet(url);
      setData(prev => {
        // 合併遠端資料：保護本地備注不被遠端空值覆蓋
        const merged: AppData = {
          ...remote,
          g1: remote.g1.length > 0
            ? remote.g1.map(rw => {
                const lw = prev.g1.find(w => w.week === rw.week);
                return { ...rw, notes: rw.notes || lw?.notes || '' };
              })
            : prev.g1,
          g2: remote.g2.length > 0
            ? remote.g2.map(rw => {
                const lw = prev.g2.find(w => w.week === rw.week);
                return { ...rw, notes: rw.notes || lw?.notes || '' };
              })
            : prev.g2,
        };
        saveData(merged);
        return merged;
      });
      setLastSync(new Date()); setSyncErr(null);
    } catch { setSyncErr('同步失敗，請確認 Apps Script URL'); }
    finally { setLoading(false); }
  }, [url]);

  useEffect(() => {
    if (url) { sync(); const t = setInterval(sync, 30000); return () => clearInterval(t); }
  }, [sync]);

  // 將本機所有資料（含備注）逐週推送到 Sheets
  // 使用情境：備注只存在電腦 localStorage、手機看不到時，按此鍵強制同步
  const pushAllData = useCallback(async () => {
    if (!url) return;
    setLoading(true);
    setSyncErr(null);
    try {
      for (const wk of data.g1) {
        await apiPush(url, { type: 'goal1', ...wk });
      }
      for (const wk of data.g2) {
        await apiPush(url, { type: 'goal2', ...wk });
      }
      // 推送完後同步驗證
      await sync();
    } catch { setSyncErr('上傳失敗，請檢查網路'); }
    finally { setLoading(false); }
  }, [url, data, sync]);

  const upG1 = useCallback((week: number, p: Partial<G1Week>) => {
    // 使用 functional setData 避免 stale closure 問題（快速輸入備注時不丟字）
    setData(prev => {
      const updated = prev.g1.map(w => w.week === week ? { ...w, ...p } : w);
      // 豁免週往後遞延：每有一週豁免，總週數 +1
      const exemptCount = updated.filter(w => w.isExempt).length;
      const neededWeeks = 10 + exemptCount;
      let g1 = updated;
      while (g1.length < neededWeeks) {
        const n = g1.length + 1;
        g1 = [...g1, { week: n, startDate: dateOffset(G1_START, n - 1), sweets: 0, junk: 0, isExempt: false, notes: '' }];
      }
      const d = { ...prev, g1 };
      saveData(d);
      if (url) {
        const wk = d.g1.find(w => w.week === week);
        if (wk) apiPush(url, { type: 'goal1', ...wk });
      }
      return d;
    });
  }, [url]);

  const upG2 = useCallback((week: number, p: Partial<G2Week>) => {
    // 使用 functional setData 避免 stale closure 問題（快速輸入備注時不丟字）
    setData(prev => {
      const updated = prev.g2.map(w => w.week === week ? { ...w, ...p } : w);
      // 容錯週往後遞延：以截止日算出基礎週數，再加上容錯次數
      const baseWeeks = Math.ceil((new Date(prev.g2End).getTime() - new Date(G2_START).getTime()) / MSW) + 1;
      const graceCount = updated.filter(w => w.isGrace).length;
      const neededWeeks = baseWeeks + graceCount;
      let g2 = updated;
      while (g2.length < neededWeeks) {
        const n = g2.length + 1;
        g2 = [...g2, { week: n, startDate: dateOffset(G2_START, n - 1), outdoor: 0, home: 0, isGrace: false, notes: '' }];
      }
      const d = { ...prev, g2 };
      saveData(d);
      if (url) {
        const wk = d.g2.find(w => w.week === week);
        if (wk) apiPush(url, { type: 'goal2', ...wk });
      }
      return d;
    });
  }, [url]);

  const applyEndDate = (newEnd: string) => {
    setData(prev => {
      const n = Math.ceil((new Date(newEnd).getTime() - new Date(G2_START).getTime()) / MSW) + 1;
      const map = Object.fromEntries(prev.g2.map(w => [w.week, w]));
      const g2 = Array.from({ length: n }, (_, i) => map[i+1] || { week: i+1, startDate: dateOffset(G2_START, i), outdoor: 0, home: 0, isGrace: false, notes: '' });
      const d = { ...prev, g2, g2End: newEnd };
      saveData(d);
      return d;
    });
    setShowEndEdit(false);
    if (url) apiPush(url, { type: 'config', key: 'goal2EndDate', value: newEnd });
  };

  const saveUrl = () => {
    localStorage.setItem('scriptUrl', tmpUrl);
    setUrl(tmpUrl); setShowSetup(false);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-violet-50 via-white to-teal-50"
      style={{ paddingBottom: 'max(2.5rem, env(safe-area-inset-bottom, 2.5rem))' }}>
      {/* Header */}
      <div className="sticky top-0 z-20 bg-white/80 backdrop-blur-md border-b border-slate-100">
        <div className="max-w-5xl mx-auto px-4 py-3 flex items-center justify-between">
          <div>
            <div className="font-bold text-slate-800 text-base">🎯 挑戰追蹤器</div>
            <div className="text-xs">
              {syncErr
                ? <span className="text-rose-400">{syncErr}</span>
                : lastSync
                  ? <span className="text-emerald-500">✓ 已同步 {lastSync.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit' })}</span>
                  : url
                    ? <span className="text-slate-400">連線中…</span>
                    : <span className="text-slate-400">離線模式（本機儲存）</span>}
            </div>
          </div>
          <div className="flex gap-1.5">
            {url && (
              <>
                <button onClick={pushAllData} title="將本機備注全部推送到 Sheets（手機看不到備注時按此）"
                  className="p-2.5 hover:bg-amber-50 rounded-xl transition-colors group relative">
                  <Upload size={15} className={loading ? 'text-amber-400' : 'text-amber-400 group-hover:text-amber-500'} />
                </button>
                <button onClick={sync} className="p-2.5 hover:bg-slate-100 rounded-xl transition-colors">
                  <RefreshCw size={15} className={loading ? 'animate-spin text-indigo-500' : 'text-slate-400'} />
                </button>
              </>
            )}
            <button onClick={() => setShowSetup(true)} className="p-2.5 hover:bg-slate-100 rounded-xl transition-colors">
              <Settings size={15} className="text-slate-400" />
            </button>
          </div>
        </div>
      </div>

      {/* Main grid */}
      <div className="max-w-5xl mx-auto p-4 md:p-6">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <G1Card
            weeks={data.g1} curW={cw1} activeW={aw1}
            onSelectW={w => setSw1(w)}
            onUpdate={upG1}
            showHist={showH1} onToggleHist={() => setShowH1(h => !h)}
          />
          <G2Card
            weeks={data.g2} curW={cw2} activeW={aw2}
            onSelectW={w => setSw2(w)}
            onUpdate={upG2}
            g2End={data.g2End}
            showHist={showH2} onToggleHist={() => setShowH2(h => !h)}
            onEditEnd={() => setShowEndEdit(true)}
          />
        </div>

        {/* Legend */}
        <div className="mt-6 flex flex-wrap gap-x-4 gap-y-2 justify-center">
          {Object.entries(BADGE).filter(([k]) => k !== 'future').map(([k, [cls, label]]) => (
            <div key={k} className="flex items-center gap-1.5 text-xs">
              <span className={`w-2 h-2 rounded-full ${DOT[k]}`} />
              <span className="text-slate-500">{label.replace(/[✓✗⚠️🛡💙⏳]/g, '').trim()}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Modals */}
      {showSetup && (
        <SetupModal tmpUrl={tmpUrl} setTmpUrl={setTmpUrl} onSave={saveUrl}
          onClose={() => setShowSetup(false)} copied={copied}
          onCopy={() => { navigator.clipboard.writeText(APPS_SCRIPT); setCopied(true); setTimeout(() => setCopied(false), 2500); }} />
      )}
      {showEndEdit && (
        <EndDateModal current={data.g2End} onSave={applyEndDate} onClose={() => setShowEndEdit(false)} />
      )}
    </div>
  );
}
