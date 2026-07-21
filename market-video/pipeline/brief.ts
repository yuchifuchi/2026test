/**
 * 分析層: 取得データから「今朝の海外→日本株」ブリーフJSONを生成
 * 数値はすべてここで機械計算する(台本・映像はこのJSONだけを参照 = 捏造防止)
 */
import * as fs from 'fs';
import * as path from 'path';
import { fetchAll, SeriesData, SymbolDef } from './fetch';

const SYMBOLS: SymbolDef[] = [
  { symbol: '^GSPC', stooq: '^spx', label: 'S&P500' },
  { symbol: '^IXIC', stooq: '^ndq', label: 'NASDAQ' },
  { symbol: 'USDJPY=X', stooq: 'usdjpy', label: 'ドル円' },
  { symbol: 'NKD=F', stooq: 'nkd.f', label: '日経平均先物' },
  { symbol: '^N225', stooq: '^nkx', label: '日経平均(前日終値)' },
];

const SECTOR_SYMBOLS: SymbolDef[] = [
  { symbol: 'XLK', stooq: 'xlk.us', label: 'ハイテク' },
  { symbol: 'XLY', stooq: 'xly.us', label: '一般消費財' },
  { symbol: 'XLP', stooq: 'xlp.us', label: '生活必需品' },
  { symbol: 'XLF', stooq: 'xlf.us', label: '金融' },
  { symbol: 'XLE', stooq: 'xle.us', label: 'エネルギー' },
  { symbol: 'XLV', stooq: 'xlv.us', label: 'ヘルスケア' },
  { symbol: 'XLI', stooq: 'xli.us', label: '資本財' },
];

// 米セクター→東京の関連業種(表示用の対応表)
const SECTOR_TO_JP: Record<string, string> = {
  'ハイテク': '半導体・電子部品',
  '一般消費財': '小売・自動車',
  '生活必需品': '食品・日用品',
  '金融': '銀行・保険',
  'エネルギー': '石油・商社',
  'ヘルスケア': '医薬品',
  '資本財': '機械・重工',
};

export interface SectorRank {
  label: string;
  changePct: number;
  jpNote: string;
}

export interface PanelData {
  label: string;
  latest: number;
  prev: number;
  changePct: number;
  latestDate: string;
  spark: number[]; // 直近20営業日の終値
  source: string;
}

export interface Brief {
  generatedAt: string;
  dateLabel: string;
  panels: PanelData[];
  headline: string;
  sectors: SectorRank[];
  sectorNote: string;
  sources: string[];
}

function toPanel(s: SeriesData): PanelData {
  const q = s.quotes;
  const latest = q[q.length - 1];
  const prev = q[q.length - 2];
  const changePct = ((latest.close - prev.close) / prev.close) * 100;
  return {
    label: s.label,
    latest: latest.close,
    prev: prev.close,
    changePct,
    latestDate: latest.date,
    spark: q.slice(-20).map((x) => x.close),
    source: s.source === 'yahoo' ? 'Yahoo Finance' : 'Stooq',
  };
}

// ルールベースの見出し(LLM不使用: 日次ブリーフは定型で十分)
function headline(panels: PanelData[]): string {
  const sp = panels.find((p) => p.label === 'S&P500');
  const sox = panels.find((p) => p.label.startsWith('SOX'));
  if (!sp) return '今朝の海外市場チェック';
  const dir = sp.changePct >= 0 ? '上昇' : '下落';
  const soxNote = sox && Math.abs(sox.changePct) > Math.abs(sp.changePct) + 0.5
    ? `、半導体は${sox.changePct >= 0 ? 'さらに強い' : 'より弱い'}`
    : '';
  return `米国株は${dir}(${sp.changePct >= 0 ? '+' : ''}${sp.changePct.toFixed(2)}%)${soxNote}`;
}

export async function buildBrief(): Promise<Brief> {
  const series = await fetchAll(SYMBOLS);
  if (series.length < 3) throw new Error('データ取得が不十分です(3系列未満)');
  const panels = series.map(toPanel);
  const sectorSeries = await fetchAll(SECTOR_SYMBOLS);
  const sectors: SectorRank[] = sectorSeries
    .map(toPanel)
    .map((p) => ({ label: p.label, changePct: p.changePct, jpNote: SECTOR_TO_JP[p.label] ?? '' }))
    .sort((a, b) => b.changePct - a.changePct);
  const top = sectors[0];
  const sectorNote = top
    ? `昨夜の米国は${top.label}が最強(${top.changePct >= 0 ? '+' : ''}${top.changePct.toFixed(2)}%)。東京は${top.jpNote}の動きに注目`
    : '';
  const brief: Brief = {
    generatedAt: new Date().toISOString(),
    dateLabel: new Date().toLocaleDateString('ja-JP', { month: 'long', day: 'numeric', weekday: 'short' }),
    panels,
    headline: headline(panels),
    sectors: sectors.slice(0, 3),
    sectorNote,
    sources: [...new Set(panels.map((p) => p.source))],
  };
  const out = path.join(__dirname, '..', 'data', 'brief.json');
  fs.writeFileSync(out, JSON.stringify(brief, null, 1));
  console.log('brief written:', out);
  console.log(JSON.stringify(brief.panels.map((p) => `${p.label}: ${p.latest.toFixed(2)} (${p.changePct >= 0 ? '+' : ''}${p.changePct.toFixed(2)}%)`), null, 1));
  return brief;
}

if (require.main === module) {
  buildBrief().catch((e) => {
    console.error(e);
    process.exit(1);
  });
}
