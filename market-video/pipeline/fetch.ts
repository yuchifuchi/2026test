/**
 * 取得層: yfinance主系(Yahoo chart API直叩き) + Stooqフォールバック
 * 設計原則: 差分取得・直列実行・2〜5秒ランダム待機・失敗時バックオフ→ソース切替
 */
import * as fs from 'fs';
import * as path from 'path';

const DATA_DIR = path.join(__dirname, '..', 'data');
const UA = 'oshi-market-video/0.1 (personal research; contact: see repo)';

export interface Quote {
  date: string; // YYYY-MM-DD
  close: number;
}
export interface SeriesData {
  symbol: string;
  label: string;
  source: 'yahoo' | 'stooq';
  fetchedAt: string;
  quotes: Quote[];
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
const jitter = () => 2000 + Math.random() * 3000;

async function fetchYahoo(symbol: string): Promise<Quote[]> {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?range=6mo&interval=1d`;
  const res = await fetch(url, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error(`yahoo ${symbol}: HTTP ${res.status}`);
  const j = (await res.json()) as any;
  const r = j.chart?.result?.[0];
  if (!r) throw new Error(`yahoo ${symbol}: empty result`);
  const ts: number[] = r.timestamp ?? [];
  const closes: (number | null)[] = r.indicators?.quote?.[0]?.close ?? [];
  return ts
    .map((t, i) => ({
      date: new Date(t * 1000).toISOString().slice(0, 10),
      close: closes[i] as number,
    }))
    .filter((q) => q.close != null && Number.isFinite(q.close));
}

async function fetchStooq(stooqSymbol: string): Promise<Quote[]> {
  const url = `https://stooq.com/q/d/l/?s=${encodeURIComponent(stooqSymbol)}&i=d`;
  const res = await fetch(url, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error(`stooq ${stooqSymbol}: HTTP ${res.status}`);
  const csv = await res.text();
  const lines = csv.trim().split('\n').slice(1);
  return lines
    .map((l) => {
      const [date, , , , close] = l.split(',');
      return { date, close: parseFloat(close) };
    })
    .filter((q) => Number.isFinite(q.close))
    .slice(-130); // 約6ヶ月分
}

export interface SymbolDef {
  symbol: string; // Yahoo
  stooq: string; // Stooqフォールバック
  label: string;
}

export async function fetchAll(symbols: SymbolDef[]): Promise<SeriesData[]> {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const out: SeriesData[] = [];
  for (const def of symbols) {
    const cacheFile = path.join(DATA_DIR, `${def.symbol.replace(/[^\w]/g, '_')}.json`);
    // 差分判定: 当日すでに取得済みならキャッシュを使う(リクエストを出さない)
    if (fs.existsSync(cacheFile)) {
      const cached = JSON.parse(fs.readFileSync(cacheFile, 'utf8')) as SeriesData;
      if (cached.fetchedAt.slice(0, 10) === new Date().toISOString().slice(0, 10)) {
        out.push(cached);
        continue;
      }
    }
    let series: SeriesData | null = null;
    // yfinance主系(リトライ2回・バックオフ) → Stooqフォールバック
    for (let attempt = 0; attempt < 2 && !series; attempt++) {
      try {
        const quotes = await fetchYahoo(def.symbol);
        series = { symbol: def.symbol, label: def.label, source: 'yahoo', fetchedAt: new Date().toISOString(), quotes };
      } catch (e) {
        console.warn(`[fetch] yahoo failed (${def.symbol}, attempt ${attempt + 1}):`, (e as Error).message);
        await sleep(3000 * (attempt + 1));
      }
    }
    if (!series) {
      try {
        const quotes = await fetchStooq(def.stooq);
        series = { symbol: def.symbol, label: def.label, source: 'stooq', fetchedAt: new Date().toISOString(), quotes };
        console.warn(`[fetch] fell back to stooq for ${def.symbol}`);
      } catch (e) {
        console.error(`[fetch] stooq also failed (${def.stooq}):`, (e as Error).message);
      }
    }
    if (series) {
      fs.writeFileSync(cacheFile, JSON.stringify(series, null, 1));
      out.push(series);
    }
    await sleep(jitter()); // 直列+ランダム待機
  }
  return out;
}
