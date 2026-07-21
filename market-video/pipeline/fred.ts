/**
 * FRED API接続層(V1.5長尺のCPI検証で使用)
 * APIキーは .env の FRED_API_KEY(https://fredaccount.stlouisfed.org で無料取得)
 */
import * as fs from 'fs';
import * as path from 'path';

function apiKey(): string {
  const envPath = path.join(__dirname, '..', '.env');
  if (fs.existsSync(envPath)) {
    const m = fs.readFileSync(envPath, 'utf8').match(/^FRED_API_KEY=(.+)$/m);
    if (m) return m[1].trim();
  }
  const k = process.env.FRED_API_KEY;
  if (!k) throw new Error('FRED_API_KEY が未設定です。.env に FRED_API_KEY=あなたのキー を記載してください');
  return k;
}

const BASE = 'https://api.stlouisfed.org/fred';

export async function fredSeries(seriesId: string, start = '2010-01-01') {
  const url = `${BASE}/series/observations?series_id=${seriesId}&observation_start=${start}&api_key=${apiKey()}&file_type=json`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`FRED ${seriesId}: HTTP ${res.status}`);
  const j = (await res.json()) as any;
  return (j.observations as { date: string; value: string }[])
    .filter((o) => o.value !== '.')
    .map((o) => ({ date: o.date, value: parseFloat(o.value) }));
}

// CPI発表日一覧(release_id=10 が Consumer Price Index)
export async function cpiReleaseDates(start = '2020-01-01') {
  const url = `${BASE}/release/dates?release_id=10&realtime_start=${start}&api_key=${apiKey()}&file_type=json&sort_order=desc&limit=60`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`FRED release dates: HTTP ${res.status}`);
  const j = (await res.json()) as any;
  return (j.release_dates as { date: string }[]).map((d) => d.date);
}

// 接続テスト: npm run test-fred
if (require.main === module) {
  (async () => {
    const cpi = await fredSeries('CPIAUCSL', '2025-01-01');
    const dates = await cpiReleaseDates('2025-01-01');
    console.log('✅ FRED接続OK');
    console.log('最新CPI:', cpi[cpi.length - 1]);
    console.log('直近のCPI発表日:', dates.slice(0, 3).join(', '));
  })().catch((e) => {
    console.error('❌ FRED接続失敗:', e.message);
    process.exit(1);
  });
}
