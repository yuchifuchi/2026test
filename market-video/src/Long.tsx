import React from 'react';
import { AbsoluteFill, Audio, interpolate, spring, staticFile, useCurrentFrame, useVideoConfig } from 'remotion';
import data from '../data/longform.json';

const UP = '#3DDC84';
const DOWN = '#FF5A6E';
const BG = '#0E1524';
const CARD = '#1A2438';
const INK = '#F2F5FA';
const MUTED = '#8B98AF';
const GOLD = '#F5C518';

const events = (data as any).events as { date: string; cpiYoY: number; surprise: string; nikkeiNextDayPct: number }[];
const downCount = events.filter((e) => e.nikkeiNextDayPct < 0).length;
const avgPct = events.reduce((s, e) => s + e.nikkeiNextDayPct, 0) / events.length;
const above = events.filter((e) => e.surprise === '予想上振れ');
const downOnAbove = above.filter((e) => e.nikkeiNextDayPct < 0).length;

// セクション開始秒(ナレーションと同期)
const SEC = { title: 0, method: 6, table: 14, stats: 50, conclusion: 62 };

function Section({ from, children }: { from: number; children: React.ReactNode }) {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const f = frame - from * fps;
  if (f < 0) return null;
  return <div style={{ opacity: interpolate(f, [0, 12], [0, 1]) }}>{children}</div>;
}

export const Long: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const isSample = (data as any).dataMode === 'sample';

  return (
    <AbsoluteFill style={{ backgroundColor: BG, fontFamily: 'sans-serif', padding: '50px 90px', color: INK }}>
      {(data as any).narrationFile ? <Audio src={staticFile((data as any).narrationFile)} /> : null}

      {/* タイトル(常時ヘッダー化) */}
      <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between' }}>
        <div>
          <div style={{ color: GOLD, fontSize: 30, fontWeight: 800, letterSpacing: 3 }}>検証シリーズ #01</div>
          <div style={{ fontSize: frame < SEC.method * fps ? 72 : 44, fontWeight: 800, marginTop: 10, transition: 'none', lineHeight: 1.3 }}>
            {(data as any).title}
          </div>
        </div>
      </div>

      {/* 検証方法 */}
      {frame >= SEC.method * fps && frame < SEC.stats * fps ? (
        <Section from={SEC.method}>
          <div style={{ color: MUTED, fontSize: 28, marginTop: 18 }}>
            検証方法: 米CPI発表日の「翌営業日」の日経平均の騰落率を、直近10回分すべて並べる(出典は概要欄)
          </div>
        </Section>
      ) : null}

      {/* 検証テーブル */}
      {frame >= SEC.table * fps && frame < SEC.stats * fps ? (
        <div style={{ marginTop: 26 }}>
          <div style={{ display: 'flex', color: MUTED, fontSize: 26, fontWeight: 700, padding: '0 24px', marginBottom: 10 }}>
            <div style={{ width: 260 }}>発表日</div>
            <div style={{ width: 260 }}>CPI(前年比)</div>
            <div style={{ width: 300 }}>結果</div>
            <div style={{ flex: 1, textAlign: 'right' }}>翌日の日経平均</div>
          </div>
          {events.map((e, i) => {
            const appear = spring({ frame: frame - (SEC.table + 1 + i * 4) * fps, fps, config: { damping: 200 } });
            const neg = e.nikkeiNextDayPct < 0;
            return (
              <div
                key={e.date}
                style={{
                  display: 'flex', alignItems: 'center', backgroundColor: CARD, borderRadius: 14,
                  padding: '14px 24px', marginBottom: 8, fontSize: 30, fontVariantNumeric: 'tabular-nums',
                  opacity: appear, transform: `translateX(${(1 - appear) * 60}px)`,
                }}
              >
                <div style={{ width: 260, color: MUTED }}>{e.date}</div>
                <div style={{ width: 260 }}>+{e.cpiYoY.toFixed(1)}%</div>
                <div style={{ width: 300, color: e.surprise === '予想上振れ' ? DOWN : e.surprise === '予想下振れ' ? UP : MUTED }}>{e.surprise}</div>
                <div style={{ flex: 1, textAlign: 'right', fontWeight: 800, color: neg ? DOWN : UP }}>
                  {neg ? '' : '+'}{e.nikkeiNextDayPct.toFixed(2)}%
                </div>
              </div>
            );
          })}
        </div>
      ) : null}

      {/* 集計結果 */}
      {frame >= SEC.stats * fps ? (
        <Section from={SEC.stats}>
          <div style={{ display: 'flex', gap: 30, marginTop: 40 }}>
            {[
              { label: '翌日に下落した回数', value: `${downCount} / ${events.length}回`, color: INK },
              { label: '平均騰落率', value: `${avgPct >= 0 ? '+' : ''}${avgPct.toFixed(2)}%`, color: avgPct >= 0 ? UP : DOWN },
              { label: '「予想上振れ」の日だけ見ると', value: `${downOnAbove} / ${above.length}回 下落`, color: DOWN },
            ].map((s) => (
              <div key={s.label} style={{ flex: 1, backgroundColor: CARD, borderRadius: 20, padding: '34px 30px', textAlign: 'center' }}>
                <div style={{ color: MUTED, fontSize: 28, fontWeight: 700 }}>{s.label}</div>
                <div style={{ fontSize: 64, fontWeight: 800, marginTop: 14, color: s.color, fontVariantNumeric: 'tabular-nums' }}>{s.value}</div>
              </div>
            ))}
          </div>
          {frame >= SEC.conclusion * fps ? (
            <div style={{ backgroundColor: '#22304C', borderRadius: 20, padding: '28px 36px', marginTop: 26, fontSize: 31, lineHeight: 1.55 }}>
              {(data as any).insight?.text}
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, marginTop: 16 }}>
                {((data as any).insight?.evidence ?? []).map((ev: any) => (
                  <div key={ev.label} style={{ backgroundColor: CARD, borderRadius: 10, padding: '8px 14px', fontSize: 21, color: MUTED }}>
                    📄 {ev.label}
                  </div>
                ))}
              </div>
            </div>
          ) : null}
        </Section>
      ) : null}

      {/* フッター */}
      <div style={{ position: 'absolute', bottom: 34, left: 90, right: 90, display: 'flex', justifyContent: 'space-between', color: MUTED, fontSize: 24 }}>
        <div>出典: {(data as any).sources.join(' / ')}</div>
        <div>本動画は情報提供であり投資助言ではありません</div>
      </div>

      {isSample ? (
        <div style={{ position: 'absolute', top: 34, right: 90, backgroundColor: GOLD, color: '#08101E', fontSize: 30, fontWeight: 800, borderRadius: 12, padding: '8px 18px', transform: 'rotate(3deg)' }}>
          サンプルデータ
        </div>
      ) : null}
    </AbsoluteFill>
  );
};
