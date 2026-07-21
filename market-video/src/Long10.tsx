import React from 'react';
import { AbsoluteFill, Audio, interpolate, spring, staticFile, useCurrentFrame, useVideoConfig } from 'remotion';
import data from '../data/long10.json';

const UP = '#3DDC84';
const BG = '#0E1524';
const CARD = '#1A2438';
const INK = '#F2F5FA';
const MUTED = '#8B98AF';
const GOLD = '#F5C518';
const FONT = "'Noto Sans JP','Noto Sans CJK JP','Hiragino Sans','Yu Gothic Medium','Meiryo',sans-serif";

const d = data as any;
const chapters = d.chapters as any[];
const starts: number[] = d.chapterStarts ?? chapters.map((_, i) => i * 110);

export const Long10: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const t = frame / fps;
  let ci = 0;
  for (let i = 0; i < starts.length; i++) if (t >= starts[i]) ci = i;
  const ch = chapters[ci];
  const chFrame = frame - starts[ci] * fps;
  const fadeIn = interpolate(chFrame, [0, 15], [0, 1]);

  return (
    <AbsoluteFill style={{ backgroundColor: BG, fontFamily: FONT, padding: '48px 90px', color: INK }}>
      {d.narrationFile ? <Audio src={staticFile(d.narrationFile)} /> : null}

      <div style={{ color: GOLD, fontSize: 28, fontWeight: 800, letterSpacing: 3 }}>
        検証シリーズ #02 ・ {ci + 1} / {chapters.length}
      </div>
      <div style={{ fontSize: 46, fontWeight: 800, marginTop: 8, lineHeight: 1.3, opacity: fadeIn }}>{ch.title}</div>

      <div style={{ marginTop: 34, opacity: fadeIn }}>
        {ch.rows ? (
          <div>
            {ch.rows.map((r: any, i: number) => {
              const appear = spring({ frame: chFrame - (20 + i * 25), fps, config: { damping: 200 } });
              return (
                <div key={i} style={{ display: 'flex', alignItems: 'center', backgroundColor: CARD, borderRadius: 16, padding: '26px 32px', marginBottom: 14, opacity: appear, transform: `translateX(${(1 - appear) * 60}px)` }}>
                  <div style={{ flex: 1.4, fontSize: 31, color: INK, lineHeight: 1.4 }}>{r.label}</div>
                  <div style={{ flex: 1, fontSize: 34, fontWeight: 800, color: GOLD, textAlign: 'center', fontVariantNumeric: 'tabular-nums' }}>{r.a}</div>
                  <div style={{ flex: 1, fontSize: 34, fontWeight: 800, color: UP, textAlign: 'center', fontVariantNumeric: 'tabular-nums' }}>{r.b}</div>
                </div>
              );
            })}
          </div>
        ) : null}

        {ch.stats ? (
          <div style={{ display: 'flex', gap: 28, marginTop: 6 }}>
            {ch.stats.map((s: any, i: number) => {
              const appear = spring({ frame: chFrame - (20 + i * 25), fps, config: { damping: 200 } });
              return (
                <div key={i} style={{ flex: 1, backgroundColor: CARD, borderRadius: 20, padding: '38px 30px', textAlign: 'center', opacity: appear, transform: `translateY(${(1 - appear) * 40}px)` }}>
                  <div style={{ color: MUTED, fontSize: 28, fontWeight: 700, lineHeight: 1.4 }}>{s.label}</div>
                  <div style={{ fontSize: 52, fontWeight: 800, marginTop: 16, color: INK, fontVariantNumeric: 'tabular-nums', lineHeight: 1.25 }}>{s.value}</div>
                </div>
              );
            })}
          </div>
        ) : null}

        {ch.note ? (
          <div style={{ backgroundColor: '#22304C', borderRadius: 18, padding: '24px 32px', marginTop: 24, fontSize: 30, lineHeight: 1.6, color: INK, opacity: interpolate(chFrame, [90, 110], [0, 1]) }}>
            {ch.note}
          </div>
        ) : null}

        {ci === chapters.length - 1 ? (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, marginTop: 20, opacity: interpolate(chFrame, [120, 140], [0, 1]) }}>
            {(d.insight?.evidence ?? []).map((ev: any) => (
              <div key={ev.label} style={{ backgroundColor: CARD, borderRadius: 10, padding: '10px 16px', fontSize: 23, color: MUTED }}>📄 {ev.label}</div>
            ))}
          </div>
        ) : null}
      </div>

      <div style={{ position: 'absolute', bottom: 30, left: 90, right: 90, display: 'flex', justifyContent: 'space-between', color: MUTED, fontSize: 23 }}>
        <div>出典: {d.sources.join(' / ')}</div>
        <div>本動画は情報提供であり投資助言ではありません。投資判断はご自身で</div>
      </div>
      <div style={{ position: 'absolute', top: 30, right: 90, backgroundColor: GOLD, color: '#08101E', fontSize: 28, fontWeight: 800, borderRadius: 12, padding: '8px 18px', transform: 'rotate(3deg)' }}>
        サンプルデータ(制作テスト)
      </div>
    </AbsoluteFill>
  );
};
