import React from 'react';
import { AbsoluteFill, Audio, interpolate, spring, staticFile, useCurrentFrame, useVideoConfig } from 'remotion';
import brief from '../data/brief.json';

const UP = '#3DDC84';
const DOWN = '#FF5A6E';
const BG = '#0E1524';
const CARD = '#1A2438';
const INK = '#F2F5FA';
const MUTED = '#8B98AF';

function Spark({ data, color, w, h }: { data: number[]; color: string; w: number; h: number }) {
  const min = Math.min(...data);
  const max = Math.max(...data);
  const pts = data
    .map((v, i) => `${(i / (data.length - 1)) * w},${h - ((v - min) / (max - min || 1)) * h}`)
    .join(' ');
  return (
    <svg width={w} height={h}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth={6} strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

export const Short: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const isSample = (brief as any).dataMode === 'sample';

  return (
    <AbsoluteFill style={{ backgroundColor: BG, fontFamily: "'Noto Sans JP','Noto Sans CJK JP','Hiragino Sans','Yu Gothic Medium','Meiryo',sans-serif", padding: 60 }}>
      {(brief as any).narrationFile ? <Audio src={staticFile((brief as any).narrationFile)} /> : null}
      {/* ヘッダー */}
      <div style={{ opacity: interpolate(frame, [0, 15], [0, 1]) }}>
        <div style={{ color: MUTED, fontSize: 40, fontWeight: 700, letterSpacing: 2 }}>
          {brief.dateLabel} 今朝の海外→日本株
        </div>
        <div style={{ color: INK, fontSize: 58, fontWeight: 800, marginTop: 18, lineHeight: 1.25 }}>
          {brief.headline}
        </div>
      </div>

      {/* 4パネル */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16, marginTop: 34 }}>
        {brief.panels.map((p, i) => {
          const appear = spring({ frame: frame - 20 - i * 12, fps, config: { damping: 200 } });
          const positive = p.changePct >= 0;
          const color = positive ? UP : DOWN;
          const drawn = Math.max(2, Math.floor(interpolate(frame, [30 + i * 12, 75 + i * 12], [2, p.spark.length], { extrapolateRight: 'clamp' })));
          return (
            <div
              key={p.label}
              style={{
                backgroundColor: CARD,
                borderRadius: 28,
                padding: '20px 36px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                transform: `translateY(${(1 - appear) * 80}px)`,
                opacity: appear,
              }}
            >
              <div>
                <div style={{ color: MUTED, fontSize: 38, fontWeight: 700 }}>{p.label}</div>
                <div style={{ color: INK, fontSize: 58, fontWeight: 800, marginTop: 6, fontVariantNumeric: 'tabular-nums' }}>
                  {p.latest.toLocaleString('ja-JP', { maximumFractionDigits: 2 })}
                </div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 30 }}>
                <Spark data={p.spark.slice(0, drawn)} color={color} w={260} h={110} />
                <div
                  style={{
                    backgroundColor: color,
                    color: '#08101E',
                    fontSize: 44,
                    fontWeight: 800,
                    borderRadius: 18,
                    padding: '14px 22px',
                    minWidth: 190,
                    textAlign: 'center',
                    fontVariantNumeric: 'tabular-nums',
                  }}
                >
                  {positive ? '+' : ''}
                  {p.changePct.toFixed(2)}%
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {/* セクター注目カード */}
      {(brief as any).sectors ? (
        <div
          style={{
            backgroundColor: '#22304C', borderRadius: 24, padding: '24px 36px', marginTop: 22,
            opacity: interpolate(frame, [95, 115], [0, 1]),
          }}
        >
          <div style={{ color: '#F5C518', fontSize: 32, fontWeight: 800 }}>昨夜の米セクター TOP3(事実)</div>
          <div style={{ display: 'flex', gap: 18, marginTop: 14 }}>
            {(brief as any).sectors.map((sct: any, j: number) => (
              <div key={sct.label} style={{ flex: 1, backgroundColor: CARD, borderRadius: 16, padding: '14px 16px', textAlign: 'center' }}>
                <div style={{ color: MUTED, fontSize: 26, fontWeight: 700 }}>{j + 1}位 {sct.label}</div>
                <div style={{ color: sct.changePct >= 0 ? UP : DOWN, fontSize: 36, fontWeight: 800, fontVariantNumeric: 'tabular-nums' }}>
                  {sct.changePct >= 0 ? '+' : ''}{sct.changePct.toFixed(2)}%
                </div>
              </div>
            ))}
          </div>
          <div style={{ color: INK, fontSize: 30, fontWeight: 700, marginTop: 14, lineHeight: 1.4 }}>
            {(brief as any).sectorNote}
          </div>
        </div>
      ) : null}

      {/* フッター: 出典+免責 */}
      <div style={{ position: 'absolute', bottom: 36, left: 60, right: 60, opacity: interpolate(frame, [80, 100], [0, 1]) }}>
        <div style={{ color: MUTED, fontSize: 30 }}>
          出典: {brief.sources.join(' / ')} ・ {brief.generatedAt.slice(0, 10)}取得
        </div>
        <div style={{ color: MUTED, fontSize: 28, marginTop: 8 }}>
          本動画は情報提供であり投資助言ではありません。投資判断はご自身で。
        </div>
      </div>

      {isSample ? (
        <div
          style={{
            position: 'absolute',
            top: 40,
            right: 40,
            backgroundColor: '#F5C518',
            color: '#08101E',
            fontSize: 34,
            fontWeight: 800,
            borderRadius: 14,
            padding: '10px 20px',
            transform: 'rotate(4deg)',
          }}
        >
          サンプルデータ
        </div>
      ) : null}
    </AbsoluteFill>
  );
};
