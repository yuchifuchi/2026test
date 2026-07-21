/**
 * ナレーション層: brief.json → 読み上げ台本(数値は機械挿入) → TTS音声
 * バックエンド優先順: ElevenLabs(本人声クローン) > Edge TTS(自然な日本語) > Open JTalk(ローカル予備)
 */
import { execSync, spawnSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';

const ROOT = path.join(__dirname, '..');
const PUB = path.join(ROOT, 'public');

function env(key: string): string | undefined {
  const envPath = path.join(ROOT, '.env');
  if (fs.existsSync(envPath)) {
    const m = fs.readFileSync(envPath, 'utf8').match(new RegExp(`^${key}=(.+)$`, 'm'));
    if (m && m[1].trim()) return m[1].trim();
  }
  return process.env[key];
}

// 台本テンプレ: 数値はすべてbriefから機械挿入(LLM不使用)
export function buildNarrationText(brief: any): string {
  const p = (label: string) => brief.panels.find((x: any) => x.label.startsWith(label));
  const pct = (x: any) => `${x.changePct >= 0 ? 'プラス' : 'マイナス'}${Math.abs(x.changePct).toFixed(2)}パーセント`;
  const sp = p('S&P500');
  const nq = p('NASDAQ');
  const fx = p('ドル円');
  const fut = p('日経平均先物');
  const parts: string[] = [];
  parts.push(`${brief.dateLabel}、今朝の海外マーケットです。`);
  if (sp) parts.push(`エスアンドピー500は${pct(sp)}。`);
  if (nq) parts.push(`ナスダックは${pct(nq)}でした。`);
  if (fx) parts.push(`ドル円は${fx.latest.toFixed(2)}円。`);
  if (fut) parts.push(`日経平均先物は${Math.round(fut.latest).toLocaleString('ja-JP')}円です。`);
  if (brief.sectorNote) parts.push(`${brief.sectorNote}。`);
  parts.push('数字の出典は概要欄に。本動画は投資助言ではありません。');
  return parts.join('');
}

function ttsOpenJtalk(text: string, outWav: string): boolean {
  const dic = '/var/lib/mecab/dic/open-jtalk/naist-jdic';
  const voice = '/usr/share/hts-voice/nitech-jp-atr503-m001/nitech_jp_atr503_m001.htsvoice';
  if (!fs.existsSync(dic) || !fs.existsSync(voice)) return false;
  const txt = path.join(PUB, 'narration.txt');
  fs.writeFileSync(txt, text);
  const r = spawnSync('open_jtalk', ['-x', dic, '-m', voice, '-r', '1.05', '-ow', outWav, txt]);
  return r.status === 0 && fs.existsSync(outWav);
}

function ttsEdge(text: string, outMedia: string): boolean {
  try {
    execSync(
      `python3 -m edge_tts --voice ja-JP-NanamiNeural --rate=+8% --text ${JSON.stringify(text)} --write-media ${JSON.stringify(outMedia)}`,
      { stdio: 'pipe', timeout: 120000 },
    );
    return fs.existsSync(outMedia) && fs.statSync(outMedia).size > 1000;
  } catch {
    return false;
  }
}

async function ttsElevenLabs(text: string, outMedia: string): Promise<boolean> {
  const key = env('ELEVENLABS_API_KEY');
  const voiceId = env('ELEVENLABS_VOICE_ID');
  if (!key || !voiceId) return false;
  const res = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${voiceId}?output_format=mp3_44100_128`, {
    method: 'POST',
    headers: { 'xi-api-key': key, 'Content-Type': 'application/json' },
    body: JSON.stringify({ text, model_id: 'eleven_multilingual_v2' }),
  });
  if (!res.ok) return false;
  fs.writeFileSync(outMedia, Buffer.from(await res.arrayBuffer()));
  return true;
}

function wavDurationSec(file: string): number {
  const b = fs.readFileSync(file);
  const byteRate = b.readUInt32LE(28);
  const dataSize = b.length - 44;
  return dataSize / byteRate;
}

export async function generateNarration(): Promise<void> {
  fs.mkdirSync(PUB, { recursive: true });
  const briefPath = path.join(ROOT, 'data', 'brief.json');
  const brief = JSON.parse(fs.readFileSync(briefPath, 'utf8'));
  const text = buildNarrationText(brief);
  console.log('[narration]', text);

  const mp3 = path.join(PUB, 'narration.mp3');
  const wav = path.join(PUB, 'narration.wav');
  for (const f of [mp3, wav]) if (fs.existsSync(f)) fs.unlinkSync(f);

  let file = '';
  let backend = '';
  if (await ttsElevenLabs(text, mp3)) {
    file = 'narration.mp3'; backend = 'elevenlabs(本人声)';
  } else if (ttsEdge(text, mp3)) {
    file = 'narration.mp3'; backend = 'edge-tts(サンプル声)';
  } else if (ttsOpenJtalk(text, wav)) {
    file = 'narration.wav'; backend = 'open_jtalk(予備サンプル声)';
  }

  if (!file) {
    console.warn('[narration] TTS不可: 音声なしで続行');
    brief.narrationFile = null;
  } else {
    const full = path.join(PUB, file);
    const durationSec = file.endsWith('.wav')
      ? wavDurationSec(full)
      : Math.max(14, fs.statSync(full).size / 16000); // mp3は概算(128kbps)
    brief.narrationFile = file;
    brief.narrationDuration = Math.round(durationSec * 10) / 10;
    console.log(`[narration] ${backend} → ${file} (${brief.narrationDuration}s)`);
  }
  fs.writeFileSync(briefPath, JSON.stringify(brief, null, 1));
}

if (require.main === module) {
  generateNarration().catch((e) => {
    console.error(e);
    process.exit(1);
  });
}
