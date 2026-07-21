/**
 * 毎朝1コマンド: データ取得→集計→動画+サムネ生成
 * 使い方: npm run morning
 */
import { execSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { buildBrief } from './brief';
import { generateNarration } from './narration';

(async () => {
  const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const outDir = path.join(__dirname, '..', 'out');
  fs.mkdirSync(outDir, { recursive: true });

  console.log('== 1/3 データ取得・集計 ==');
  const brief = await buildBrief();
  console.log('見出し:', brief.headline);

  console.log('== 1.5/3 ナレーション生成 ==');
  await generateNarration();

  console.log('== 2/3 動画レンダリング ==');
  const mp4 = path.join(outDir, `short-${stamp}.mp4`);
  execSync(`npx remotion render src/index.ts Short "${mp4}"`, { stdio: 'inherit', cwd: path.join(__dirname, '..') });

  console.log('== 2.5/3 編集用パッケージ出力 ==');
  const visual = path.join(outDir, `short-${stamp}-visual.mp4`);
  execSync(`npx remotion render src/index.ts Short "${visual}" --muted`, { stdio: 'inherit', cwd: path.join(__dirname, '..') });
  for (const f of ['narration.mp3', 'narration.wav', 'narration.txt']) {
    const src = path.join(__dirname, '..', 'public', f);
    if (fs.existsSync(src)) fs.copyFileSync(src, path.join(outDir, `short-${stamp}-${f}`));
  }

  console.log('== 3/3 サムネイル生成 ==');
  const png = path.join(outDir, `thumb-${stamp}.png`);
  execSync(`npx remotion still src/index.ts Short "${png}" --frame=150`, { stdio: 'inherit', cwd: path.join(__dirname, '..') });

  console.log('\n✅ 完了');
  console.log('動画:', mp4);
  console.log('サムネ:', png);
  console.log('→ プレビューして問題なければ YouTube Studio からアップロードしてください');
})().catch((e) => {
  console.error('❌ 失敗:', e.message ?? e);
  process.exit(1);
});
