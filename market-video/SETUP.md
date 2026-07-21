# セットアップ手順(ローカルPC)

## 0. 前提

- Node.js 20以上(https://nodejs.org からLTSをインストール)
- Git

## 1. クローンとインストール

```bash
git clone -b claude/app-concept-market-research-8pqclv <このリポジトリのURL>
cd 2026test/market-video
npm install
```

## 2. FRED APIキーの取得(無料・5分)

1. https://fredaccount.stlouisfed.org/login/secure/ を開き「Create New Account」でアカウント作成(メールアドレスのみ)
2. ログイン後、右上のメニューから **My Account → API Keys** を開く
3. **Request API Key** をクリック。用途欄は「Personal research / data visualization」程度でOK
4. 64文字のキーが即時発行される
5. このフォルダに `.env` ファイルを作成し、次の1行を記載:

```
FRED_API_KEY=ここに発行されたキーを貼り付け
```

6. 接続テスト:

```bash
npm run test-fred
# 「✅ FRED接続OK」と最新CPI・発表日が表示されれば成功
```

※ `.env` は `.gitignore` 済みでGitにはコミットされません。キーは他人に共有しないでください。

## 3. 毎朝の実行(1コマンド)

```bash
npm run morning
```

- データ取得(米指標5+セクターETF7)→ 集計 → 動画(out/short-日付.mp4)→ サムネ生成まで自動
- 初回はチャート用ブラウザのダウンロードが入るため数分かかることがあります
- 生成後、プレビューして問題なければYouTube Studioからアップロード

### 自動起動の登録(毎朝6:30)

- **Windows**: タスクスケジューラ → 基本タスクの作成 → 毎日6:30 → プログラム `cmd`、引数 `/c cd /d C:\path\to\market-video && npm run morning`
- **Mac**: `crontab -e` で `30 6 * * 1-5 cd /path/to/market-video && npm run morning >> morning.log 2>&1`
- 平日のみ(月〜金)推奨。米国休場日は前日と同値になるため、その日は投稿スキップでOK

## 4. トラブルシューティング

- `データ取得が不十分です` → ネットワーク確認。yfinance→Stooqの順で自動フォールバックするため、両方失敗が続く場合は時間をおいて再実行
- レンダリング失敗(ブラウザ起動系)→ `npx remotion browser ensure` を一度実行
- 動画の数値がおかしい → data/ フォルダのキャッシュJSONを削除して再実行

## 5. 次のステップ(音声)

ElevenLabs(https://elevenlabs.io)でアカウント作成 → Voices → Instant Voice Clone で自分の声を10分以上録音・登録 → APIキーを `.env` に `ELEVENLABS_API_KEY=...` として追記。音声層の実装が入り次第、`npm run morning` に読み上げが自動追加されます。
