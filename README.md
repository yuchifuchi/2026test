# 2026test — 個人開発アプリプロジェクト

役に立つ×マネタイズできるアプリをバイブコーディングで開発するプロジェクト。
**方向性はエンタメ(ゲーム)に転換済み**(野球関連・RPG系、アイテム課金)。

## 進行状況

- [x] フェーズ1: 市場調査・分析(汎用アプリ向け)
- [x] フェーズ1.5: 候補A(SaaS)の競合調査 → その後エンタメへ方向転換
- [x] フェーズ1.6: ゲーム向け市場・法規制・競合調査
- [ ] フェーズ2: 要件定義
- [ ] フェーズ3: 実装(MVP)
- [ ] フェーズ4: リリース・改善

## ドキュメント

| ドキュメント | 内容 |
|---|---|
| [docs/game-market-research.md](docs/game-market-research.md) | **ゲーム市場調査**(成功事例 / アイテム課金の法規制 / 野球ゲーム競合 / 技術選定) |
| [docs/game-candidates.md](docs/game-candidates.md) | **ゲーム候補5案の評価と推奨案(高校野球監督シム)** |
| [docs/market-research.md](docs/market-research.md) | (アーカイブ)汎用アプリ向け市場調査 |
| [docs/app-candidates.md](docs/app-candidates.md) | (アーカイブ)SaaS系候補5案の評価 |
| [docs/competitor-analysis.md](docs/competitor-analysis.md) | (アーカイブ)候補A(フリーランスSaaS)の競合分析 |

## 現在の推奨案

**G1: 高校野球監督シミュレーション**(架空チーム・架空選手、RPG的育成・3年周回)

- 需要の空白: 公式スマホ版『栄冠クロス』が2026年3月にサ終し、スマホの監督シム本命が不在
- 先行実証: 同ジャンルの個人開発『私を甲子園に連れてって』が累計1,290万円・最高月商257万円
- 課金設計: 通貨なし・ガチャなしの直接販売(殿堂チケット・広告除去・エディット拡張)+リワード広告 — 資金決済法・景表法リスクを構造的に回避
- 技術: React/Next.js(Web/PWA)→ Capacitorでモバイル化+ストアIAP

詳細は [docs/game-candidates.md](docs/game-candidates.md) を参照。
