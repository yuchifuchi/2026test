# 2026test — 個人開発アプリプロジェクト

役に立つ×マネタイズできるアプリをバイブコーディングで開発するプロジェクト。
**方向性はエンタメ(ゲーム)に転換済み**(野球関連・RPG系、アイテム課金)。

## 進行状況

- [x] フェーズ1: 市場調査・分析(汎用アプリ向け)
- [x] フェーズ1.5: 候補A(SaaS)の競合調査 → その後エンタメへ方向転換
- [x] フェーズ1.6: ゲーム向け市場・法規制・競合調査
- [x] フェーズ1.7: 日常系・若年層向け・SNS系の追加調査(非ゲーム路線の比較)
- [x] フェーズ2: 方向性決定(L2: 推し活管理)→ 要件定義([docs/requirements.md](docs/requirements.md))
- [ ] フェーズ3: 実装(MVP)
- [ ] フェーズ4: リリース・改善

## ドキュメント

| ドキュメント | 内容 |
|---|---|
| [docs/game-market-research.md](docs/game-market-research.md) | ゲーム市場調査(成功事例 / アイテム課金の法規制 / 野球ゲーム競合 / 技術選定) |
| [docs/game-candidates.md](docs/game-candidates.md) | ゲーム候補5案の評価と推奨案(G1: 高校野球監督シム) |
| [docs/lifestyle-app-research.md](docs/lifestyle-app-research.md) | 日常系・若年層向け・SNS系の調査と候補6案(L1〜L6)、全候補比較表 |
| [docs/oshikatsu-deep-dive.md](docs/oshikatsu-deep-dive.md) | L2深掘り: 推し活管理アプリの競合分析・ユーザー実態・コンセプト案 |
| [docs/requirements.md](docs/requirements.md) | **要件定義書: 当落ノート(仮称)— 機能要件・通知仕様・データモデル・画面構成・技術選定・課金設計** |
| [docs/market-research.md](docs/market-research.md) | (アーカイブ)汎用アプリ向け市場調査 |
| [docs/app-candidates.md](docs/app-candidates.md) | (アーカイブ)SaaS系候補5案の評価 |
| [docs/competitor-analysis.md](docs/competitor-analysis.md) | (アーカイブ)候補A(フリーランスSaaS)の競合分析 |

## 方向性の候補(検討中)

| 路線 | 最有力候補 | スコア | 収益の型 |
|---|---|---|---|
| ゲーム | G1: 高校野球監督シム | 26 | 広告+アイテム直接課金(先行事例: 月商257万円) |
| 日常系 | L1: 習慣×ペット育成×ウィジェット | 24 | 低額サブスク(Finch $30M ARR / HabitKit 年$602K) |
| 若年層 | **L2: 推し活の予算・当落・遠征管理(深掘り済み・最有力)** | 24 | サブスク+アフィリエイト(市場3.5兆円・一気通貫アプリ不在) |
| バズ型 | L3: 診断・分析シェア(TikTok自走) | 24 | バズ→サブスク(IsTalk: 月商250万円) |

詳細比較は [docs/lifestyle-app-research.md](docs/lifestyle-app-research.md) の比較表を参照。
