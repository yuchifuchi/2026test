// 公演カタログ(読み取り専用データ)の型定義
// Phase 1: アプリ同梱JSON / Phase 2以降: CDN配信JSONに差し替え(型は共通)

export interface CatalogArtist {
  id: string;
  name: string;
  kana?: string;
  genre: string;
  color: string; // 推し色の初期値(メンバーカラーではなくグループ代表色)
  icon?: string;
  officialUrl?: string;
  xAccount?: string;
}

export interface CatalogSlot {
  // 申込枠(先行/一般)。templateIdはBUILTIN_TEMPLATESに対応
  templateId: 'fc' | 'official' | 'pia' | 'lawson' | 'eplus' | 'general' | 'custom';
  name: string; // 表示名(例: FC先行(FRUITS ZIPPER FAMILY))
  applyStart?: string; // ISO8601
  applyEnd?: string;
  announceAt?: string;
  url?: string; // 申込/告知ページURL
}

export interface CatalogDateDetail {
  date: string; // 'YYYY-MM-DD'
  venue?: string; // ツアーで日程ごとに会場が異なる場合
  city?: string;
}

export interface CatalogEvent {
  id: string;
  artistIds: string[]; // 合同イベント・フェスは複数
  title: string;
  venue?: string; // 単一会場の場合
  city?: string;
  dates: string[]; // 'YYYY-MM-DD'(複数日程)
  datesDetail?: CatalogDateDetail[]; // 日程ごとの会場(ツアー用)
  price?: string; // 例: '指定席 ¥8,000(税込)'
  slots: CatalogSlot[];
  sourceUrl: string; // 出典(必須)
  verified: boolean; // 検品済みフラグ
  updatedAt: string;
}

export interface Catalog {
  version: number;
  generatedAt: string;
  artists: CatalogArtist[];
  events: CatalogEvent[];
}
