import { Catalog, CatalogArtist, CatalogEvent } from './types';

// Phase 1: アプリ同梱のカタログ(検品済みデータのみ)
// Phase 2: 起動時にCDNから取得し、失敗時は同梱/前回取得分へフォールバック
// eslint-disable-next-line @typescript-eslint/no-require-imports
const bundled = require('./data.json') as Catalog;

export function getCatalog(): Catalog {
  return bundled;
}

export function listCatalogArtists(): CatalogArtist[] {
  return bundled.artists;
}

export function getCatalogArtist(id: string): CatalogArtist | undefined {
  return bundled.artists.find((a) => a.id === id);
}

export function eventsForArtist(artistId: string): CatalogEvent[] {
  const today = new Date().toISOString().slice(0, 10);
  return bundled.events
    .filter((e) => e.artistId === artistId)
    .filter((e) => e.dates.some((d) => d >= today)) // 未来の日程が残っている公演のみ
    .sort((a, b) => (a.dates[0] ?? '').localeCompare(b.dates[0] ?? ''));
}
