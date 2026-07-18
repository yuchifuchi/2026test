import { drizzle } from 'drizzle-orm/expo-sqlite';
import * as SQLite from 'expo-sqlite';

export const sqlite = SQLite.openDatabaseSync('oshinote.db');
export const db = drizzle(sqlite);

// PRAGMA user_version によるシンプルなバージョン管理マイグレーション
const MIGRATIONS: string[] = [
  // v1: 初期スキーマ
  `
  CREATE TABLE IF NOT EXISTS oshi (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, genre TEXT,
    color1 TEXT NOT NULL, color2 TEXT, icon TEXT,
    sort_order INTEGER NOT NULL DEFAULT 0, created_at INTEGER NOT NULL
  );
  CREATE TABLE IF NOT EXISTS meigi (
    id TEXT PRIMARY KEY, label TEXT NOT NULL, memo TEXT,
    sort_order INTEGER NOT NULL DEFAULT 0
  );
  CREATE TABLE IF NOT EXISTS fc_membership (
    id TEXT PRIMARY KEY, oshi_id TEXT NOT NULL, name TEXT NOT NULL,
    renewal_date INTEGER, annual_fee REAL, notify_enabled INTEGER NOT NULL DEFAULT 1
  );
  CREATE TABLE IF NOT EXISTS event (
    id TEXT PRIMARY KEY, oshi_id TEXT NOT NULL, title TEXT NOT NULL,
    type TEXT NOT NULL DEFAULT 'live', venue TEXT, city TEXT, memo TEXT
  );
  CREATE TABLE IF NOT EXISTS performance (
    id TEXT PRIMARY KEY, event_id TEXT NOT NULL, date INTEGER NOT NULL,
    open_time TEXT, start_time TEXT
  );
  CREATE TABLE IF NOT EXISTS application (
    id TEXT PRIMARY KEY, event_id TEXT NOT NULL,
    template_id TEXT NOT NULL, template_name TEXT NOT NULL,
    meigi_id TEXT, quantity INTEGER NOT NULL DEFAULT 1,
    status TEXT NOT NULL DEFAULT 'planned',
    close_at INTEGER, announce_at INTEGER, payment_deadline INTEGER,
    seat_info TEXT, total_amount REAL, url TEXT, memo TEXT,
    created_at INTEGER NOT NULL
  );
  CREATE TABLE IF NOT EXISTS application_choice (
    id TEXT PRIMARY KEY, application_id TEXT NOT NULL,
    performance_id TEXT NOT NULL, priority INTEGER NOT NULL DEFAULT 1
  );
  CREATE TABLE IF NOT EXISTS expense (
    id TEXT PRIMARY KEY, oshi_id TEXT, event_id TEXT, application_id TEXT,
    category TEXT NOT NULL, amount REAL NOT NULL, date INTEGER NOT NULL, memo TEXT
  );
  CREATE TABLE IF NOT EXISTS budget (
    id TEXT PRIMARY KEY, month TEXT NOT NULL, oshi_id TEXT, limit_amount REAL NOT NULL
  );
  CREATE INDEX IF NOT EXISTS idx_application_event ON application(event_id);
  CREATE INDEX IF NOT EXISTS idx_performance_event ON performance(event_id);
  CREATE INDEX IF NOT EXISTS idx_expense_date ON expense(date);
  `,
  // v2: 公演カタログとの紐付け
  `ALTER TABLE oshi ADD COLUMN catalog_artist_id TEXT;`,
];

export function migrate(): void {
  const row = sqlite.getFirstSync<{ user_version: number }>('PRAGMA user_version');
  const current = row?.user_version ?? 0;
  for (let v = current; v < MIGRATIONS.length; v++) {
    sqlite.execSync(MIGRATIONS[v]);
    sqlite.execSync(`PRAGMA user_version = ${v + 1}`);
  }
}

export function newId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}
