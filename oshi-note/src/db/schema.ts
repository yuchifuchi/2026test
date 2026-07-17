import { integer, real, sqliteTable, text } from 'drizzle-orm/sqlite-core';

// 日時は unixepoch ミリ秒(integer)で保存する

export const oshi = sqliteTable('oshi', {
  id: text('id').primaryKey(),
  name: text('name').notNull(),
  genre: text('genre'),
  color1: text('color1').notNull(),
  color2: text('color2'),
  icon: text('icon'), // 絵文字
  sortOrder: integer('sort_order').notNull().default(0),
  createdAt: integer('created_at').notNull(),
});

export const meigi = sqliteTable('meigi', {
  id: text('id').primaryKey(),
  label: text('label').notNull(),
  memo: text('memo'),
  sortOrder: integer('sort_order').notNull().default(0),
});

export const fcMembership = sqliteTable('fc_membership', {
  id: text('id').primaryKey(),
  oshiId: text('oshi_id').notNull(),
  name: text('name').notNull(),
  renewalDate: integer('renewal_date'),
  annualFee: real('annual_fee'),
  notifyEnabled: integer('notify_enabled').notNull().default(1),
});

export const event = sqliteTable('event', {
  id: text('id').primaryKey(),
  oshiId: text('oshi_id').notNull(),
  title: text('title').notNull(),
  type: text('type').notNull().default('live'), // live/stage/release/stream/other
  venue: text('venue'),
  city: text('city'),
  memo: text('memo'),
});

export const performance = sqliteTable('performance', {
  id: text('id').primaryKey(),
  eventId: text('event_id').notNull(),
  date: integer('date').notNull(),
  openTime: text('open_time'),
  startTime: text('start_time'),
});

export const application = sqliteTable('application', {
  id: text('id').primaryKey(),
  eventId: text('event_id').notNull(),
  templateId: text('template_id').notNull(), // BUILTIN_TEMPLATES の id
  templateName: text('template_name').notNull(), // 表示用スナップショット
  meigiId: text('meigi_id'),
  quantity: integer('quantity').notNull().default(1),
  status: text('status').notNull().default('planned'),
  closeAt: integer('close_at'), // 申込締切
  announceAt: integer('announce_at'), // 当落発表
  paymentDeadline: integer('payment_deadline'), // 入金締切
  seatInfo: text('seat_info'),
  totalAmount: real('total_amount'),
  url: text('url'),
  memo: text('memo'),
  createdAt: integer('created_at').notNull(),
});

export const applicationChoice = sqliteTable('application_choice', {
  id: text('id').primaryKey(),
  applicationId: text('application_id').notNull(),
  performanceId: text('performance_id').notNull(),
  priority: integer('priority').notNull().default(1), // 第N希望
});

export const expense = sqliteTable('expense', {
  id: text('id').primaryKey(),
  oshiId: text('oshi_id'),
  eventId: text('event_id'),
  applicationId: text('application_id'),
  category: text('category').notNull(), // ticket/goods/travel/lodging/stream/other
  amount: real('amount').notNull(),
  date: integer('date').notNull(),
  memo: text('memo'),
});

export const budget = sqliteTable('budget', {
  id: text('id').primaryKey(),
  month: text('month').notNull(), // 'YYYY-MM'
  oshiId: text('oshi_id'), // null = 全体
  limitAmount: real('limit_amount').notNull(),
});

export type OshiRow = typeof oshi.$inferSelect;
export type MeigiRow = typeof meigi.$inferSelect;
export type FcRow = typeof fcMembership.$inferSelect;
export type EventRow = typeof event.$inferSelect;
export type PerformanceRow = typeof performance.$inferSelect;
export type ApplicationRow = typeof application.$inferSelect;
export type ExpenseRow = typeof expense.$inferSelect;
