import { eq } from 'drizzle-orm';
import { ApplicationStatus } from '../domain/types';
import { db, newId } from './client';
import {
  application, applicationChoice, event, fcMembership, meigi, oshi, performance,
  ApplicationRow, EventRow, FcRow, MeigiRow, OshiRow, PerformanceRow,
} from './schema';

// ---- 推し ----

export function listOshi(): OshiRow[] {
  return db.select().from(oshi).all().sort((a, b) => a.sortOrder - b.sortOrder);
}

export function createOshi(input: {
  name: string;
  color1: string;
  icon?: string;
  genre?: string;
  catalogArtistId?: string;
}): OshiRow {
  const row: OshiRow = {
    id: newId(),
    name: input.name,
    genre: input.genre ?? null,
    catalogArtistId: input.catalogArtistId ?? null,
    color1: input.color1,
    color2: null,
    icon: input.icon ?? null,
    sortOrder: listOshi().length,
    createdAt: Date.now(),
  };
  db.insert(oshi).values(row).run();
  return row;
}

export function deleteOshi(id: string): void {
  db.delete(oshi).where(eq(oshi.id, id)).run();
}

// ---- 名義 ----

export function listMeigi(): MeigiRow[] {
  return db.select().from(meigi).all();
}

export function findOrCreateMeigi(label: string): MeigiRow {
  const trimmed = label.trim();
  const found = listMeigi().find((m) => m.label === trimmed);
  if (found) return found;
  const row: MeigiRow = { id: newId(), label: trimmed, memo: null, sortOrder: listMeigi().length };
  db.insert(meigi).values(row).run();
  return row;
}

// ---- FC ----

export function listFc(): FcRow[] {
  return db.select().from(fcMembership).all();
}

// ---- 申込(公演とセットで作成) ----

export interface NewApplicationInput {
  oshiId: string;
  eventTitle: string;
  performanceDate: Date;
  venue?: string;
  templateId: string;
  templateName: string;
  meigiLabel?: string;
  quantity: number;
  closeAt?: Date | null;
  announceAt?: Date | null;
  url?: string;
  memo?: string;
}

export function createApplicationWithEvent(input: NewApplicationInput): ApplicationRow {
  const eventRow: EventRow = {
    id: newId(),
    oshiId: input.oshiId,
    title: input.eventTitle,
    type: 'live',
    venue: input.venue ?? null,
    city: null,
    memo: null,
  };
  db.insert(event).values(eventRow).run();

  const perfRow: PerformanceRow = {
    id: newId(),
    eventId: eventRow.id,
    date: input.performanceDate.getTime(),
    openTime: null,
    startTime: null,
  };
  db.insert(performance).values(perfRow).run();

  const meigiRow = input.meigiLabel ? findOrCreateMeigi(input.meigiLabel) : null;

  const appRow: ApplicationRow = {
    id: newId(),
    eventId: eventRow.id,
    templateId: input.templateId,
    templateName: input.templateName,
    meigiId: meigiRow?.id ?? null,
    quantity: input.quantity,
    status: 'planned',
    closeAt: input.closeAt ? input.closeAt.getTime() : null,
    announceAt: input.announceAt ? input.announceAt.getTime() : null,
    paymentDeadline: null,
    seatInfo: null,
    totalAmount: null,
    url: input.url ?? null,
    memo: input.memo ?? null,
    createdAt: Date.now(),
  };
  db.insert(application).values(appRow).run();
  db.insert(applicationChoice)
    .values({ id: newId(), applicationId: appRow.id, performanceId: perfRow.id, priority: 1 })
    .run();
  return appRow;
}

export interface ApplicationView extends ApplicationRow {
  eventTitle: string;
  oshiId: string;
  oshiName: string;
  oshiColor: string;
  meigiLabel: string | null;
  performanceDate: Date | null;
}

export function listApplications(): ApplicationView[] {
  const apps = db.select().from(application).all();
  const events = new Map(db.select().from(event).all().map((e) => [e.id, e]));
  const oshis = new Map(listOshi().map((o) => [o.id, o]));
  const meigis = new Map(listMeigi().map((m) => [m.id, m]));
  const perfs = db.select().from(performance).all();
  const firstPerfByEvent = new Map<string, PerformanceRow>();
  for (const p of perfs.sort((a, b) => a.date - b.date)) {
    if (!firstPerfByEvent.has(p.eventId)) firstPerfByEvent.set(p.eventId, p);
  }
  return apps.map((a) => {
    const ev = events.get(a.eventId);
    const os = ev ? oshis.get(ev.oshiId) : undefined;
    const perf = ev ? firstPerfByEvent.get(ev.id) : undefined;
    return {
      ...a,
      eventTitle: ev?.title ?? '(不明な公演)',
      oshiId: ev?.oshiId ?? '',
      oshiName: os?.name ?? '',
      oshiColor: os?.color1 ?? '#8E8E93',
      meigiLabel: a.meigiId ? (meigis.get(a.meigiId)?.label ?? null) : null,
      performanceDate: perf ? new Date(perf.date) : null,
    };
  });
}

export function getApplication(id: string): ApplicationView | undefined {
  return listApplications().find((a) => a.id === id);
}

export function updateApplication(
  id: string,
  patch: Partial<{
    status: ApplicationStatus;
    paymentDeadline: Date | null;
    seatInfo: string | null;
    totalAmount: number | null;
    announceAt: Date | null;
    closeAt: Date | null;
    memo: string | null;
  }>,
): void {
  const values: Record<string, unknown> = {};
  if (patch.status !== undefined) values.status = patch.status;
  if (patch.paymentDeadline !== undefined)
    values.paymentDeadline = patch.paymentDeadline ? patch.paymentDeadline.getTime() : null;
  if (patch.announceAt !== undefined)
    values.announceAt = patch.announceAt ? patch.announceAt.getTime() : null;
  if (patch.closeAt !== undefined) values.closeAt = patch.closeAt ? patch.closeAt.getTime() : null;
  if (patch.seatInfo !== undefined) values.seatInfo = patch.seatInfo;
  if (patch.totalAmount !== undefined) values.totalAmount = patch.totalAmount;
  if (patch.memo !== undefined) values.memo = patch.memo;
  if (Object.keys(values).length === 0) return;
  db.update(application).set(values).where(eq(application.id, id)).run();
}

export function deleteApplication(id: string): void {
  db.delete(applicationChoice).where(eq(applicationChoice.applicationId, id)).run();
  db.delete(application).where(eq(application.id, id)).run();
}
