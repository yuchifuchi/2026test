import { format, isValid, parse } from 'date-fns';
import { ja } from 'date-fns/locale';

// 'YYYY-MM-DD' (+ 任意で 'HH:mm') → Date。不正なら null
export function parseDateInput(dateStr: string, timeStr?: string): Date | null {
  const d = dateStr.trim();
  if (!d) return null;
  const t = (timeStr ?? '').trim();
  const parsed = t
    ? parse(`${d} ${t}`, 'yyyy-MM-dd HH:mm', new Date())
    : parse(d, 'yyyy-MM-dd', new Date());
  return isValid(parsed) ? parsed : null;
}

export function fmtDate(d: Date | null | undefined): string {
  return d ? format(d, 'M/d(EEE)', { locale: ja }) : '未設定';
}

export function fmtDateTime(d: Date | null | undefined): string {
  return d ? format(d, 'M/d(EEE) HH:mm', { locale: ja }) : '未設定';
}

export function toInputDate(d: Date | null | undefined): string {
  return d ? format(d, 'yyyy-MM-dd') : '';
}

export function toInputTime(d: Date | null | undefined): string {
  return d ? format(d, 'HH:mm') : '';
}

export function daysUntil(d: Date, now: Date = new Date()): number {
  const a = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const b = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  return Math.round((b - a) / 86400000);
}
