import { addDays, isSameDay, set } from 'date-fns';
import { ApplicationStatus, NotificationKind } from './types';

export interface PlannedNotification {
  key: string; // 重複防止用の一意キー
  kind: NotificationKind;
  fireAt: Date;
  title: string;
  body: string;
}

export interface AppForPlan {
  id: string;
  title: string; // 公演名
  status: ApplicationStatus;
  closeAt?: Date | null; // 申込締切
  announceAt?: Date | null; // 当落発表
  paymentDeadline?: Date | null; // 入金締切
}

export interface FcForPlan {
  id: string;
  name: string;
  renewalDate?: Date | null;
}

export interface PerformanceForPlan {
  id: string;
  title: string;
  date: Date;
}

const at = (base: Date, hours: number, minutes = 0) =>
  set(base, { hours, minutes, seconds: 0, milliseconds: 0 });

export function planForApplication(a: AppForPlan): PlannedNotification[] {
  const out: PlannedNotification[] = [];
  if (a.status === 'planned' && a.closeAt) {
    out.push({
      key: `app:${a.id}:closeSoon`,
      kind: 'applyCloseSoon',
      fireAt: at(addDays(a.closeAt, -1), 20),
      title: '申込締切が近づいています',
      body: `「${a.title}」の申込締切は明日です。忘れずに申し込みましょう`,
    });
    out.push({
      key: `app:${a.id}:closeToday`,
      kind: 'applyCloseToday',
      fireAt: at(a.closeAt, 12),
      title: '今日が申込締切です',
      body: `「${a.title}」の申込は今日まで。まだの場合は今すぐ!`,
    });
  }
  if (a.status === 'applied' && a.announceAt) {
    out.push({
      key: `app:${a.id}:announce`,
      kind: 'announce',
      fireAt: a.announceAt,
      title: '当落発表の時間です',
      body: `「${a.title}」の結果を確認して、アプリに記録しましょう`,
    });
    out.push({
      key: `app:${a.id}:announceFollow`,
      kind: 'announceFollowUp',
      fireAt: at(a.announceAt, 21),
      title: '当落結果は記録しましたか?',
      body: `「${a.title}」の結果が未記録です。当選なら入金期限の管理を始めます`,
    });
  }
  if (a.status === 'won' && a.paymentDeadline) {
    out.push({
      key: `app:${a.id}:payEve`,
      kind: 'paymentEve',
      fireAt: at(addDays(a.paymentDeadline, -1), 20),
      title: '入金締切は明日です',
      body: `「${a.title}」のチケット代の入金を忘れずに`,
    });
    out.push({
      key: `app:${a.id}:payDay`,
      kind: 'paymentDay',
      fireAt: at(a.paymentDeadline, 9),
      title: '⚠️ 今日中に入金してください',
      body: `「${a.title}」の入金締切は今日です。期限を過ぎると当選が無効になります`,
    });
  }
  return out;
}

export function planForFc(fc: FcForPlan): PlannedNotification[] {
  if (!fc.renewalDate) return [];
  return [
    {
      key: `fc:${fc.id}:week`,
      kind: 'fcRenewalWeek',
      fireAt: at(addDays(fc.renewalDate, -7), 20),
      title: 'FC更新が近づいています',
      body: `「${fc.name}」の更新期限まであと1週間です`,
    },
    {
      key: `fc:${fc.id}:eve`,
      kind: 'fcRenewalEve',
      fireAt: at(addDays(fc.renewalDate, -1), 20),
      title: 'FC更新期限は明日です',
      body: `「${fc.name}」の更新を忘れずに`,
    },
  ];
}

export function planForPerformance(p: PerformanceForPlan): PlannedNotification[] {
  return [
    {
      key: `perf:${p.id}:eve`,
      kind: 'eventEve',
      fireAt: at(addDays(p.date, -1), 21),
      title: '明日は現場です!',
      body: `「${p.title}」チケット・身分証・充電の準備はOK?`,
    },
  ];
}

// iOSのスケジュール通知上限(64件)を考慮して直近60件に絞る
const MAX_SCHEDULED = 60;

export function planAll(
  input: {
    applications: AppForPlan[];
    fcs?: FcForPlan[];
    performances?: PerformanceForPlan[];
  },
  now: Date,
): PlannedNotification[] {
  const all = [
    ...input.applications.flatMap(planForApplication),
    ...(input.fcs ?? []).flatMap(planForFc),
    ...(input.performances ?? []).flatMap(planForPerformance),
  ];
  return all
    .filter((n) => n.fireAt.getTime() > now.getTime())
    .sort((a, b) => a.fireAt.getTime() - b.fireAt.getTime())
    .slice(0, MAX_SCHEDULED);
}

// 当落発表が「今日」の申込があるか(ホーム画面の強調表示用)
export function isAnnounceToday(a: AppForPlan, now: Date): boolean {
  return a.status === 'applied' && !!a.announceAt && isSameDay(a.announceAt, now);
}
