import { planAll, planForApplication, planForFc } from '../notificationPlanner';
import { canTransition, nextStatuses } from '../statusMachine';

const d = (s: string) => new Date(s);

describe('notificationPlanner', () => {
  test('申込予定: 締切前日20時と当日12時に通知される', () => {
    const plans = planForApplication({
      id: 'a1',
      title: 'ドームツアー東京',
      status: 'planned',
      closeAt: d('2026-08-10T23:59:00'),
    });
    expect(plans).toHaveLength(2);
    expect(plans[0].fireAt).toEqual(d('2026-08-09T20:00:00'));
    expect(plans[1].fireAt).toEqual(d('2026-08-10T12:00:00'));
  });

  test('当落待ち: 発表時刻ちょうど+同日21時のフォロー', () => {
    const plans = planForApplication({
      id: 'a2',
      title: 'アリーナ公演',
      status: 'applied',
      announceAt: d('2026-08-15T13:00:00'),
    });
    expect(plans.map((p) => p.kind)).toEqual(['announce', 'announceFollowUp']);
    expect(plans[0].fireAt).toEqual(d('2026-08-15T13:00:00'));
    expect(plans[1].fireAt).toEqual(d('2026-08-15T21:00:00'));
  });

  test('当選(入金待ち): 入金締切の前日20時と当日朝9時に通知される', () => {
    const plans = planForApplication({
      id: 'a3',
      title: 'ファンミ大阪',
      status: 'won',
      paymentDeadline: d('2026-08-20T23:59:00'),
    });
    expect(plans.map((p) => p.kind)).toEqual(['paymentEve', 'paymentDay']);
    expect(plans[0].fireAt).toEqual(d('2026-08-19T20:00:00'));
    expect(plans[1].fireAt).toEqual(d('2026-08-20T09:00:00'));
  });

  test('入金済・落選などは申込系の通知を出さない', () => {
    for (const status of ['paid', 'ticketed', 'lost', 'expired'] as const) {
      const plans = planForApplication({
        id: 'x',
        title: 't',
        status,
        closeAt: d('2026-08-10T23:59:00'),
        announceAt: d('2026-08-15T13:00:00'),
        paymentDeadline: d('2026-08-20T23:59:00'),
      });
      expect(plans).toHaveLength(0);
    }
  });

  test('planAll: 過去の通知は除外され、時刻順に並ぶ', () => {
    const now = d('2026-08-15T00:00:00');
    const plans = planAll(
      {
        applications: [
          { id: 'past', title: '過去', status: 'planned', closeAt: d('2026-08-01T23:59:00') },
          { id: 'fut', title: '未来', status: 'won', paymentDeadline: d('2026-08-20T23:59:00') },
        ],
        fcs: [{ id: 'fc1', name: '推しFC', renewalDate: d('2026-08-18T00:00:00') }],
      },
      now,
    );
    expect(plans.every((p) => p.fireAt.getTime() > now.getTime())).toBe(true);
    const times = plans.map((p) => p.fireAt.getTime());
    expect([...times].sort((a, b) => a - b)).toEqual(times);
    // FC更新7日前(8/11)は過去なので除外、前日(8/17 20時)は含まれる
    expect(plans.some((p) => p.kind === 'fcRenewalEve')).toBe(true);
    expect(plans.some((p) => p.kind === 'fcRenewalWeek')).toBe(false);
  });

  test('FC更新: 7日前と前日の20時', () => {
    const plans = planForFc({ id: 'f1', name: 'FC', renewalDate: d('2026-09-10T00:00:00') });
    expect(plans[0].fireAt).toEqual(d('2026-09-03T20:00:00'));
    expect(plans[1].fireAt).toEqual(d('2026-09-09T20:00:00'));
  });
});

describe('statusMachine', () => {
  test('正しい遷移のみ許可される', () => {
    expect(canTransition('applied', 'won')).toBe(true);
    expect(canTransition('applied', 'lost')).toBe(true);
    expect(canTransition('won', 'paid')).toBe(true);
    expect(canTransition('lost', 'won')).toBe(true); // 復活当選
    expect(canTransition('planned', 'won')).toBe(false);
    expect(canTransition('paid', 'expired')).toBe(false);
    expect(nextStatuses('ticketed')).toEqual([]);
  });
});
