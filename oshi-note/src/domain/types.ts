// 申込ステータス(当落ライフサイクル)
// won は「当選=入金待ち」を含む(入金期限は paymentDeadline で管理)
export type ApplicationStatus =
  | 'planned' // 申込予定
  | 'applied' // 申込済(当落待ち)
  | 'won' // 当選(入金待ち)
  | 'paid' // 入金済(発券待ち)
  | 'ticketed' // 発券済
  | 'lost' // 落選
  | 'expired'; // 期限切れ

export const STATUS_LABEL: Record<ApplicationStatus, string> = {
  planned: '申込予定',
  applied: '当落待ち',
  won: '当選・入金待ち',
  paid: '入金済',
  ticketed: '発券済',
  lost: '落選',
  expired: '期限切れ',
};

export const STATUS_COLOR: Record<ApplicationStatus, string> = {
  planned: '#8E8E93',
  applied: '#5A9CF8',
  won: '#F2545B',
  paid: '#34C759',
  ticketed: '#2AA198',
  lost: '#C7C7CC',
  expired: '#FF3B30',
};

export type NotificationKind =
  | 'applyCloseSoon' // 申込締切 前日
  | 'applyCloseToday' // 申込締切 当日
  | 'announce' // 当落発表
  | 'announceFollowUp' // 結果未記録フォロー
  | 'paymentEve' // 入金締切 前日
  | 'paymentDay' // 入金締切 当日
  | 'fcRenewalWeek' // FC更新 7日前
  | 'fcRenewalEve' // FC更新 前日
  | 'eventEve'; // 公演 前日

// 先行種別テンプレ(発表日・入金期間の初期値。すべて画面上で上書き可能)
export interface TemplateDef {
  id: string;
  name: string;
  announceOffsetDays: number; // 申込締切 → 当落発表 の日数目安
  paymentWindowDays: number; // 当落発表 → 入金締切 の日数目安
}

export const BUILTIN_TEMPLATES: TemplateDef[] = [
  { id: 'fc', name: 'FC先行', announceOffsetDays: 7, paymentWindowDays: 3 },
  { id: 'pia', name: 'ぴあ先行', announceOffsetDays: 5, paymentWindowDays: 3 },
  { id: 'lawson', name: 'ローチケ先行', announceOffsetDays: 5, paymentWindowDays: 3 },
  { id: 'eplus', name: 'イープラス先行', announceOffsetDays: 5, paymentWindowDays: 3 },
  { id: 'general', name: '一般発売', announceOffsetDays: 0, paymentWindowDays: 1 },
  { id: 'custom', name: 'その他', announceOffsetDays: 7, paymentWindowDays: 3 },
];

export const OSHI_COLORS = [
  '#F2545B', '#FF8A5C', '#F5C518', '#7BC950', '#34C759',
  '#2AA198', '#5AC8FA', '#5A9CF8', '#7B61FF', '#C969E6',
  '#FF6FA5', '#8E8E93',
];
