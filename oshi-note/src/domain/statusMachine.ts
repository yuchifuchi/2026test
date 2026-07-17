import { ApplicationStatus } from './types';

// 当落ライフサイクルの許可された遷移
export const TRANSITIONS: Record<ApplicationStatus, ApplicationStatus[]> = {
  planned: ['applied'],
  applied: ['won', 'lost'],
  won: ['paid', 'expired'],
  paid: ['ticketed'],
  ticketed: [],
  lost: ['won'], // 復活当選
  expired: [],
};

export function canTransition(from: ApplicationStatus, to: ApplicationStatus): boolean {
  return TRANSITIONS[from]?.includes(to) ?? false;
}

export function nextStatuses(from: ApplicationStatus): ApplicationStatus[] {
  return TRANSITIONS[from] ?? [];
}
