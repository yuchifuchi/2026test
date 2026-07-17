import { create } from 'zustand';
import { newId } from '../db/client';
import {
  ApplicationView, createApplicationWithEvent, createOshi, deleteApplication, deleteOshi,
  listApplications, listFc, listOshi, NewApplicationInput, updateApplication,
} from '../db/repos';
import { OshiRow } from '../db/schema';
import { planAll } from '../domain/notificationPlanner';
import { ApplicationStatus } from '../domain/types';
import { resyncScheduledNotifications } from '../services/notificationService';

function buildPlansAndResync(applications: ApplicationView[]): void {
  const plans = planAll(
    {
      applications: applications.map((a) => ({
        id: a.id,
        title: a.eventTitle,
        status: a.status as ApplicationStatus,
        closeAt: a.closeAt ? new Date(a.closeAt) : null,
        announceAt: a.announceAt ? new Date(a.announceAt) : null,
        paymentDeadline: a.paymentDeadline ? new Date(a.paymentDeadline) : null,
      })),
      fcs: listFc().map((f) => ({
        id: f.id,
        name: f.name,
        renewalDate: f.renewalDate && f.notifyEnabled ? new Date(f.renewalDate) : null,
      })),
      // 参戦が確定している(当選以降の)公演のみ前日リマインド
      performances: applications
        .filter((a) => ['won', 'paid', 'ticketed'].includes(a.status) && a.performanceDate)
        .map((a) => ({ id: a.id, title: a.eventTitle, date: a.performanceDate as Date })),
    },
    new Date(),
  );
  void resyncScheduledNotifications(plans);
}

interface AppState {
  oshis: OshiRow[];
  applications: ApplicationView[];
  loaded: boolean;
  refresh: () => void;
  addOshi: (input: { name: string; color1: string; icon?: string; genre?: string }) => void;
  removeOshi: (id: string) => void;
  addApplication: (input: NewApplicationInput) => void;
  patchApplication: (
    id: string,
    patch: Parameters<typeof updateApplication>[1],
  ) => void;
  removeApplication: (id: string) => void;
}

export const useAppStore = create<AppState>((set, get) => ({
  oshis: [],
  applications: [],
  loaded: false,

  refresh: () => {
    const applications = listApplications();
    set({ oshis: listOshi(), applications, loaded: true });
    buildPlansAndResync(applications);
  },

  addOshi: (input) => {
    createOshi(input);
    get().refresh();
  },

  removeOshi: (id) => {
    deleteOshi(id);
    get().refresh();
  },

  addApplication: (input) => {
    createApplicationWithEvent(input);
    get().refresh();
  },

  patchApplication: (id, patch) => {
    updateApplication(id, patch);
    get().refresh();
  },

  removeApplication: (id) => {
    deleteApplication(id);
    get().refresh();
  },
}));

export { newId };
