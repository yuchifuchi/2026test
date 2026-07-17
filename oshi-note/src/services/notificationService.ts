import * as Notifications from 'expo-notifications';
import { Platform } from 'react-native';
import { PlannedNotification } from '../domain/notificationPlanner';

let handlerConfigured = false;

export function configureNotificationHandler(): void {
  if (handlerConfigured) return;
  handlerConfigured = true;
  Notifications.setNotificationHandler({
    handleNotification: async () => ({
      shouldShowBanner: true,
      shouldShowList: true,
      shouldPlaySound: true,
      shouldSetBadge: false,
    }),
  });
}

export async function ensurePermissions(): Promise<boolean> {
  const settings = await Notifications.getPermissionsAsync();
  if (settings.granted) return true;
  const req = await Notifications.requestPermissionsAsync();
  return req.granted;
}

// 方式: 全キャンセル→全再登録(データ変更との整合を常に保つ最も単純で堅牢な方法)
export async function resyncScheduledNotifications(plans: PlannedNotification[]): Promise<void> {
  if (Platform.OS === 'web') return; // Webはスケジュール通知非対応
  try {
    await Notifications.cancelAllScheduledNotificationsAsync();
    for (const p of plans) {
      await Notifications.scheduleNotificationAsync({
        content: { title: p.title, body: p.body, sound: true },
        trigger: {
          type: Notifications.SchedulableTriggerInputTypes.DATE,
          date: p.fireAt,
        },
      });
    }
  } catch (e) {
    // 通知の失敗でアプリ本体を止めない(権限未許可など)
    console.warn('notification resync failed', e);
  }
}
