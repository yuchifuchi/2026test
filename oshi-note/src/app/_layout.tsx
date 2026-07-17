import { Stack } from 'expo-router';
import { useEffect } from 'react';
import { migrate } from '../db/client';
import { configureNotificationHandler, ensurePermissions } from '../services/notificationService';
import { useAppStore } from '../state/store';

// アプリ起動時に一度だけ: DBマイグレーション → 通知ハンドラ設定
migrate();
configureNotificationHandler();

export default function RootLayout() {
  const refresh = useAppStore((s) => s.refresh);

  useEffect(() => {
    refresh(); // データ読込+通知スケジュール再計算
    void ensurePermissions();
  }, [refresh]);

  return (
    <Stack>
      <Stack.Screen name="(tabs)" options={{ headerShown: false }} />
      <Stack.Screen
        name="oshi-form"
        options={{ presentation: 'modal', title: '推しを追加' }}
      />
      <Stack.Screen
        name="application-form"
        options={{ presentation: 'modal', title: '申込を登録' }}
      />
      <Stack.Screen name="application/[id]" options={{ title: '申込の詳細' }} />
    </Stack>
  );
}
