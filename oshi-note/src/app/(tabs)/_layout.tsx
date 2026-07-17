import { Tabs } from 'expo-router';
import React from 'react';
import { Text } from 'react-native';

function icon(emoji: string) {
  return ({ focused }: { focused: boolean }) => (
    <Text style={{ fontSize: 20, opacity: focused ? 1 : 0.45 }}>{emoji}</Text>
  );
}

export default function TabsLayout() {
  return (
    <Tabs
      screenOptions={{
        tabBarActiveTintColor: '#F2545B',
        headerTitleStyle: { fontWeight: '700' },
      }}
    >
      <Tabs.Screen name="index" options={{ title: 'ホーム', tabBarIcon: icon('🏠') }} />
      <Tabs.Screen name="touraku" options={{ title: '当落', tabBarIcon: icon('🎫') }} />
      <Tabs.Screen name="calendar" options={{ title: 'カレンダー', tabBarIcon: icon('📅') }} />
      <Tabs.Screen name="budget" options={{ title: '収支', tabBarIcon: icon('💰') }} />
      <Tabs.Screen name="oshi" options={{ title: '推し', tabBarIcon: icon('💖') }} />
    </Tabs>
  );
}
