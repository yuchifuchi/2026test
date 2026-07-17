import { useRouter } from 'expo-router';
import React from 'react';
import { Alert, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { Button, Empty } from '../../components/ui';
import { useAppStore } from '../../state/store';

export default function OshiScreen() {
  const router = useRouter();
  const { oshis, applications, removeOshi } = useAppStore();

  const confirmDelete = (id: string, name: string) => {
    Alert.alert('推しを削除', `「${name}」を削除しますか?(申込データは残ります)`, [
      { text: 'キャンセル', style: 'cancel' },
      { text: '削除', style: 'destructive', onPress: () => removeOshi(id) },
    ]);
  };

  return (
    <ScrollView style={styles.container} contentContainerStyle={{ padding: 16 }}>
      {oshis.length === 0 ? (
        <Empty title="推しが登録されていません" hint="推しを登録すると申込・収支を推しごとに管理できます" />
      ) : (
        oshis.map((o) => {
          const count = applications.filter((a) => a.oshiId === o.id).length;
          return (
            <Pressable key={o.id} style={styles.card} onLongPress={() => confirmDelete(o.id, o.name)}>
              <View style={[styles.avatar, { backgroundColor: o.color1 }]}>
                <Text style={{ fontSize: 22 }}>{o.icon ?? '💖'}</Text>
              </View>
              <View style={{ flex: 1 }}>
                <Text style={styles.name}>{o.name}</Text>
                <Text style={styles.sub}>
                  {o.genre ? `${o.genre} ・ ` : ''}申込 {count}件
                </Text>
              </View>
              <View style={[styles.colorDot, { backgroundColor: o.color1 }]} />
            </Pressable>
          );
        })
      )}
      <View style={{ height: 12 }} />
      <Button title="＋ 推しを追加する" onPress={() => router.push('/oshi-form')} />
      {oshis.length > 0 ? (
        <Text style={styles.hint}>長押しで削除できます</Text>
      ) : null}
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#F6F6F8' },
  card: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: '#fff',
    borderRadius: 14,
    padding: 14,
    marginBottom: 10,
    gap: 12,
  },
  avatar: {
    width: 44, height: 44, borderRadius: 22,
    alignItems: 'center', justifyContent: 'center',
  },
  name: { fontSize: 16, fontWeight: '700', color: '#1C1C1E' },
  sub: { fontSize: 12, color: '#8E8E93', marginTop: 2 },
  colorDot: { width: 14, height: 14, borderRadius: 7 },
  hint: { fontSize: 12, color: '#9A9AA0', textAlign: 'center', marginTop: 10 },
});
