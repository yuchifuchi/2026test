import { useRouter } from 'expo-router';
import React, { useState } from 'react';
import { Alert, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { Button, Field } from '../components/ui';
import { OSHI_COLORS } from '../domain/types';
import { useAppStore } from '../state/store';

export default function OshiFormScreen() {
  const router = useRouter();
  const addOshi = useAppStore((s) => s.addOshi);
  const [name, setName] = useState('');
  const [genre, setGenre] = useState('');
  const [icon, setIcon] = useState('');
  const [color, setColor] = useState(OSHI_COLORS[0]);

  const save = () => {
    if (!name.trim()) {
      Alert.alert('入力エラー', '推しの名前を入力してください');
      return;
    }
    addOshi({
      name: name.trim(),
      color1: color,
      genre: genre.trim() || undefined,
      icon: icon.trim() || undefined,
    });
    router.back();
  };

  return (
    <ScrollView style={styles.container} contentContainerStyle={{ padding: 16 }}>
      <Field label="推しの名前 *" value={name} onChangeText={setName} placeholder="例: 山田太郎 / ◯◯(グループ名)" />
      <Field label="ジャンル" value={genre} onChangeText={setGenre} placeholder="例: アイドル / K-POP / 2.5次元 / VTuber" />
      <Field label="アイコン絵文字" value={icon} onChangeText={setIcon} placeholder="例: 🎤(未入力なら💖)" maxLength={2} />

      <Text style={styles.label}>推し色</Text>
      <View style={styles.palette}>
        {OSHI_COLORS.map((c) => (
          <Pressable
            key={c}
            onPress={() => setColor(c)}
            style={[styles.swatch, { backgroundColor: c }, color === c && styles.swatchSelected]}
          />
        ))}
      </View>

      <View style={{ height: 16 }} />
      <Button title="保存する" color={color} onPress={save} />
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#F6F6F8' },
  label: { fontSize: 12, color: '#6B6B70', marginBottom: 8, fontWeight: '600' },
  palette: { flexDirection: 'row', flexWrap: 'wrap', gap: 10 },
  swatch: { width: 40, height: 40, borderRadius: 20 },
  swatchSelected: { borderWidth: 3, borderColor: '#1C1C1E' },
});
