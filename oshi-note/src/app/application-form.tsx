import { addDays } from 'date-fns';
import { useRouter } from 'expo-router';
import React, { useState } from 'react';
import { Alert, ScrollView, StyleSheet, Text, View } from 'react-native';
import { Button, Chip, Field } from '../components/ui';
import { BUILTIN_TEMPLATES } from '../domain/types';
import { useAppStore } from '../state/store';
import { parseDateInput, toInputDate } from '../utils/dates';

export default function ApplicationFormScreen() {
  const router = useRouter();
  const { oshis, addApplication } = useAppStore();

  const [oshiId, setOshiId] = useState(oshis[0]?.id ?? '');
  const [eventTitle, setEventTitle] = useState('');
  const [perfDate, setPerfDate] = useState('');
  const [venue, setVenue] = useState('');
  const [templateId, setTemplateId] = useState(BUILTIN_TEMPLATES[0].id);
  const [meigiLabel, setMeigiLabel] = useState('');
  const [quantity, setQuantity] = useState('1');
  const [closeDate, setCloseDate] = useState('');
  const [closeTime, setCloseTime] = useState('23:59');
  const [announceDate, setAnnounceDate] = useState('');
  const [announceTime, setAnnounceTime] = useState('13:00');
  const [url, setUrl] = useState('');

  const template = BUILTIN_TEMPLATES.find((t) => t.id === templateId) ?? BUILTIN_TEMPLATES[0];

  // 申込締切を入れたら、発表日の初期値をテンプレから提案
  const onCloseDateChange = (v: string) => {
    setCloseDate(v);
    if (!announceDate) {
      const parsed = parseDateInput(v);
      if (parsed) setAnnounceDate(toInputDate(addDays(parsed, template.announceOffsetDays)));
    }
  };

  const save = () => {
    if (!oshiId) {
      Alert.alert('推しが未登録です', '先に「推し」タブから推しを登録してください');
      return;
    }
    if (!eventTitle.trim()) {
      Alert.alert('入力エラー', '公演名を入力してください');
      return;
    }
    const performanceDate = parseDateInput(perfDate);
    if (!performanceDate) {
      Alert.alert('入力エラー', '公演日を YYYY-MM-DD 形式で入力してください(例: 2026-09-20)');
      return;
    }
    const closeAt = closeDate ? parseDateInput(closeDate, closeTime) : null;
    if (closeDate && !closeAt) {
      Alert.alert('入力エラー', '申込締切の日付形式が正しくありません');
      return;
    }
    const announceAt = announceDate ? parseDateInput(announceDate, announceTime) : null;
    if (announceDate && !announceAt) {
      Alert.alert('入力エラー', '当落発表の日付形式が正しくありません');
      return;
    }
    addApplication({
      oshiId,
      eventTitle: eventTitle.trim(),
      performanceDate,
      venue: venue.trim() || undefined,
      templateId: template.id,
      templateName: template.name,
      meigiLabel: meigiLabel.trim() || undefined,
      quantity: Math.max(1, parseInt(quantity, 10) || 1),
      closeAt,
      announceAt,
      url: url.trim() || undefined,
    });
    router.back();
  };

  return (
    <ScrollView style={styles.container} contentContainerStyle={{ padding: 16, paddingBottom: 48 }}>
      <Text style={styles.label}>推し *</Text>
      <View style={styles.chipRow}>
        {oshis.map((o) => (
          <Chip key={o.id} label={o.name} color={o.color1} selected={oshiId === o.id} onPress={() => setOshiId(o.id)} />
        ))}
        {oshis.length === 0 ? <Text style={styles.hint}>「推し」タブから先に推しを登録してください</Text> : null}
      </View>

      <Field label="公演名 *" value={eventTitle} onChangeText={setEventTitle} placeholder="例: ◯◯ LIVE TOUR 2026 東京" />
      <Field label="公演日 * (YYYY-MM-DD)" value={perfDate} onChangeText={setPerfDate} placeholder="2026-09-20" keyboardType="numbers-and-punctuation" />
      <Field label="会場" value={venue} onChangeText={setVenue} placeholder="例: 東京ドーム" />

      <Text style={styles.label}>申込枠(先行種別)</Text>
      <View style={styles.chipRow}>
        {BUILTIN_TEMPLATES.map((t) => (
          <Chip key={t.id} label={t.name} selected={templateId === t.id} onPress={() => setTemplateId(t.id)} />
        ))}
      </View>

      <Field label="名義" value={meigiLabel} onChangeText={setMeigiLabel} placeholder="例: 自名義 / 母名義" />
      <Field label="枚数" value={quantity} onChangeText={setQuantity} keyboardType="number-pad" />

      <View style={styles.row}>
        <View style={{ flex: 2 }}>
          <Field label="申込締切 (YYYY-MM-DD)" value={closeDate} onChangeText={onCloseDateChange} placeholder="2026-08-10" keyboardType="numbers-and-punctuation" />
        </View>
        <View style={{ flex: 1, marginLeft: 10 }}>
          <Field label="時刻" value={closeTime} onChangeText={setCloseTime} placeholder="23:59" keyboardType="numbers-and-punctuation" />
        </View>
      </View>

      <View style={styles.row}>
        <View style={{ flex: 2 }}>
          <Field label="当落発表 (YYYY-MM-DD)" value={announceDate} onChangeText={setAnnounceDate} placeholder="2026-08-17" keyboardType="numbers-and-punctuation" />
        </View>
        <View style={{ flex: 1, marginLeft: 10 }}>
          <Field label="時刻" value={announceTime} onChangeText={setAnnounceTime} placeholder="13:00" keyboardType="numbers-and-punctuation" />
        </View>
      </View>
      <Text style={styles.hint}>
        締切を入れると発表日の目安を自動入力します({template.name}: 締切+{template.announceOffsetDays}日)。実際の発表日に合わせて修正してください
      </Text>

      <Field label="申込ページURL" value={url} onChangeText={setUrl} placeholder="https://..." autoCapitalize="none" />

      <View style={{ height: 8 }} />
      <Button title="登録する(通知を自動セット)" onPress={save} />
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#F6F6F8' },
  label: { fontSize: 12, color: '#6B6B70', marginBottom: 6, fontWeight: '600' },
  chipRow: { flexDirection: 'row', flexWrap: 'wrap', marginBottom: 10 },
  row: { flexDirection: 'row' },
  hint: { fontSize: 12, color: '#9A9AA0', marginBottom: 12, lineHeight: 18 },
});
