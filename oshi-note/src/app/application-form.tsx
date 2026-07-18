import { addDays } from 'date-fns';
import { useRouter } from 'expo-router';
import React, { useMemo, useState } from 'react';
import { Alert, Linking, Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { eventsForArtist } from '../catalog';
import { CatalogEvent, CatalogSlot } from '../catalog/types';
import { Button, Chip, Field } from '../components/ui';
import { BUILTIN_TEMPLATES } from '../domain/types';
import { useAppStore } from '../state/store';
import { fmtDate, parseDateInput, toInputDate, toInputTime } from '../utils/dates';

export default function ApplicationFormScreen() {
  const router = useRouter();
  const { oshis, addApplication } = useAppStore();

  const [oshiId, setOshiId] = useState(oshis[0]?.id ?? '');
  const [catalogEventId, setCatalogEventId] = useState<string | undefined>(undefined);
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

  const selectedOshi = oshis.find((o) => o.id === oshiId);
  const catalogEvents = useMemo(
    () => (selectedOshi?.catalogArtistId ? eventsForArtist(selectedOshi.catalogArtistId) : []),
    [selectedOshi?.catalogArtistId],
  );
  const catalogEvent: CatalogEvent | undefined = catalogEvents.find((e) => e.id === catalogEventId);
  const template = BUILTIN_TEMPLATES.find((t) => t.id === templateId) ?? BUILTIN_TEMPLATES[0];

  const pickOshi = (id: string) => {
    setOshiId(id);
    setCatalogEventId(undefined);
  };

  const pickCatalogEvent = (e: CatalogEvent) => {
    if (catalogEventId === e.id) {
      setCatalogEventId(undefined);
      return;
    }
    setCatalogEventId(e.id);
    setEventTitle(e.title);
    setVenue(e.venue ?? '');
    setPerfDate(e.dates.length === 1 ? e.dates[0] : '');
    if (e.slots.length > 0) pickSlot(e.slots[0]);
  };

  const pickSlot = (s: CatalogSlot) => {
    setTemplateId(s.templateId);
    if (s.applyEnd) {
      const d = new Date(s.applyEnd);
      setCloseDate(toInputDate(d));
      setCloseTime(toInputTime(d));
    }
    if (s.announceAt) {
      const d = new Date(s.announceAt);
      setAnnounceDate(toInputDate(d));
      setAnnounceTime(toInputTime(d));
    }
    if (s.url) setUrl(s.url);
  };

  // 手入力時: 申込締切を入れたら発表日の初期値をテンプレから提案
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
      Alert.alert('入力エラー', '公演日を選択するか、YYYY-MM-DD 形式で入力してください');
      return;
    }
    const closeAt = closeDate ? parseDateInput(closeDate, closeTime) : null;
    const announceAt = announceDate ? parseDateInput(announceDate, announceTime) : null;
    if ((closeDate && !closeAt) || (announceDate && !announceAt)) {
      Alert.alert('入力エラー', '締切/発表日の日付形式が正しくありません');
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
    const openUrl = url.trim();
    if (openUrl) {
      Alert.alert('登録しました', '通知を自動セットしました。このまま申込ページを開きますか?', [
        { text: 'あとで', style: 'cancel', onPress: () => router.back() },
        {
          text: '申込ページを開く',
          onPress: () => {
            void Linking.openURL(openUrl);
            router.back();
          },
        },
      ]);
    } else {
      router.back();
    }
  };

  return (
    <ScrollView style={styles.container} contentContainerStyle={{ padding: 16, paddingBottom: 48 }}>
      <Text style={styles.label}>推し *</Text>
      <View style={styles.chipRow}>
        {oshis.map((o) => (
          <Chip key={o.id} label={o.name} color={o.color1} selected={oshiId === o.id} onPress={() => pickOshi(o.id)} />
        ))}
        {oshis.length === 0 ? <Text style={styles.hint}>「推し」タブから先に推しを登録してください</Text> : null}
      </View>

      {catalogEvents.length > 0 ? (
        <>
          <Text style={styles.label}>公演を選ぶ(カタログから自動表示)</Text>
          {catalogEvents.map((e) => (
            <Pressable
              key={e.id}
              onPress={() => pickCatalogEvent(e)}
              style={[styles.eventCard, catalogEventId === e.id && { borderColor: selectedOshi?.color1 ?? '#F2545B' }]}
            >
              <Text style={styles.eventTitle}>{e.title}</Text>
              <Text style={styles.eventSub}>
                {e.venue ?? ''}{e.city ? `(${e.city})` : ''} ・ {e.dates.length}日程
              </Text>
            </Pressable>
          ))}
          <Text style={styles.hint}>
            情報は自動収集+検品済みですが、申込前に必ず公式ページでご確認ください。カタログに無い公演は下に手入力できます
          </Text>
        </>
      ) : selectedOshi?.catalogArtistId ? (
        <Text style={styles.hint}>この推しの今後の公演はまだカタログにありません(手入力できます)</Text>
      ) : null}

      {catalogEvent && catalogEvent.dates.length > 1 ? (
        <>
          <Text style={styles.label}>公演日を選ぶ *</Text>
          <View style={styles.chipRow}>
            {catalogEvent.dates.map((d) => (
              <Chip
                key={d}
                label={fmtDate(new Date(`${d}T00:00:00`))}
                color={selectedOshi?.color1}
                selected={perfDate === d}
                onPress={() => setPerfDate(d)}
              />
            ))}
          </View>
        </>
      ) : null}

      {catalogEvent && catalogEvent.slots.length > 0 ? (
        <>
          <Text style={styles.label}>申込枠を選ぶ(締切・発表日を自動入力)</Text>
          <View style={styles.chipRow}>
            {catalogEvent.slots.map((s, i) => (
              <Chip
                key={`${s.templateId}-${i}`}
                label={s.name}
                color={selectedOshi?.color1}
                selected={templateId === s.templateId}
                onPress={() => pickSlot(s)}
              />
            ))}
          </View>
        </>
      ) : null}

      <Field label="公演名 *" value={eventTitle} onChangeText={setEventTitle} placeholder="例: ◯◯ LIVE TOUR 2026 東京" />
      {!catalogEvent || catalogEvent.dates.length <= 1 ? (
        <Field label="公演日 * (YYYY-MM-DD)" value={perfDate} onChangeText={setPerfDate} placeholder="2026-09-20" keyboardType="numbers-and-punctuation" />
      ) : null}
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
  eventCard: {
    backgroundColor: '#fff', borderRadius: 12, padding: 12, marginBottom: 8,
    borderWidth: 1.5, borderColor: 'transparent',
  },
  eventTitle: { fontSize: 14, fontWeight: '700', color: '#1C1C1E' },
  eventSub: { fontSize: 12, color: '#8E8E93', marginTop: 3 },
});
