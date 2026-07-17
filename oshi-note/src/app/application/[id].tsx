import { addDays } from 'date-fns';
import { useLocalSearchParams, useRouter } from 'expo-router';
import React, { useState } from 'react';
import { Alert, ScrollView, StyleSheet, Text, View } from 'react-native';
import { Button, Field } from '../../components/ui';
import { nextStatuses } from '../../domain/statusMachine';
import { ApplicationStatus, BUILTIN_TEMPLATES, STATUS_COLOR, STATUS_LABEL } from '../../domain/types';
import { useAppStore } from '../../state/store';
import { fmtDate, fmtDateTime, parseDateInput, toInputDate } from '../../utils/dates';

const ACTION_LABEL: Record<ApplicationStatus, string> = {
  planned: '申込予定にする',
  applied: '申込済にする',
  won: '🎉 当選を記録',
  lost: '落選を記録',
  paid: '入金済にする',
  ticketed: '発券済にする',
  expired: '期限切れにする',
};

export default function ApplicationDetailScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const router = useRouter();
  const { applications, patchApplication, removeApplication } = useAppStore();
  const app = applications.find((a) => a.id === id);

  const [payDate, setPayDate] = useState('');
  const [payTime, setPayTime] = useState('23:59');
  const [amount, setAmount] = useState('');
  const [seat, setSeat] = useState('');

  if (!app) {
    return (
      <View style={[styles.container, { justifyContent: 'center', alignItems: 'center' }]}>
        <Text>申込が見つかりません</Text>
      </View>
    );
  }

  const status = app.status as ApplicationStatus;
  const template = BUILTIN_TEMPLATES.find((t) => t.id === app.templateId);

  const transition = (to: ApplicationStatus) => {
    if (to === 'won') {
      // 当選記録: 入金締切の初期値をテンプレ(発表+N日)から提案
      const base = app.announceAt ? new Date(app.announceAt) : new Date();
      const suggested = addDays(base, template?.paymentWindowDays ?? 3);
      setPayDate(toInputDate(suggested));
      patchApplication(app.id, { status: 'won', paymentDeadline: suggested });
      Alert.alert('🎉 当選おめでとうございます!', '入金締切を確認・修正してください。前日と当日朝に通知します');
      return;
    }
    patchApplication(app.id, { status: to });
  };

  const savePaymentDeadline = () => {
    const d = parseDateInput(payDate, payTime);
    if (!d) {
      Alert.alert('入力エラー', '入金締切を YYYY-MM-DD 形式で入力してください');
      return;
    }
    patchApplication(app.id, { paymentDeadline: d });
    Alert.alert('保存しました', '入金締切の前日20時と当日朝9時に通知します');
  };

  const saveWonDetails = () => {
    patchApplication(app.id, {
      seatInfo: seat.trim() || null,
      totalAmount: amount ? Number(amount) : null,
    });
  };

  const confirmDelete = () => {
    Alert.alert('申込を削除', 'この申込を削除しますか?', [
      { text: 'キャンセル', style: 'cancel' },
      {
        text: '削除',
        style: 'destructive',
        onPress: () => {
          removeApplication(app.id);
          router.back();
        },
      },
    ]);
  };

  return (
    <ScrollView style={styles.container} contentContainerStyle={{ padding: 16, paddingBottom: 48 }}>
      <View style={styles.header}>
        <View style={[styles.colorBar, { backgroundColor: app.oshiColor }]} />
        <View style={{ flex: 1 }}>
          <Text style={styles.title}>{app.eventTitle}</Text>
          <Text style={styles.sub}>
            {app.oshiName} ・ {fmtDate(app.performanceDate)} ・ {app.templateName}
            {app.meigiLabel ? ` ・ ${app.meigiLabel}名義` : ''} ・ {app.quantity}枚
          </Text>
        </View>
        <View style={[styles.badge, { backgroundColor: STATUS_COLOR[status] }]}>
          <Text style={styles.badgeText}>{STATUS_LABEL[status]}</Text>
        </View>
      </View>

      <View style={styles.section}>
        <Text style={styles.sectionTitle}>期限</Text>
        <Text style={styles.line}>申込締切: {fmtDateTime(app.closeAt ? new Date(app.closeAt) : null)}</Text>
        <Text style={styles.line}>当落発表: {fmtDateTime(app.announceAt ? new Date(app.announceAt) : null)}</Text>
        <Text style={[styles.line, status === 'won' && { fontWeight: '700', color: '#F2545B' }]}>
          入金締切: {fmtDateTime(app.paymentDeadline ? new Date(app.paymentDeadline) : null)}
        </Text>
      </View>

      {status === 'won' ? (
        <View style={styles.section}>
          <Text style={styles.sectionTitle}>入金締切の修正</Text>
          <View style={{ flexDirection: 'row' }}>
            <View style={{ flex: 2 }}>
              <Field label="日付 (YYYY-MM-DD)" value={payDate} onChangeText={setPayDate} placeholder="2026-08-20" keyboardType="numbers-and-punctuation" />
            </View>
            <View style={{ flex: 1, marginLeft: 10 }}>
              <Field label="時刻" value={payTime} onChangeText={setPayTime} placeholder="23:59" keyboardType="numbers-and-punctuation" />
            </View>
          </View>
          <Button title="入金締切を保存" variant="outline" onPress={savePaymentDeadline} />
          <View style={{ height: 12 }} />
          <Field label="チケット代(合計)" value={amount} onChangeText={setAmount} placeholder="例: 19800" keyboardType="number-pad" />
          <Field label="座席・整理番号" value={seat} onChangeText={setSeat} placeholder="当選メールに記載があれば" />
          <Button title="当選情報を保存" variant="outline" onPress={saveWonDetails} />
        </View>
      ) : null}

      <View style={styles.section}>
        <Text style={styles.sectionTitle}>ステータスを更新</Text>
        {nextStatuses(status).map((to) => (
          <View key={to} style={{ marginBottom: 8 }}>
            <Button
              title={ACTION_LABEL[to]}
              color={to === 'won' ? '#F2545B' : to === 'lost' || to === 'expired' ? '#8E8E93' : '#34C759'}
              onPress={() => transition(to)}
            />
          </View>
        ))}
        {nextStatuses(status).length === 0 ? (
          <Text style={styles.sub}>このステータスから先の操作はありません</Text>
        ) : null}
      </View>

      <Button title="この申込を削除" variant="outline" color="#FF3B30" onPress={confirmDelete} />
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#F6F6F8' },
  header: {
    flexDirection: 'row', alignItems: 'center', gap: 12,
    backgroundColor: '#fff', borderRadius: 14, padding: 14, marginBottom: 12,
  },
  colorBar: { width: 5, alignSelf: 'stretch', borderRadius: 3 },
  title: { fontSize: 16, fontWeight: '700', color: '#1C1C1E' },
  sub: { fontSize: 12, color: '#8E8E93', marginTop: 3, lineHeight: 17 },
  badge: { borderRadius: 8, paddingHorizontal: 8, paddingVertical: 4 },
  badgeText: { color: '#fff', fontSize: 11, fontWeight: '700' },
  section: { backgroundColor: '#fff', borderRadius: 14, padding: 14, marginBottom: 12 },
  sectionTitle: { fontSize: 14, fontWeight: '700', marginBottom: 8, color: '#1C1C1E' },
  line: { fontSize: 14, color: '#3A3A3C', marginBottom: 4 },
});
