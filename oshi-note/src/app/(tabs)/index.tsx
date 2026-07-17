import { useRouter } from 'expo-router';
import React, { useMemo } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { Button, Empty } from '../../components/ui';
import { ApplicationView } from '../../db/repos';
import { ApplicationStatus, STATUS_COLOR, STATUS_LABEL } from '../../domain/types';
import { useAppStore } from '../../state/store';
import { daysUntil, fmtDate, fmtDateTime } from '../../utils/dates';

interface Deadline {
  app: ApplicationView;
  label: string;
  date: Date;
  urgent: boolean;
}

// 申込ごとの「次の期限」を1つに絞る
function nextDeadline(a: ApplicationView, now: Date): Deadline | null {
  const status = a.status as ApplicationStatus;
  if (status === 'planned' && a.closeAt) {
    const d = new Date(a.closeAt);
    return { app: a, label: '申込締切', date: d, urgent: daysUntil(d, now) <= 1 };
  }
  if (status === 'applied' && a.announceAt) {
    const d = new Date(a.announceAt);
    return { app: a, label: '当落発表', date: d, urgent: daysUntil(d, now) <= 0 };
  }
  if (status === 'won' && a.paymentDeadline) {
    const d = new Date(a.paymentDeadline);
    return { app: a, label: '入金締切', date: d, urgent: daysUntil(d, now) <= 1 };
  }
  return null;
}

export default function HomeScreen() {
  const router = useRouter();
  const { applications, oshis } = useAppStore();

  const deadlines = useMemo(() => {
    const now = new Date();
    return applications
      .map((a) => nextDeadline(a, now))
      .filter((x): x is Deadline => !!x && x.date.getTime() > now.getTime() - 86400000)
      .sort((x, y) => x.date.getTime() - y.date.getTime());
  }, [applications]);

  return (
    <ScrollView style={styles.container} contentContainerStyle={{ padding: 16 }}>
      <Text style={styles.sectionTitle}>次の期限</Text>
      {deadlines.length === 0 ? (
        <Empty
          title="期限はすべてクリア!"
          hint={
            oshis.length === 0
              ? 'まずは「推し」タブから推しを登録して、申込を追加しましょう'
              : '「当落」タブから申込を登録すると、締切や当落発表の通知が自動でセットされます'
          }
        />
      ) : (
        deadlines.map((d) => {
          const days = daysUntil(d.date);
          return (
            <Pressable
              key={d.app.id}
              style={[styles.card, d.urgent && styles.cardUrgent]}
              onPress={() => router.push({ pathname: '/application/[id]', params: { id: d.app.id } })}
            >
              <View style={[styles.colorBar, { backgroundColor: d.app.oshiColor }]} />
              <View style={{ flex: 1 }}>
                <Text style={styles.cardTitle}>{d.app.eventTitle}</Text>
                <Text style={styles.cardSub}>
                  {d.app.oshiName}
                  {d.app.meigiLabel ? ` ・ ${d.app.meigiLabel}名義` : ''} ・ {d.app.templateName}
                </Text>
                <Text style={[styles.deadline, d.urgent && { color: '#FF3B30' }]}>
                  {d.label}: {fmtDateTime(d.date)}
                  {days === 0 ? ' (今日!)' : days === 1 ? ' (明日)' : days > 1 ? ` (あと${days}日)` : ''}
                </Text>
              </View>
              <View style={[styles.badge, { backgroundColor: STATUS_COLOR[d.app.status as ApplicationStatus] }]}>
                <Text style={styles.badgeText}>{STATUS_LABEL[d.app.status as ApplicationStatus]}</Text>
              </View>
            </Pressable>
          );
        })
      )}

      <Text style={styles.sectionTitle}>参戦予定</Text>
      {applications.filter((a) => ['won', 'paid', 'ticketed'].includes(a.status) && a.performanceDate).length === 0 ? (
        <Empty title="参戦予定はまだありません" hint="当選を記録するとここに表示されます" />
      ) : (
        applications
          .filter((a) => ['won', 'paid', 'ticketed'].includes(a.status) && a.performanceDate)
          .sort((a, b) => (a.performanceDate as Date).getTime() - (b.performanceDate as Date).getTime())
          .map((a) => (
            <Pressable
              key={a.id}
              style={styles.card}
              onPress={() => router.push({ pathname: '/application/[id]', params: { id: a.id } })}
            >
              <View style={[styles.colorBar, { backgroundColor: a.oshiColor }]} />
              <View style={{ flex: 1 }}>
                <Text style={styles.cardTitle}>{a.eventTitle}</Text>
                <Text style={styles.cardSub}>{fmtDate(a.performanceDate)}</Text>
              </View>
            </Pressable>
          ))
      )}

      <View style={{ height: 12 }} />
      <Button title="＋ 申込を登録する" onPress={() => router.push('/application-form')} />
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#F6F6F8' },
  sectionTitle: { fontSize: 16, fontWeight: '700', marginTop: 8, marginBottom: 10, color: '#1C1C1E' },
  card: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: '#fff',
    borderRadius: 14,
    padding: 14,
    marginBottom: 10,
    gap: 12,
  },
  cardUrgent: { borderWidth: 1.5, borderColor: '#FF3B30' },
  colorBar: { width: 5, alignSelf: 'stretch', borderRadius: 3 },
  cardTitle: { fontSize: 15, fontWeight: '600', color: '#1C1C1E' },
  cardSub: { fontSize: 12, color: '#8E8E93', marginTop: 2 },
  deadline: { fontSize: 13, fontWeight: '600', color: '#3A3A3C', marginTop: 6 },
  badge: { borderRadius: 8, paddingHorizontal: 8, paddingVertical: 4 },
  badgeText: { color: '#fff', fontSize: 11, fontWeight: '700' },
});
