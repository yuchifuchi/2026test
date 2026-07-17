import { useRouter } from 'expo-router';
import React, { useMemo, useState } from 'react';
import { Pressable, ScrollView, StyleSheet, Text, View } from 'react-native';
import { Button, Chip, Empty } from '../../components/ui';
import { ApplicationStatus, STATUS_COLOR, STATUS_LABEL } from '../../domain/types';
import { useAppStore } from '../../state/store';
import { fmtDate, fmtDateTime } from '../../utils/dates';

type Filter = 'active' | 'done' | 'all';

const ACTIVE: ApplicationStatus[] = ['planned', 'applied', 'won'];

export default function TourakuScreen() {
  const router = useRouter();
  const applications = useAppStore((s) => s.applications);
  const [filter, setFilter] = useState<Filter>('active');

  const list = useMemo(() => {
    const filtered = applications.filter((a) => {
      const st = a.status as ApplicationStatus;
      if (filter === 'active') return ACTIVE.includes(st);
      if (filter === 'done') return !ACTIVE.includes(st);
      return true;
    });
    return filtered.sort((a, b) => b.createdAt - a.createdAt);
  }, [applications, filter]);

  return (
    <View style={styles.container}>
      <View style={styles.filterRow}>
        <Chip label="進行中" selected={filter === 'active'} onPress={() => setFilter('active')} />
        <Chip label="完了" selected={filter === 'done'} onPress={() => setFilter('done')} />
        <Chip label="すべて" selected={filter === 'all'} onPress={() => setFilter('all')} />
      </View>
      <ScrollView contentContainerStyle={{ padding: 16, paddingTop: 4 }}>
        {list.length === 0 ? (
          <Empty
            title="申込がありません"
            hint="申込を登録すると、締切・当落発表・入金期限の通知が自動でセットされます"
          />
        ) : (
          list.map((a) => {
            const st = a.status as ApplicationStatus;
            const deadline =
              st === 'planned' && a.closeAt
                ? `申込締切 ${fmtDateTime(new Date(a.closeAt))}`
                : st === 'applied' && a.announceAt
                  ? `当落発表 ${fmtDateTime(new Date(a.announceAt))}`
                  : st === 'won' && a.paymentDeadline
                    ? `入金締切 ${fmtDateTime(new Date(a.paymentDeadline))}`
                    : null;
            return (
              <Pressable
                key={a.id}
                style={styles.card}
                onPress={() => router.push({ pathname: '/application/[id]', params: { id: a.id } })}
              >
                <View style={[styles.colorBar, { backgroundColor: a.oshiColor }]} />
                <View style={{ flex: 1 }}>
                  <Text style={styles.cardTitle}>{a.eventTitle}</Text>
                  <Text style={styles.cardSub}>
                    {fmtDate(a.performanceDate)} ・ {a.templateName}
                    {a.meigiLabel ? ` ・ ${a.meigiLabel}名義` : ''} ・ {a.quantity}枚
                  </Text>
                  {deadline ? <Text style={styles.deadline}>{deadline}</Text> : null}
                </View>
                <View style={[styles.badge, { backgroundColor: STATUS_COLOR[st] }]}>
                  <Text style={styles.badgeText}>{STATUS_LABEL[st]}</Text>
                </View>
              </Pressable>
            );
          })
        )}
        <View style={{ height: 12 }} />
        <Button title="＋ 申込を登録する" onPress={() => router.push('/application-form')} />
      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#F6F6F8' },
  filterRow: { flexDirection: 'row', paddingHorizontal: 16, paddingTop: 12 },
  card: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: '#fff',
    borderRadius: 14,
    padding: 14,
    marginBottom: 10,
    gap: 12,
  },
  colorBar: { width: 5, alignSelf: 'stretch', borderRadius: 3 },
  cardTitle: { fontSize: 15, fontWeight: '600', color: '#1C1C1E' },
  cardSub: { fontSize: 12, color: '#8E8E93', marginTop: 2 },
  deadline: { fontSize: 13, fontWeight: '600', color: '#3A3A3C', marginTop: 6 },
  badge: { borderRadius: 8, paddingHorizontal: 8, paddingVertical: 4 },
  badgeText: { color: '#fff', fontSize: 11, fontWeight: '700' },
});
