import React from 'react';
import { View } from 'react-native';
import { Empty } from '../../components/ui';

// M2で実装: 月カレンダー(公演・発表日・締切のプロット)
export default function CalendarScreen() {
  return (
    <View style={{ flex: 1, backgroundColor: '#F6F6F8', justifyContent: 'center' }}>
      <Empty title="カレンダーは準備中" hint="次のアップデートで公演・締切を月表示できるようになります" />
    </View>
  );
}
