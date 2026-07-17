import React from 'react';
import { View } from 'react-native';
import { Empty } from '../../components/ui';

// M2で実装: 支出記録・月次予算・推し別集計
export default function BudgetScreen() {
  return (
    <View style={{ flex: 1, backgroundColor: '#F6F6F8', justifyContent: 'center' }}>
      <Empty title="収支は準備中" hint="次のアップデートで支出記録と予算管理ができるようになります" />
    </View>
  );
}
