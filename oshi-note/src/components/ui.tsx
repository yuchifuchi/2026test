import React from 'react';
import {
  Pressable, StyleSheet, Text, TextInput, TextInputProps, View, ViewStyle,
} from 'react-native';

export function Button(props: {
  title: string;
  onPress: () => void;
  color?: string;
  variant?: 'solid' | 'outline';
  style?: ViewStyle;
}) {
  const color = props.color ?? '#F2545B';
  const solid = (props.variant ?? 'solid') === 'solid';
  return (
    <Pressable
      onPress={props.onPress}
      style={({ pressed }) => [
        styles.button,
        solid ? { backgroundColor: color } : { borderWidth: 1.5, borderColor: color },
        pressed && { opacity: 0.7 },
        props.style,
      ]}
    >
      <Text style={[styles.buttonText, { color: solid ? '#fff' : color }]}>{props.title}</Text>
    </Pressable>
  );
}

export function Field(props: TextInputProps & { label: string }) {
  const { label, ...rest } = props;
  return (
    <View style={styles.field}>
      <Text style={styles.label}>{label}</Text>
      <TextInput
        placeholderTextColor="#B0B0B6"
        {...rest}
        style={[styles.input, rest.style]}
      />
    </View>
  );
}

export function Chip(props: {
  label: string;
  selected: boolean;
  onPress: () => void;
  color?: string;
}) {
  const color = props.color ?? '#F2545B';
  return (
    <Pressable
      onPress={props.onPress}
      style={[
        styles.chip,
        props.selected ? { backgroundColor: color, borderColor: color } : { borderColor: '#D0D0D6' },
      ]}
    >
      <Text style={{ color: props.selected ? '#fff' : '#444', fontSize: 13 }}>{props.label}</Text>
    </Pressable>
  );
}

export function Empty(props: { title: string; hint?: string }) {
  return (
    <View style={styles.empty}>
      <Text style={styles.emptyTitle}>{props.title}</Text>
      {props.hint ? <Text style={styles.emptyHint}>{props.hint}</Text> : null}
    </View>
  );
}

const styles = StyleSheet.create({
  button: {
    borderRadius: 12,
    paddingVertical: 13,
    paddingHorizontal: 18,
    alignItems: 'center',
  },
  buttonText: { fontSize: 15, fontWeight: '600' },
  field: { marginBottom: 14 },
  label: { fontSize: 12, color: '#6B6B70', marginBottom: 5, fontWeight: '600' },
  input: {
    borderWidth: 1,
    borderColor: '#D8D8DE',
    borderRadius: 10,
    paddingHorizontal: 12,
    paddingVertical: 10,
    fontSize: 15,
    backgroundColor: '#fff',
    color: '#1C1C1E',
  },
  chip: {
    borderWidth: 1.5,
    borderRadius: 16,
    paddingHorizontal: 12,
    paddingVertical: 6,
    marginRight: 8,
    marginBottom: 8,
  },
  empty: { alignItems: 'center', paddingVertical: 48, paddingHorizontal: 24 },
  emptyTitle: { fontSize: 15, fontWeight: '600', color: '#6B6B70', marginBottom: 6 },
  emptyHint: { fontSize: 13, color: '#9A9AA0', textAlign: 'center', lineHeight: 19 },
});
