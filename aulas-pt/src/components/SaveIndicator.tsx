import React from 'react';
import { StyleSheet, Text } from 'react-native';
import { colors } from '../theme';

type Status = 'idle' | 'saving' | 'saved' | 'error';

export function SaveIndicator({ status }: { status: Status }) {
  if (status === 'idle') return null;
  const label =
    status === 'saving'
      ? 'A guardar…'
      : status === 'saved'
        ? 'Guardado'
        : 'Erro ao guardar';
  return (
    <Text
      style={[
        styles.text,
        status === 'error' && styles.error,
        status === 'saved' && styles.saved,
      ]}
    >
      {label}
    </Text>
  );
}

const styles = StyleSheet.create({
  text: {
    fontSize: 12,
    color: colors.inkMuted,
    fontWeight: '600',
  },
  saved: {
    color: colors.teal,
  },
  error: {
    color: colors.danger,
  },
});
