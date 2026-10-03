import React from 'react';
import { StyleSheet, Text, View } from 'react-native';
import { formatEuros } from '../lib/format';
import { colors, radii, spacing } from '../theme';

type Props = {
  nameA: string;
  nameB: string;
  daysA: number[];
  daysB: number[];
  priceCents: number;
};

export function MonthSummary({
  nameA,
  nameB,
  daysA,
  daysB,
  priceCents,
}: Props) {
  const total = daysA.length + daysB.length;
  const amount = total * priceCents;

  return (
    <View style={styles.wrap}>
      <Text style={styles.title}>Resumo do mês</Text>
      <PersonLine name={nameA} days={daysA} />
      <PersonLine name={nameB} days={daysB} />
      <View style={styles.totalRow}>
        <Text style={styles.totalLabel}>
          {total} aulas × {formatEuros(priceCents)}
        </Text>
        <Text style={styles.totalValue}>{formatEuros(amount)}</Text>
      </View>
    </View>
  );
}

function PersonLine({ name, days }: { name: string; days: number[] }) {
  const list = days.length ? days.join(', ') : '—';
  return (
    <View style={styles.personBlock}>
      <Text style={styles.personName}>
        {name}: {days.length} {days.length === 1 ? 'aula' : 'aulas'}
      </Text>
      <Text style={styles.personDays}>({list})</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  wrap: {
    backgroundColor: colors.surfaceCard,
    borderRadius: radii.lg,
    padding: spacing.md,
    borderWidth: 1,
    borderColor: colors.border,
    gap: spacing.sm,
  },
  title: {
    fontSize: 15,
    fontWeight: '700',
    color: colors.ink,
    marginBottom: 2,
  },
  personBlock: {
    gap: 2,
  },
  personName: {
    fontSize: 15,
    fontWeight: '600',
    color: colors.ink,
  },
  personDays: {
    fontSize: 13,
    color: colors.inkMuted,
  },
  totalRow: {
    marginTop: spacing.sm,
    paddingTop: spacing.sm,
    borderTopWidth: 1,
    borderTopColor: colors.border,
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: spacing.sm,
  },
  totalLabel: {
    flex: 1,
    fontSize: 14,
    color: colors.inkMuted,
  },
  totalValue: {
    fontSize: 18,
    fontWeight: '800',
    color: colors.teal,
  },
});
