import React, { useMemo } from 'react';
import { Pressable, StyleSheet, Text, View } from 'react-native';
import { buildMonthGrid, toDayKey, weekdayLabels } from '../lib/format';
import type { Person } from '../lib/types';
import { colors, radii, spacing } from '../theme';

type Props = {
  year: number;
  monthIndex: number;
  person: Person;
  sessionSet: Set<string>;
  onToggle: (day: number) => void;
};

export function MonthCalendar({
  year,
  monthIndex,
  person,
  sessionSet,
  onToggle,
}: Props) {
  const weeks = useMemo(
    () => buildMonthGrid(year, monthIndex),
    [year, monthIndex],
  );
  const labels = weekdayLabels();
  const today = new Date();
  const isCurrentMonth =
    today.getFullYear() === year && today.getMonth() === monthIndex;
  const todayDay = today.getDate();

  return (
    <View style={styles.wrap}>
      <View style={styles.weekRow}>
        {labels.map((label) => (
          <Text key={label} style={styles.weekday}>
            {label}
          </Text>
        ))}
      </View>
      {weeks.map((week, wi) => (
        <View key={wi} style={styles.weekRow}>
          {week.map((day, di) => {
            if (day == null) {
              return <View key={`e-${wi}-${di}`} style={styles.cell} />;
            }
            const key = `${person}:${toDayKey(year, monthIndex, day)}`;
            const done = sessionSet.has(key);
            const isToday = isCurrentMonth && day === todayDay;
            return (
              <Pressable
                key={day}
                onPress={() => onToggle(day)}
                style={({ pressed }) => [
                  styles.cell,
                  styles.day,
                  done && styles.dayDone,
                  isToday && !done && styles.dayToday,
                  pressed && styles.dayPressed,
                ]}
                accessibilityRole="button"
                accessibilityLabel={
                  done
                    ? `Dia ${day}, horário feito. Tocar para desmarcar.`
                    : `Dia ${day}. Tocar para marcar horário feito.`
                }
              >
                <Text style={[styles.dayText, done && styles.dayTextDone]}>
                  {day}
                </Text>
              </Pressable>
            );
          })}
        </View>
      ))}
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
  },
  weekRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    marginBottom: 6,
  },
  weekday: {
    flex: 1,
    textAlign: 'center',
    fontSize: 12,
    fontWeight: '600',
    color: colors.inkMuted,
    letterSpacing: 0.3,
  },
  cell: {
    flex: 1,
    aspectRatio: 1,
    alignItems: 'center',
    justifyContent: 'center',
    padding: 2,
  },
  day: {
    borderRadius: radii.sm,
    backgroundColor: colors.surface,
  },
  dayDone: {
    backgroundColor: colors.done,
  },
  dayToday: {
    borderWidth: 2,
    borderColor: colors.accent,
  },
  dayPressed: {
    opacity: 0.85,
    transform: [{ scale: 0.96 }],
  },
  dayText: {
    fontSize: 16,
    fontWeight: '600',
    color: colors.ink,
  },
  dayTextDone: {
    color: colors.doneText,
  },
});
