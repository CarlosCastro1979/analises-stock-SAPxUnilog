import { Link } from 'expo-router';
import React, { useMemo, useState } from 'react';
import {
  ActivityIndicator,
  Linking,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  View,
} from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import { MonthCalendar } from '../components/MonthCalendar';
import { MonthSummary } from '../components/MonthSummary';
import { useHousehold } from '../context/HouseholdContext';
import {
  buildWhatsAppText,
  monthTitle,
  personName,
} from '../lib/format';
import type { Person } from '../lib/types';
import { colors, radii, spacing } from '../theme';

export function HomeScreen() {
  const { data } = useHousehold();
  const {
    settings,
    loading,
    error,
    sessionSet,
    toggleSession,
    daysForPersonMonth,
  } = data;

  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [monthIndex, setMonthIndex] = useState(now.getMonth());
  const [person, setPerson] = useState<Person>('a');

  const daysA = useMemo(
    () => daysForPersonMonth('a', year, monthIndex),
    [daysForPersonMonth, year, monthIndex],
  );
  const daysB = useMemo(
    () => daysForPersonMonth('b', year, monthIndex),
    [daysForPersonMonth, year, monthIndex],
  );

  const nameA = personName(settings, 'a');
  const nameB = personName(settings, 'b');
  const priceCents = settings?.price_cents ?? 2500;

  const shiftMonth = (delta: number) => {
    const d = new Date(year, monthIndex + delta, 1);
    setYear(d.getFullYear());
    setMonthIndex(d.getMonth());
  };

  const openWhatsApp = async () => {
    const text = buildWhatsAppText({
      year,
      monthIndex,
      nameA,
      nameB,
      daysA,
      daysB,
      priceCents,
    });
    const url = `https://wa.me/?text=${encodeURIComponent(text)}`;
    await Linking.openURL(url);
  };

  return (
    <View style={styles.root}>
      <SafeAreaView edges={['top']} style={styles.topSafe}>
        <View style={styles.header}>
          <View>
            <Text style={styles.brand}>Aulas PT</Text>
            <Text style={styles.sub}>Horários feitos · preço fixo</Text>
          </View>
          <Link href="/settings" asChild>
            <Pressable style={({ pressed }) => [styles.gear, pressed && { opacity: 0.8 }]}>
              <Text style={styles.gearText}>Definições</Text>
            </Pressable>
          </Link>
        </View>
      </SafeAreaView>

      <ScrollView
        contentContainerStyle={styles.scroll}
        keyboardShouldPersistTaps="handled"
      >
        <View style={styles.monthNav}>
          <Pressable onPress={() => shiftMonth(-1)} style={styles.navBtn}>
            <Text style={styles.navBtnText}>‹</Text>
          </Pressable>
          <Text style={styles.monthLabel}>{monthTitle(year, monthIndex)}</Text>
          <Pressable onPress={() => shiftMonth(1)} style={styles.navBtn}>
            <Text style={styles.navBtnText}>›</Text>
          </Pressable>
        </View>

        <View style={styles.tabs}>
          {(['a', 'b'] as Person[]).map((p) => {
            const active = person === p;
            const label = p === 'a' ? nameA : nameB;
            return (
              <Pressable
                key={p}
                onPress={() => setPerson(p)}
                style={[styles.tab, active && styles.tabActive]}
              >
                <Text style={[styles.tabText, active && styles.tabTextActive]}>
                  {label}
                </Text>
              </Pressable>
            );
          })}
        </View>

        {loading && !settings ? (
          <ActivityIndicator color={colors.teal} style={{ marginVertical: 40 }} />
        ) : (
          <>
            {error ? (
              <Text style={styles.errorBanner}>
                {error.includes('Could not find the table') ||
                error.includes('schema cache')
                  ? 'As tabelas Aulas PT ainda não existem no Supabase. Corre o SQL em aulas-pt/supabase.sql.'
                  : error}
              </Text>
            ) : null}

            <Text style={styles.hint}>
              Toca num dia para marcar ou desmarcar horário feito
            </Text>

            <MonthCalendar
              year={year}
              monthIndex={monthIndex}
              person={person}
              sessionSet={sessionSet}
              onToggle={(day) => void toggleSession(person, year, monthIndex, day)}
            />

            <MonthSummary
              nameA={nameA}
              nameB={nameB}
              daysA={daysA}
              daysB={daysB}
              priceCents={priceCents}
            />

            <Pressable
              onPress={() => void openWhatsApp()}
              style={({ pressed }) => [
                styles.waBtn,
                pressed && { transform: [{ scale: 0.98 }] },
              ]}
            >
              <Text style={styles.waBtnText}>Enviar no WhatsApp</Text>
            </Pressable>
          </>
        )}
      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1, backgroundColor: colors.surface },
  topSafe: { backgroundColor: colors.bg },
  header: {
    backgroundColor: colors.bg,
    paddingHorizontal: spacing.lg,
    paddingBottom: spacing.md,
    paddingTop: spacing.sm,
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'flex-end',
  },
  brand: {
    fontSize: 28,
    fontWeight: '800',
    color: '#fff',
    letterSpacing: -0.5,
  },
  sub: {
    marginTop: 2,
    fontSize: 13,
    color: 'rgba(255,255,255,0.75)',
  },
  gear: {
    backgroundColor: 'rgba(255,255,255,0.14)',
    paddingHorizontal: 12,
    paddingVertical: 8,
    borderRadius: radii.sm,
  },
  gearText: {
    color: '#fff',
    fontWeight: '600',
    fontSize: 13,
  },
  scroll: {
    padding: spacing.md,
    gap: spacing.md,
    paddingBottom: spacing.xl,
  },
  monthNav: {
    flexDirection: 'row',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  navBtn: {
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: colors.surfaceCard,
    borderWidth: 1,
    borderColor: colors.border,
    alignItems: 'center',
    justifyContent: 'center',
  },
  navBtnText: {
    fontSize: 28,
    lineHeight: 30,
    color: colors.teal,
    fontWeight: '600',
  },
  monthLabel: {
    fontSize: 20,
    fontWeight: '700',
    color: colors.ink,
  },
  tabs: {
    flexDirection: 'row',
    backgroundColor: colors.surfaceCard,
    borderRadius: radii.md,
    padding: 4,
    borderWidth: 1,
    borderColor: colors.border,
    gap: 4,
  },
  tab: {
    flex: 1,
    paddingVertical: 12,
    borderRadius: radii.sm,
    alignItems: 'center',
  },
  tabActive: {
    backgroundColor: colors.teal,
  },
  tabText: {
    fontSize: 15,
    fontWeight: '600',
    color: colors.inkMuted,
  },
  tabTextActive: {
    color: '#fff',
  },
  hint: {
    fontSize: 13,
    color: colors.inkMuted,
    textAlign: 'center',
  },
  errorBanner: {
    backgroundColor: '#FDE8E4',
    color: colors.danger,
    padding: spacing.md,
    borderRadius: radii.sm,
    fontSize: 13,
    lineHeight: 18,
  },
  waBtn: {
    backgroundColor: colors.whatsapp,
    borderRadius: radii.md,
    paddingVertical: 16,
    alignItems: 'center',
  },
  waBtnText: {
    color: '#fff',
    fontSize: 17,
    fontWeight: '800',
  },
});
