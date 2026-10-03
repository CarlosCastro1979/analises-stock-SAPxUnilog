import React, { useEffect, useState } from 'react';
import {
  Alert,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  View,
} from 'react-native';
import { router } from 'expo-router';
import { SaveIndicator } from '../components/SaveIndicator';
import { useHousehold } from '../context/HouseholdContext';
import { formatEuros, parseEurosToCents } from '../lib/format';
import { colors, radii, spacing } from '../theme';

export function SettingsScreen() {
  const { householdCode, leaveHousehold, data } = useHousehold();
  const { settings, scheduleSettingsSave, saveStatus, loading } = data;

  const [nameA, setNameA] = useState('');
  const [nameB, setNameB] = useState('');
  const [priceText, setPriceText] = useState('');

  useEffect(() => {
    if (!settings) return;
    setNameA(settings.name_a);
    setNameB(settings.name_b);
    setPriceText((settings.price_cents / 100).toFixed(2).replace('.', ','));
  }, [settings?.household_code]); // only seed when household loads

  // Keep fields in sync if the other phone updates names/price
  useEffect(() => {
    if (!settings) return;
    // Avoid clobbering while user is mid-edit of the same field: only sync when
    // remote values differ and local save is idle.
    if (saveStatus === 'saving') return;
    setNameA((prev) => (prev === settings.name_a ? prev : settings.name_a));
    setNameB((prev) => (prev === settings.name_b ? prev : settings.name_b));
    const remotePrice = (settings.price_cents / 100).toFixed(2).replace('.', ',');
    setPriceText((prev) => {
      const localCents = parseEurosToCents(prev);
      if (localCents === settings.price_cents) return prev;
      return remotePrice;
    });
  }, [settings?.name_a, settings?.name_b, settings?.price_cents, saveStatus]);

  const onNameA = (v: string) => {
    setNameA(v);
    scheduleSettingsSave({ name_a: v.trim() || 'Pessoa 1' });
  };
  const onNameB = (v: string) => {
    setNameB(v);
    scheduleSettingsSave({ name_b: v.trim() || 'Pessoa 2' });
  };
  const onPrice = (v: string) => {
    setPriceText(v);
    const cents = parseEurosToCents(v);
    if (cents != null) {
      scheduleSettingsSave({ price_cents: cents });
    }
  };

  const confirmLeave = () => {
    Alert.alert(
      'Mudar de casa?',
      'Vais precisar de voltar a introduzir o código da casa neste telemóvel. Os dados no servidor mantêm-se.',
      [
        { text: 'Cancelar', style: 'cancel' },
        {
          text: 'Sair',
          style: 'destructive',
          onPress: () => {
            void leaveHousehold().then(() => router.replace('/'));
          },
        },
      ],
    );
  };

  return (
    <ScrollView
      contentContainerStyle={styles.scroll}
      keyboardShouldPersistTaps="handled"
    >
      <View style={styles.rowBetween}>
        <Text style={styles.section}>Nomes e preço</Text>
        <SaveIndicator status={saveStatus} />
      </View>
      <Text style={styles.hint}>
        Grava sozinho cerca de 1 segundo depois de parares de escrever.
      </Text>

      <Field label="Pessoa 1" value={nameA} onChangeText={onNameA} editable={!loading} />
      <Field label="Pessoa 2" value={nameB} onChangeText={onNameB} editable={!loading} />
      <Field
        label="Preço por aula (€)"
        value={priceText}
        onChangeText={onPrice}
        keyboardType="decimal-pad"
        editable={!loading}
      />
      <Text style={styles.pricePreview}>
        Guardado: {formatEuros(settings?.price_cents ?? 0)} · total do mês = aulas × preço
      </Text>

      <View style={styles.divider} />

      <Text style={styles.section}>Casa</Text>
      <Text style={styles.codeLabel}>Código da casa</Text>
      <Text style={styles.codeValue}>{householdCode}</Text>
      <Text style={styles.hint}>
        Os dois telemóveis usam o mesmo código para ver os mesmos dados.
      </Text>

      <Pressable
        onPress={confirmLeave}
        style={({ pressed }) => [styles.leaveBtn, pressed && { opacity: 0.85 }]}
      >
        <Text style={styles.leaveText}>Usar outro código</Text>
      </Pressable>
    </ScrollView>
  );
}

function Field({
  label,
  value,
  onChangeText,
  keyboardType,
  editable = true,
}: {
  label: string;
  value: string;
  onChangeText: (v: string) => void;
  keyboardType?: 'default' | 'decimal-pad';
  editable?: boolean;
}) {
  return (
    <View style={styles.field}>
      <Text style={styles.fieldLabel}>{label}</Text>
      <TextInput
        value={value}
        onChangeText={onChangeText}
        style={styles.input}
        editable={editable}
        keyboardType={keyboardType}
        autoCorrect={false}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  scroll: {
    padding: spacing.lg,
    gap: spacing.md,
    backgroundColor: colors.surface,
  },
  rowBetween: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  section: {
    fontSize: 18,
    fontWeight: '800',
    color: colors.ink,
  },
  hint: {
    fontSize: 13,
    color: colors.inkMuted,
    lineHeight: 18,
    marginTop: -6,
  },
  field: { gap: 6 },
  fieldLabel: {
    fontSize: 13,
    fontWeight: '600',
    color: colors.inkMuted,
  },
  input: {
    backgroundColor: colors.surfaceCard,
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: radii.sm,
    paddingHorizontal: spacing.md,
    paddingVertical: 12,
    fontSize: 16,
    color: colors.ink,
  },
  pricePreview: {
    fontSize: 13,
    color: colors.teal,
    fontWeight: '600',
  },
  divider: {
    height: 1,
    backgroundColor: colors.border,
    marginVertical: spacing.sm,
  },
  codeLabel: {
    fontSize: 13,
    fontWeight: '600',
    color: colors.inkMuted,
  },
  codeValue: {
    fontSize: 20,
    fontWeight: '700',
    color: colors.ink,
    marginTop: -8,
  },
  leaveBtn: {
    marginTop: spacing.sm,
    borderWidth: 1,
    borderColor: colors.danger,
    borderRadius: radii.sm,
    paddingVertical: 14,
    alignItems: 'center',
  },
  leaveText: {
    color: colors.danger,
    fontWeight: '700',
    fontSize: 15,
  },
});
