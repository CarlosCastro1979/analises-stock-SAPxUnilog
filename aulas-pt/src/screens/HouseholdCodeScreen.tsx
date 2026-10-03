import React, { useState } from 'react';
import {
  ActivityIndicator,
  KeyboardAvoidingView,
  Platform,
  Pressable,
  StyleSheet,
  Text,
  TextInput,
  View,
} from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import { LinearGradientFallback } from '../components/LinearGradientFallback';
import { useHousehold } from '../context/HouseholdContext';
import { colors, radii, spacing } from '../theme';

export function HouseholdCodeScreen() {
  const { joinHousehold } = useHousehold();
  const [code, setCode] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onContinue = async () => {
    setBusy(true);
    setError(null);
    try {
      await joinHousehold(code);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Não foi possível entrar.');
    } finally {
      setBusy(false);
    }
  };

  return (
    <LinearGradientFallback style={styles.bg}>
      <SafeAreaView style={styles.safe}>
        <KeyboardAvoidingView
          behavior={Platform.OS === 'ios' ? 'padding' : undefined}
          style={styles.kav}
        >
          <View style={styles.hero}>
            <Text style={styles.brand}>Aulas PT</Text>
            <Text style={styles.tagline}>
              Marca os horários feitos. Os dois vêem o mesmo calendário.
            </Text>
          </View>

          <View style={styles.card}>
            <Text style={styles.label}>Código da casa</Text>
            <Text style={styles.hint}>
              Escolham a mesma palavra nos dois telemóveis. Sem conta, sem link.
            </Text>
            <TextInput
              value={code}
              onChangeText={setCode}
              placeholder="ex.: casa-silva"
              placeholderTextColor={colors.inkMuted}
              autoCapitalize="none"
              autoCorrect={false}
              style={styles.input}
              onSubmitEditing={() => void onContinue()}
              returnKeyType="go"
            />
            {error ? <Text style={styles.error}>{error}</Text> : null}
            <Pressable
              onPress={() => void onContinue()}
              disabled={busy || !code.trim()}
              style={({ pressed }) => [
                styles.btn,
                (!code.trim() || busy) && styles.btnDisabled,
                pressed && styles.btnPressed,
              ]}
            >
              {busy ? (
                <ActivityIndicator color="#fff" />
              ) : (
                <Text style={styles.btnText}>Entrar</Text>
              )}
            </Pressable>
          </View>
        </KeyboardAvoidingView>
      </SafeAreaView>
    </LinearGradientFallback>
  );
}

const styles = StyleSheet.create({
  bg: { flex: 1 },
  safe: { flex: 1 },
  kav: {
    flex: 1,
    justifyContent: 'center',
    paddingHorizontal: spacing.lg,
    gap: spacing.xl,
  },
  hero: { gap: spacing.sm },
  brand: {
    fontSize: 48,
    fontWeight: '800',
    color: '#fff',
    letterSpacing: -1,
  },
  tagline: {
    fontSize: 17,
    lineHeight: 24,
    color: 'rgba(255,255,255,0.88)',
    maxWidth: 320,
  },
  card: {
    backgroundColor: colors.surfaceCard,
    borderRadius: radii.lg,
    padding: spacing.lg,
    gap: spacing.sm,
  },
  label: {
    fontSize: 16,
    fontWeight: '700',
    color: colors.ink,
  },
  hint: {
    fontSize: 14,
    color: colors.inkMuted,
    marginBottom: spacing.xs,
    lineHeight: 20,
  },
  input: {
    borderWidth: 1,
    borderColor: colors.border,
    borderRadius: radii.sm,
    paddingHorizontal: spacing.md,
    paddingVertical: 14,
    fontSize: 17,
    color: colors.ink,
    backgroundColor: colors.surface,
  },
  error: {
    color: colors.danger,
    fontSize: 13,
  },
  btn: {
    marginTop: spacing.sm,
    backgroundColor: colors.accent,
    borderRadius: radii.sm,
    paddingVertical: 14,
    alignItems: 'center',
  },
  btnDisabled: { opacity: 0.5 },
  btnPressed: { transform: [{ scale: 0.98 }] },
  btnText: {
    color: '#fff',
    fontSize: 17,
    fontWeight: '700',
  },
});
