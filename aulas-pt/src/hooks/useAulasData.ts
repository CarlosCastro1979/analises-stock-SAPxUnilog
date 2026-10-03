import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { supabase } from '../lib/supabase';
import type { Person, SessionRow, Settings } from '../lib/types';
import { toDayKey } from '../lib/format';

const DEFAULT_PRICE_CENTS = 2500;

type SaveStatus = 'idle' | 'saving' | 'saved' | 'error';

export function useAulasData(householdCode: string | null) {
  const [settings, setSettings] = useState<Settings | null>(null);
  const [sessions, setSessions] = useState<SessionRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [saveStatus, setSaveStatus] = useState<SaveStatus>('idle');
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const savedFlash = useRef<ReturnType<typeof setTimeout> | null>(null);
  const settingsRef = useRef<Settings | null>(null);

  useEffect(() => {
    settingsRef.current = settings;
  }, [settings]);

  const markSaved = useCallback(() => {
    setSaveStatus('saved');
    if (savedFlash.current) clearTimeout(savedFlash.current);
    savedFlash.current = setTimeout(() => setSaveStatus('idle'), 1600);
  }, []);

  const ensureHousehold = useCallback(async (code: string) => {
    const { data, error: selErr } = await supabase
      .from('aulas_pt_settings')
      .select('*')
      .eq('household_code', code)
      .maybeSingle();

    if (selErr) throw selErr;

    if (data) {
      setSettings(data as Settings);
      return data as Settings;
    }

    const insertPayload = {
      household_code: code,
      name_a: 'Pessoa 1',
      name_b: 'Pessoa 2',
      price_cents: DEFAULT_PRICE_CENTS,
      updated_at: new Date().toISOString(),
    };

    const { data: created, error: insErr } = await supabase
      .from('aulas_pt_settings')
      .insert(insertPayload)
      .select('*')
      .single();

    if (insErr) throw insErr;
    setSettings(created as Settings);
    return created as Settings;
  }, []);

  const loadSessions = useCallback(async (code: string) => {
    const { data, error: sessErr } = await supabase
      .from('aulas_pt_sessions')
      .select('household_code, person, day')
      .eq('household_code', code);

    if (sessErr) throw sessErr;
    setSessions((data ?? []) as SessionRow[]);
  }, []);

  const refresh = useCallback(async () => {
    if (!householdCode) {
      setLoading(false);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      await ensureHousehold(householdCode);
      await loadSessions(householdCode);
    } catch (e: unknown) {
      const message =
        e && typeof e === 'object' && 'message' in e
          ? String((e as { message: string }).message)
          : 'Não foi possível carregar os dados.';
      setError(message);
    } finally {
      setLoading(false);
    }
  }, [householdCode, ensureHousehold, loadSessions]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  // Realtime
  useEffect(() => {
    if (!householdCode) return;

    const channel = supabase
      .channel(`aulas-pt-${householdCode}`)
      .on(
        'postgres_changes',
        {
          event: '*',
          schema: 'public',
          table: 'aulas_pt_settings',
          filter: `household_code=eq.${householdCode}`,
        },
        (payload) => {
          if (payload.eventType === 'DELETE') return;
          const next = payload.new as Settings;
          setSettings(next);
        },
      )
      .on(
        'postgres_changes',
        {
          event: '*',
          schema: 'public',
          table: 'aulas_pt_sessions',
          filter: `household_code=eq.${householdCode}`,
        },
        (payload) => {
          if (payload.eventType === 'DELETE') {
            const old = payload.old as SessionRow;
            setSessions((prev) =>
              prev.filter(
                (s) =>
                  !(
                    s.person === old.person &&
                    s.day === old.day
                  ),
              ),
            );
            return;
          }
          const row = payload.new as SessionRow;
          setSessions((prev) => {
            const without = prev.filter(
              (s) => !(s.person === row.person && s.day === row.day),
            );
            return [...without, row];
          });
        },
      )
      .subscribe();

    return () => {
      void supabase.removeChannel(channel);
    };
  }, [householdCode]);

  const persistSettings = useCallback(
    async (patch: Partial<Pick<Settings, 'name_a' | 'name_b' | 'price_cents'>>) => {
      if (!householdCode) return;
      const current = settingsRef.current;
      if (!current) return;

      const next: Settings = {
        ...current,
        ...patch,
        updated_at: new Date().toISOString(),
      };
      setSettings(next);
      setSaveStatus('saving');

      const { error: upErr } = await supabase.from('aulas_pt_settings').upsert({
        household_code: householdCode,
        name_a: next.name_a,
        name_b: next.name_b,
        price_cents: next.price_cents,
        updated_at: next.updated_at,
      });

      if (upErr) {
        setSaveStatus('error');
        setError(upErr.message);
        return;
      }
      markSaved();
    },
    [householdCode, markSaved],
  );

  const scheduleSettingsSave = useCallback(
    (patch: Partial<Pick<Settings, 'name_a' | 'name_b' | 'price_cents'>>) => {
      if (!householdCode) return;
      const current = settingsRef.current;
      if (!current) return;
      const optimistic = { ...current, ...patch };
      setSettings(optimistic);
      settingsRef.current = optimistic;

      if (saveTimer.current) clearTimeout(saveTimer.current);
      saveTimer.current = setTimeout(() => {
        void persistSettings(patch);
      }, 1000);
    },
    [householdCode, persistSettings],
  );

  const toggleSession = useCallback(
    async (person: Person, year: number, monthIndex: number, day: number) => {
      if (!householdCode) return;
      const dayKey = toDayKey(year, monthIndex, day);
      const exists = sessions.some(
        (s) => s.person === person && s.day === dayKey,
      );

      // Optimistic
      if (exists) {
        setSessions((prev) =>
          prev.filter((s) => !(s.person === person && s.day === dayKey)),
        );
        const { error: delErr } = await supabase
          .from('aulas_pt_sessions')
          .delete()
          .eq('household_code', householdCode)
          .eq('person', person)
          .eq('day', dayKey);
        if (delErr) {
          setError(delErr.message);
          await loadSessions(householdCode);
        }
      } else {
        const row: SessionRow = {
          household_code: householdCode,
          person,
          day: dayKey,
        };
        setSessions((prev) => [...prev, row]);
        const { error: insErr } = await supabase.from('aulas_pt_sessions').insert(row);
        if (insErr) {
          setError(insErr.message);
          await loadSessions(householdCode);
        }
      }
    },
    [householdCode, sessions, loadSessions],
  );

  const daysForPersonMonth = useCallback(
    (person: Person, year: number, monthIndex: number): number[] => {
      const prefix = `${year}-${String(monthIndex + 1).padStart(2, '0')}-`;
      return sessions
        .filter((s) => s.person === person && s.day.startsWith(prefix))
        .map((s) => Number(s.day.slice(8, 10)))
        .filter((n) => Number.isFinite(n))
        .sort((a, b) => a - b);
    },
    [sessions],
  );

  const sessionSet = useMemo(() => {
    const set = new Set<string>();
    for (const s of sessions) set.add(`${s.person}:${s.day}`);
    return set;
  }, [sessions]);

  useEffect(() => {
    return () => {
      if (saveTimer.current) clearTimeout(saveTimer.current);
      if (savedFlash.current) clearTimeout(savedFlash.current);
    };
  }, []);

  return {
    settings,
    sessions,
    sessionSet,
    loading,
    error,
    saveStatus,
    refresh,
    scheduleSettingsSave,
    persistSettings,
    toggleSession,
    daysForPersonMonth,
  };
}
