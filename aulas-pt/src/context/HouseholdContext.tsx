import React, {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
} from 'react';
import {
  clearHouseholdCode,
  loadHouseholdCode,
  saveHouseholdCode,
} from '../lib/storage';
import { useAulasData } from '../hooks/useAulasData';

type HouseholdContextValue = {
  ready: boolean;
  householdCode: string | null;
  joinHousehold: (code: string) => Promise<void>;
  leaveHousehold: () => Promise<void>;
  data: ReturnType<typeof useAulasData>;
};

const HouseholdContext = createContext<HouseholdContextValue | null>(null);

export function HouseholdProvider({ children }: { children: React.ReactNode }) {
  const [ready, setReady] = useState(false);
  const [householdCode, setHouseholdCode] = useState<string | null>(null);
  const data = useAulasData(householdCode);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const code = await loadHouseholdCode();
      if (!cancelled) {
        setHouseholdCode(code);
        setReady(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const joinHousehold = useCallback(async (code: string) => {
    const cleaned = code.trim().toLowerCase().replace(/\s+/g, '-');
    if (!cleaned) throw new Error('Indica um código da casa.');
    await saveHouseholdCode(cleaned);
    setHouseholdCode(cleaned);
  }, []);

  const leaveHousehold = useCallback(async () => {
    await clearHouseholdCode();
    setHouseholdCode(null);
  }, []);

  const value = useMemo(
    () => ({
      ready,
      householdCode,
      joinHousehold,
      leaveHousehold,
      data,
    }),
    [ready, householdCode, joinHousehold, leaveHousehold, data],
  );

  return (
    <HouseholdContext.Provider value={value}>{children}</HouseholdContext.Provider>
  );
}

export function useHousehold() {
  const ctx = useContext(HouseholdContext);
  if (!ctx) throw new Error('useHousehold must be used within HouseholdProvider');
  return ctx;
}
