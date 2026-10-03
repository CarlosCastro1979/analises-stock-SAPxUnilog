import AsyncStorage from '@react-native-async-storage/async-storage';

const HOUSEHOLD_KEY = 'aulas_pt_household_code';

export async function loadHouseholdCode(): Promise<string | null> {
  const value = await AsyncStorage.getItem(HOUSEHOLD_KEY);
  return value?.trim() ? value.trim() : null;
}

export async function saveHouseholdCode(code: string): Promise<void> {
  await AsyncStorage.setItem(HOUSEHOLD_KEY, code.trim());
}

export async function clearHouseholdCode(): Promise<void> {
  await AsyncStorage.removeItem(HOUSEHOLD_KEY);
}
