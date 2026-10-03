export type Person = 'a' | 'b';

export type Settings = {
  household_code: string;
  name_a: string;
  name_b: string;
  price_cents: number;
  updated_at?: string;
};

export type SessionRow = {
  household_code: string;
  person: Person;
  day: string; // YYYY-MM-DD
};
