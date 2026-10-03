import type { Person, Settings } from './types';

const MONTHS_PT = [
  'janeiro',
  'fevereiro',
  'março',
  'abril',
  'maio',
  'junho',
  'julho',
  'agosto',
  'setembro',
  'outubro',
  'novembro',
  'dezembro',
];

const WEEKDAYS_PT = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'];

export function monthLabel(year: number, monthIndex: number): string {
  return `${MONTHS_PT[monthIndex]} ${year}`;
}

export function monthTitle(year: number, monthIndex: number): string {
  const name = MONTHS_PT[monthIndex];
  return name.charAt(0).toUpperCase() + name.slice(1) + ` ${year}`;
}

export function weekdayLabels(): string[] {
  return WEEKDAYS_PT;
}

/** Format cents as Portuguese euros, e.g. 2500 → "25,00 €" */
export function formatEuros(cents: number): string {
  const safe = Number.isFinite(cents) ? cents : 0;
  const euros = safe / 100;
  const fixed = euros.toFixed(2).replace('.', ',');
  return `${fixed} €`;
}

/** Parse user-typed euro amount (accepts 25, 25,00, 25.00) → cents */
export function parseEurosToCents(input: string): number | null {
  const trimmed = input.trim().replace(/\s/g, '').replace('€', '');
  if (!trimmed) return null;
  const normalized = trimmed.replace(',', '.');
  if (!/^\d+(\.\d{0,2})?$/.test(normalized)) return null;
  const value = Number(normalized);
  if (!Number.isFinite(value) || value < 0) return null;
  return Math.round(value * 100);
}

export function toDayKey(year: number, monthIndex: number, day: number): string {
  const m = String(monthIndex + 1).padStart(2, '0');
  const d = String(day).padStart(2, '0');
  return `${year}-${m}-${d}`;
}

export function personName(settings: Settings | null, person: Person): string {
  if (!settings) return person === 'a' ? 'Pessoa 1' : 'Pessoa 2';
  return person === 'a' ? settings.name_a : settings.name_b;
}

export function buildWhatsAppText(opts: {
  year: number;
  monthIndex: number;
  nameA: string;
  nameB: string;
  daysA: number[];
  daysB: number[];
  priceCents: number;
}): string {
  const label = monthLabel(opts.year, opts.monthIndex);
  const listA = opts.daysA.length ? opts.daysA.join(', ') : '—';
  const listB = opts.daysB.length ? opts.daysB.join(', ') : '—';
  const total = opts.daysA.length + opts.daysB.length;
  const amount = formatEuros(total * opts.priceCents);
  const price = formatEuros(opts.priceCents);

  return [
    `Aulas PT — ${label}`,
    '',
    `${opts.nameA}: ${opts.daysA.length} aulas (${listA})`,
    `${opts.nameB}: ${opts.daysB.length} aulas (${listB})`,
    '',
    `${total} aulas × ${price} = ${amount}`,
  ].join('\n');
}

/** Calendar grid: Monday-first weeks for the given month */
export function buildMonthGrid(year: number, monthIndex: number): (number | null)[][] {
  const first = new Date(year, monthIndex, 1);
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  // JS: 0=Sun … 6=Sat → Monday-first index 0=Mon … 6=Sun
  const startPad = (first.getDay() + 6) % 7;
  const cells: (number | null)[] = [];
  for (let i = 0; i < startPad; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);
  const weeks: (number | null)[][] = [];
  for (let i = 0; i < cells.length; i += 7) {
    weeks.push(cells.slice(i, i + 7));
  }
  return weeks;
}
