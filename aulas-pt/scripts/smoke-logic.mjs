/** Smoke-test pure formatting / WhatsApp text (no Expo runtime). */

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

function formatEuros(cents) {
  return `${(cents / 100).toFixed(2).replace('.', ',')} €`;
}

function parseEurosToCents(input) {
  const trimmed = input.trim().replace(/\s/g, '').replace('€', '');
  if (!trimmed) return null;
  const normalized = trimmed.replace(',', '.');
  if (!/^\d+(\.\d{0,2})?$/.test(normalized)) return null;
  return Math.round(Number(normalized) * 100);
}

function buildWhatsAppText({ year, monthIndex, nameA, nameB, daysA, daysB, priceCents }) {
  const label = `${MONTHS_PT[monthIndex]} ${year}`;
  const listA = daysA.length ? daysA.join(', ') : '—';
  const listB = daysB.length ? daysB.join(', ') : '—';
  const total = daysA.length + daysB.length;
  return [
    `Aulas PT — ${label}`,
    '',
    `${nameA}: ${daysA.length} aulas (${listA})`,
    `${nameB}: ${daysB.length} aulas (${listB})`,
    '',
    `${total} aulas × ${formatEuros(priceCents)} = ${formatEuros(total * priceCents)}`,
  ].join('\n');
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

assert(parseEurosToCents('25') === 2500, 'parse 25');
assert(parseEurosToCents('25,00') === 2500, 'parse 25,00');
assert(parseEurosToCents('12.5') === 1250, 'parse 12.5');
assert(formatEuros(35000) === '350,00 €', 'format 350');

const text = buildWhatsAppText({
  year: 2026,
  monthIndex: 9,
  nameA: 'Pessoa 1',
  nameB: 'Pessoa 2',
  daysA: [1, 3, 6, 8, 10, 13, 15, 20],
  daysB: [2, 4, 7, 9, 14, 18],
  priceCents: 2500,
});

const expected = `Aulas PT — outubro 2026

Pessoa 1: 8 aulas (1, 3, 6, 8, 10, 13, 15, 20)
Pessoa 2: 6 aulas (2, 4, 7, 9, 14, 18)

14 aulas × 25,00 € = 350,00 €`;

assert(text === expected, `WhatsApp text mismatch:\n${text}`);
console.log('smoke-logic OK');
