"""Probe ZFACT valor heuristics — NF 99642 regression."""
import re

SAP_VALOR_MAX = 5_000_000

def norm_nf_key(nf):
    s = re.sub(r'\D', '', str(nf or ''))
    if not s:
        return ''
    return s.lstrip('0') or '0'

def parse_sap_num(v):
    if v is None or v == '':
        return 0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(' ', '')
    if not s or s == '-':
        return 0
    s = re.sub(r'^R\$', '', s, flags=re.I)
    if ',' in s:
        return float(s.replace('.', '').replace(',', '.'))
    if '.' in s:
        parts = s.split('.')
        if len(parts) > 2:
            return float(s.replace('.', ''))
        int_part, frac = parts[0], parts[1] if len(parts) > 1 else ''
        if not frac or re.fullmatch(r'0+', frac):
            return float(int_part)
        return float(s)
    return float(s)

def has_sap_valor_formatting(v):
    s = str(v).strip().replace(' ', '')
    return bool(re.match(r'^R\$', s, re.I)) or ',' in s

def looks_like_sap_doc_or_nf_number(v):
    if v is None or v == '':
        return False
    if has_sap_valor_formatting(v):
        return False
    key = norm_nf_key(v)
    if len(key) >= 8:
        return True
    s = re.sub(r'^R\$', '', str(v).strip().replace(' ', ''), flags=re.I)
    if re.fullmatch(r'\d{7,}', s):
        return True
    if isinstance(v, int) and abs(v) >= 1_000_000:
        return True
    return False

def looks_like_sap_valor_cell(v, trust_column=False):
    if v is None or v == '':
        return False
    n = parse_sap_num(v)
    if not (n > 0) or n > SAP_VALOR_MAX:
        return False
    if looks_like_sap_doc_or_nf_number(v):
        return False
    s = str(v).strip().replace(' ', '')
    if re.match(r'^R\$', s, re.I) or ',' in s:
        return True
    if trust_column:
        if isinstance(v, int) and v >= 1_000_000_000:
            return False
        return True
    if re.fullmatch(r'\d{1,3}', s) and n < 1000:
        return False
    return n >= 10 or '.' in s or (isinstance(v, float))

def sap_parsed_valor(v):
    return parse_sap_num(v) if looks_like_sap_valor_cell(v, trust_column=True) else 0

def sap_row_line_valor(r):
    v = r.get('valorLine')
    if v is None and not r.get('valorDoc'):
        v = r.get('valorNF')
    return sap_parsed_valor(v)

def build_sum(rows):
    total = 0
    for r in rows:
        if r.get('_hasMaterialCol') and not (r.get('material') or '').strip():
            continue
        total += sap_row_line_valor(r)
    return total

rows = [
    {'material': '7806814 CUBA', 'valorLine': '0,00', '_hasMaterialCol': True},
    {'material': '7806814 CUBA', 'valorLine': 308.91, '_hasMaterialCol': True},
    {'material': '7806814 CUBA', 'valorLine': 679.59, '_hasMaterialCol': True},
    {'material': '7801734 COLÔMBIA', 'valorLine': '0,00', '_hasMaterialCol': True},
    {'material': '7801734 COLÔMBIA', 'valorLine': 988.5, '_hasMaterialCol': True},
    {'material': '7802415 TIMOR', 'valorLine': '0,00', '_hasMaterialCol': True},
    {'material': '7802415 TIMOR', 'valorLine': 786, '_hasMaterialCol': True},
]

print('786 trustColumn (OLD order):', looks_like_sap_valor_cell(786, trust_column=False))
print('786 trustColumn (NEW):', looks_like_sap_valor_cell(786, trust_column=True))
print('sap_parsed_valor(786) OLD logic:', sap_parsed_valor(786) if looks_like_sap_valor_cell(786, trust_column=False) else 0)
print('sap_parsed_valor(786) NEW logic:', sap_parsed_valor(786))
total = build_sum(rows)
print('99642 sum:', total, 'OK' if abs(total - 2763) < 0.01 else f'FAIL expected 2763')
