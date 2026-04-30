"""Extract activities from Dovrebanen Excel into project.json data[].

Source file: utkast dovrebanen brubytte.xlsx
Schema: 4 columns (Task Name | Start | Duration | Finish), 1 sheet, 53 rows.
Hierarchy is encoded as leading spaces in Task Name (3 spaces per level).

Section model: 3 disciplines (rows)
  0: Tverrfaglig / leveranser
  1: Konstruksjon / bru
  2: Geoteknikk

Work-type model: 3 activity kinds (colors)
  proj: Prosjektering / utarbeidelse
  lev:  Leveranse / milepæl
  kvs:  Kvalitetssikring / godkjenning
"""
import openpyxl, json, datetime, sys, io, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

XLSX = r"C:\Users\MHKK\fremdrift\Dovrebanen\utkast dovrebanen brubytte.xlsx"

QUARTERS = [
    ('q2-26', datetime.date(2026, 5, 1),  datetime.date(2026, 6, 30)),
    ('q3-26', datetime.date(2026, 7, 1),  datetime.date(2026, 9, 30)),
    ('q4-26', datetime.date(2026, 10, 1), datetime.date(2026, 12, 31)),
    ('q1-27', datetime.date(2027, 1, 1),  datetime.date(2027, 3, 31)),
    ('q2-27', datetime.date(2027, 4, 1),  datetime.date(2027, 6, 30)),
    ('q3-27', datetime.date(2027, 7, 1),  datetime.date(2027, 9, 30)),
]

def to_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    return None

def section_for(name_lower):
    if any(k in name_lower for k in ['kvalitetsplan', 'oppstart', 'kontrakt', 'månedsrapport',
                                      'tidlig oppstart', 'oppstartsaktivit', 'prosjektstyring',
                                      'gjennomgang av grunnlagsdata', 'utlysning', 'bistand byggherre',
                                      'sluttfrist', 'df1', 'df2']):
        return 0
    if any(k in name_lower for k in ['geoteknikk', 'grunnund', 'fundament', 'grunnforhold']):
        return 2
    return 1

def worktype_for(name_lower):
    # Deliverables / milestones
    if any(k in name_lower for k in ['df1', 'df2', 'df3', 'df4', 'df 4', 'df5', 'df6', 'df7',
                                      'leveranse', 'levere', 'levering',
                                      'godkjent ', 'sluttfrist', 'utlysning',
                                      'kvalitetsplan']):
        return 'lev'
    # Quality / approval / review
    if any(k in name_lower for k in ['godkjen', 'kontroll', 'kvalitetsikr', 'kvalitetssikr',
                                      'gjennomsyn', 'oppretting', 'høringsutkast', 'tfk',
                                      'rams', 'ks ']):
        return 'kvs'
    return 'proj'

def duration_days(s):
    if not s: return 0
    m = re.search(r'(\d+)', str(s))
    return int(m.group(1)) if m else 0

def quarters_for(start, finish):
    out = []
    for qid, qfrom, qto in QUARTERS:
        if start <= qto and finish >= qfrom:
            out.append(qid)
    return out

def clean_label(s, max_len=70):
    s = re.sub(r'\s+', ' ', s).strip().rstrip('*').rstrip()
    replacements = {
        'Detaljprosjektering - Analyse og utarbeide fagmodeller': 'Detaljprosjektering – fagmodeller (MMI300)',
        'Konseptfase - Optimaliseringer, utarbeide og analysere alternativer': 'Konseptfase – alternativer & optimalisering',
        '3. parts kontroll og godkjenning Bane NOR teknologi': '3. parts kontroll & godkjenning Bane NOR Teknologi',
        'Levere høringsutkast arbeidsunderlag til gjennomgang (modell og tegninger)': 'Høringsutkast arbeidsunderlag (DF7)',
        'Utføre supplerende grunnundersøkelser': 'Supplerende grunnundersøkelser',
        'Bane NOR Utlysningsperiode': 'Bane NOR utlysningsperiode',
        'Bistand byggherre evt avklaringer konkurransegrunnlag': 'Bistand byggherre — avklaringer KG',
        'Levere komplett konkurransegrunnlag (DF 4) - MMI375 (modell og tegninger)': 'DF4 – Konkurransegrunnlag (MMI375)',
        'Levering endelig arbeidsunderlag (DF7) - MMI400': 'DF7 – Arbeidsunderlag (MMI400)',
        'Godkjent endelig arbeidsunderlag (DF7) - MMI400': 'DF7 godkjent (MMI400)',
        'KS og Leveranse av konseptfase og valg av brokonsept (DF3)': 'DF3 – Valg av brukonsept',
        'Konkurransegrunnlag høringsutkast til gjennomsyn BN': 'Høringsutkast konkurransegrunnlag',
        'Kvalitetsplan': 'DF1 – Kvalitetsplan',
        'Bane NOR godkjenningsperiode': 'Bane NOR godkjenning',
        'Oppfølging I byggetid': 'Oppfølging i byggetid',
        'Sluttfrist': 'Sluttfrist – prosjektslutt',
    }
    if s in replacements:
        s = replacements[s]
    if len(s) > max_len:
        s = s[:max_len-1] + '…'
    return s

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb['Sheet1']

# Read all rows: (raw_name, start, dur, finish, indent_level)
rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    name, start, dur, finish = r[:4]
    if not name: continue
    raw = str(name)
    # leading spaces / 3 = indent level
    stripped = raw.lstrip()
    indent = (len(raw) - len(stripped)) // 3
    s = to_date(start)
    f = to_date(finish)
    if not s or not f: continue
    # Drop trailing * marker on names (used in source for sub-project markers)
    name_clean = stripped.strip().rstrip('*').rstrip()
    rows.append((name_clean, s, f, indent, duration_days(dur), raw))

print(f"Total rows: {len(rows)}", file=sys.stderr)
for r in rows: print(f"  L{r[3]} {r[4]:>3}d {r[1]} → {r[2]} | {r[0][:60]}", file=sys.stderr)

data = []
seen = set()

# Pick rule:
#   - Top-level (indent 0) parents EXCEPT "Byggeplan" (which is the whole project span)
#   - Level-1 phase containers (indent 1)
#   - Specific named milestones from deeper levels (DF1 Kvalitetsplan, DF3, DF4, DF7, Sluttfrist)
KEEP_DEEP_PATTERNS = [
    'kvalitetsplan',                              # DF1
    'leveranse av konseptfase',                   # DF3
    'levere komplett konkurransegrunnlag',        # DF4
    'levering endelig arbeidsunderlag',           # DF7 levering
    'godkjent endelig arbeidsunderlag',           # DF7 godkjent
    'godkjent konkurransegrunnlag',               # KG godkjent milestone
    'bane nor godkjenningsperiode',               # spans both KG and DF7 phases
    'høringsutkast arbeidsunderlag',              # only the "Levere ..." one
    'konkurransegrunnlag høringsutkast',          # KG høringsutkast
]
SKIP_NAMES = {
    'byggeplan - brufornyelse breivegen og kvam - dovrebanen',  # whole-project parent
    'utarbeide arbeidsunderlag',                                # parent of L1 group, redundant
    'tfk høringsutkast arbeidsunderlag',                        # 1-day step
    'oppretting etter gjennomgang',                             # internal step
    'kvalitetsikre endelig arbeidsunderlag',                    # internal step
    'utarbeide arbeidsunderlag - mmi400',                       # subsumed by DF7 levering
}

for stripped, s, f, indent, days, raw in rows:
    nl = stripped.lower()
    if nl in SKIP_NAMES: continue

    keep = False
    if indent == 0: keep = True
    elif indent == 1: keep = True
    elif any(p in nl for p in KEEP_DEEP_PATTERNS): keep = True
    if not keep: continue

    # Skip pure 0-day milestones except DF/Sluttfrist named ones
    is_named_milestone = any(k in nl for k in ['df3', 'df 4', 'df4', 'df6', 'df7', 'sluttfrist',
                                                'godkjent konkurransegrunnlag', 'kvalitetsplan'])
    if days == 0 and not is_named_milestone: continue

    si = section_for(nl)
    wt = worktype_for(nl)
    label = clean_label(stripped)

    quarters = quarters_for(s, f)
    if not quarters: continue
    for q in quarters:
        key = (si, q, wt, label)
        if key in seen: continue
        seen.add(key)
        data.append([si, q, wt, label, 1, 0])

# Sort by section, quarter, worktype, label
qorder = {q[0]: i for i, q in enumerate(QUARTERS)}
data.sort(key=lambda d: (d[0], qorder.get(d[1], 99), d[2], d[3]))

print(f"Activities: {len(data)}", file=sys.stderr)
print(json.dumps(data, ensure_ascii=False))
