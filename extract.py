"""Extract activities from Dovrebanen Excel into project.json data[].

Section model: 5 disciplines (rows)
  0: Tverrfaglig / leveranser
  1: Konstruksjon / bru
  2: Geoteknikk
  3: Bane (spor, signal, kontaktledning)
  4: Vei og adkomst

Work-type model: 3 activity kinds (colors)
  proj: Prosjektering / utarbeidelse
  lev:  Leveranse / milepæl
  kvs:  Kvalitetssikring / godkjenning
"""
import openpyxl, json, datetime, sys, io, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

XLSX = r"C:\Users\MHKK\fremdrift\Dovrebanen\Fremdriftsplan Prosjektering Bruer Dovrebanen excel.xlsx"

QUARTERS = [
    ('q2-26', datetime.date(2026, 5, 1),  datetime.date(2026, 6, 30)),
    ('q3-26', datetime.date(2026, 7, 1),  datetime.date(2026, 9, 30)),
    ('q4-26', datetime.date(2026, 10, 1), datetime.date(2026, 12, 31)),
    ('q1-27', datetime.date(2027, 1, 1),  datetime.date(2027, 3, 31)),
    ('q2-27', datetime.date(2027, 4, 1),  datetime.date(2027, 6, 30)),
    ('q3-27', datetime.date(2027, 7, 1),  datetime.date(2027, 9, 30)),
    ('q4-27', datetime.date(2027, 10, 1), datetime.date(2027, 12, 31)),
]

def parse_date(v):
    if v is None: return None
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    s = str(v).strip()
    for fmt in ('%d %B %Y %H:%M', '%d %B %Y'):
        try: return datetime.datetime.strptime(s, fmt).date()
        except: pass
    return None

def section_for(name, parent_chain):
    n = name.lower()
    full = (name + ' ' + parent_chain).lower()
    # Tverrfaglig leveranser & oppstart
    if any(k in n for k in ['df1', 'df2', 'kvalitetsplan', 'oppstart', 'kontrakt', 'månedsrapport']): return 0
    # Geoteknikk
    if any(k in full for k in ['geoteknikk', 'grunnund', 'fundament', 'grunnforhold']): return 2
    # Bane (spor/signal/KL)
    if any(k in full for k in ['signal', 'sikringsanlegg', 'kontaktledning', 'kl-anlegg', ' kl ', ' kl,', 'kl/', 'sporkonst', ' spor ', 'skinne', 'banestrøm']): return 3
    # Vei
    if any(k in full for k in [' vei ', ' vei,', 'adkomst', 'trafikk', 'veibru', 'veiinfra', 'arbeidsvarsling', 'statens vegvesen']): return 4
    # Konstruksjon (default for bridge work)
    return 1

def worktype_for(name):
    n = name.lower()
    # Deliverables / milestones
    if any(k in n for k in ['df1', 'df2', 'df3', 'df4', 'df5', 'df6', 'milepæ', 'milep', 'leveranse', 'overlevering', 'rapporter', 'månedsrapport', 'kvalitetsplan']): return 'lev'
    # Quality / approval / review
    if any(k in n for k in ['godkjen', 'kontroll', 'kvalitet', 'gjennomsyn', 'innarbeide', 'oppretting', 'tfk', 'tfg', 'uak', 'sha', 'rams', 'ym']): return 'kvs'
    # Default — design/project work
    return 'proj'

def duration_days(dur):
    if not dur: return 0
    m = re.search(r'(\d+)', str(dur))
    return int(m.group(1)) if m else 0

# Discard noise labels (placeholders, single-word disciplines, abbreviations)
NOISE = {'ek', 'kon', 'sk', 'slk', 'spor', 'kl', 'vei', 'konstruksjon', 'geoteknikk', 'sha', 'ym', 'rams'}

def clean_label(s, max_len=70):
    s = re.sub(r'\s+', ' ', s).strip().rstrip('*').rstrip()
    # Friendly relabels
    replacements = {
        'Detaljprosjektering - Analyse og utarbeide fagmodeller': 'Detaljprosjektering – fagmodeller (MMI300)',
        'Konseptfase - Optimaliseringer, utarbeide og analysere alternativer': 'Konseptfase – alternativer & optimalisering',
        '3. parts kontroll og godkjenning Bane NOR teknologi': '3. parts kontroll & godkjenning Bane NOR Teknologi',
        'Levere høringsutkast arbeidsunderlag til gjennomgang (modell, tegninger og beskrivelser)': 'Høringsutkast arbeidsunderlag',
        'DF6 - Overlevering av slutdokumentasjon iht krav Kap C4': 'DF6 – Sluttdokumentasjon',
        'Oppfølging I byggetid': 'Oppfølging i byggetid',
        'Utføre supplerende grunnundersøkelser': 'Supplerende grunnundersøkelser',
    }
    if s in replacements:
        s = replacements[s]
    if len(s) > max_len:
        s = s[:max_len-1] + '…'
    return s

def is_noise(name):
    return name.strip().lower() in NOISE

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb['Task_Table']
header = [c.value for c in ws[1]]

def find_col(*candidates):
    for cand in candidates:
        for i, h in enumerate(header):
            if h and cand.lower() == str(h).lower(): return i
    return None

idx_name   = find_col('Name')
idx_start  = find_col('Start')
idx_finish = find_col('Finish')
idx_level  = find_col('Outline Level')
idx_dur    = find_col('Duration')

rows_raw = []
parent_at = {}
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r[idx_name]: continue
    name   = str(r[idx_name]).strip()
    start  = r[idx_start]
    finish = r[idx_finish]
    try: level = int(r[idx_level])
    except: level = 0
    dur = r[idx_dur]
    parent_at[level] = name
    for k in list(parent_at.keys()):
        if k > level: del parent_at[k]
    parent_chain = ' '.join(parent_at[k] for k in sorted(parent_at.keys()) if k < level)
    rows_raw.append((name, start, finish, level, parent_chain, dur))

# Use L3 phase containers + DF deliverables (skip L4-L5 which are sub-steps)
data = []
seen = set()

# Skip these L3 entries (covered by other items or too procedural)
SKIP_L3 = {
    'igangsettelse av oppdraget',
    'godkjenne plan for grunnundersøkelser', 'godkjenne plan for grunnundersokelser',
    'godkjenne psb',
    'tfk høringsutkast arbeidsunderlag', 'tfk hoeringsutkast arbeidsunderlag',
    'oppretting etter gjennomgang',
    'kvalitetsikre endelig arbeidsunderlag',
    'sluttfrist', 'sluttfrist*',
    'grunnlagsdata mottatt',
    'levere høringsutkast arbeidsunderlag til gjennomgang (modell, tegninger og beskrivelser)',
    'levere hoeringsutkast arbeidsunderlag til gjennomgang (modell, tegninger og beskrivelser)',
}

for name, start, finish, level, parent_chain, dur in rows_raw:
    days = duration_days(dur)
    s, f = parse_date(start), parse_date(finish)
    if not s or not f: continue
    if is_noise(name): continue

    nl_strip = name.lower().rstrip('*').strip()
    if level != 3: continue
    if nl_strip in SKIP_L3: continue
    # Substring-based skips for verbose names
    if nl_strip.startswith('levere høringsutkast'): continue
    if nl_strip.startswith('utarbeide arbeidsunderlag'): continue  # handled by DF7 levering entry
    # Skip pure 0-day entries unless they are key DF or named milestones
    is_df_key = any(name.startswith(k) for k in ['DF1', 'DF3', 'DF4', 'DF6'])
    is_named_milestone = ('DF7' in name) or ('Kontrakt' in name and level == 3)
    if days == 0 and not (is_df_key or is_named_milestone): continue
    # DF2 handled separately (rolled-up)
    if name.startswith('DF2'): continue

    si = section_for(name, parent_chain)
    wt = worktype_for(name)
    label = clean_label(name)

    quarters = [qid for qid, qf, qt in QUARTERS if s <= qt and f >= qf]
    if not quarters: continue

    for q in quarters:
        key = (si, q, wt, label)
        if key in seen: continue
        seen.add(key)
        data.append([si, q, wt, label, 1, 0])

# Add rolled-up DF2 monthly reporting across all quarters
for q in [q[0] for q in QUARTERS]:
    key = (0, q, 'lev', 'Månedsrapportering (DF2)')
    if key not in seen:
        seen.add(key)
        data.append([0, q, 'lev', 'Månedsrapportering (DF2)', 1, 0])

qorder = {q[0]: i for i, q in enumerate(QUARTERS)}
data.sort(key=lambda d: (d[0], qorder.get(d[1], 99), d[2], d[3]))

print(f"Activities: {len(data)}", file=sys.stderr)
print(json.dumps(data, ensure_ascii=False))
