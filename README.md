# Dovrebanen — Brubytte Breivegen og Kvam

COWI's prosjekteringsoppdrag for Bane NOR. Two-page web overview:

1. **Anleggsgjennomføringsplan** (landing page) — fasebeskrivelse fra forarbeider til ferdigstillelse, med togstatus per fase
2. **Fremdriftsplan** — Gantt-stil tidsplan over prosjekteringen Q2 2026 → Q3 2027

**Live:** https://khakpour27.github.io/dovrebanen-fremdrift/

## Files

- `index.html` — landing page: anleggsgjennomføringsplan (8 faser, statisk innhold)
- `fremdriftsplan.html` — single-file Gantt fremdriftsplan, data-driven from `project.json`
- `project.json` — all fremdriftsplan data (sections, columns, work types, activities, meta)
- `extract.py` — one-off Python script that read the source Excel and produced the `data[]` array

## Updating fremdriftsplan data

Edit `project.json` directly. The Gantt page reloads from it on every refresh.

To re-extract from the source Excel after schedule changes:
```
python extract.py > _data.json
# paste _data.json into project.json's "data" field
```

The Excel source lives at `C:\Users\MHKK\fremdrift\Dovrebanen\utkast dovrebanen brubytte.xlsx`.

## Updating anleggsgjennomføringsplan

The phase table on the landing page is hardcoded in `index.html`. Edit the `<tbody>` rows directly to change phase descriptions, togstatus pills, or activities.

## Local preview

```
python -m http.server 3003
# open http://localhost:3003/
```

## Project context

- **Client:** Bane NOR
- **Vendor:** COWI
- **Scope:** Prosjektering av brubytte ved Breivegen og Kvam
- **Timeline:** 18. mai 2026 — 28. september 2027 (sluttfrist)
- **Key milestones:** DF1 (Kvalitetsplan), DF3 (Brukonsept), DF4 (Konkurransegrunnlag, MMI375), DF7 (Arbeidsunderlag, MMI400)
