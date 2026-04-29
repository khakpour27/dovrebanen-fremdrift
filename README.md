# Dovrebanen Fremdriftsplan

Web-based fremdriftsplan for COWI's prosjekteringsoppdrag for Bane NOR — brufornyelse Breivegen og Kvam, Dovrebanen.

**Live:** https://khakpour27.github.io/dovrebanen-fremdrift/

## Files

- `index.html` — single-file Gantt-style fremdriftsplan (data-driven)
- `project.json` — all project data (sections, columns, work types, activities, metadata)
- `extract.py` — one-off Python script that read the source Excel and produced the `data[]` array

## Updating data

Edit `project.json` directly. The page reloads from it on every refresh.

To re-extract from the source Excel after schedule changes:
```
python extract.py > _data.json
# manually paste _data.json into project.json's "data" field
```

The Excel source lives at `C:\Users\MHKK\fremdrift\Dovrebanen\Fremdriftsplan Prosjektering Bruer Dovrebanen excel.xlsx`.

## Local preview

```
python -m http.server 3003
# open http://localhost:3003/
```

## Project context

- **Client:** Bane NOR
- **Vendor:** COWI
- **Scope:** Prosjektering av brubytte ved Breivegen og Kvam
- **Timeline:** 18. mai 2026 — 31. desember 2027
- **Key milestones:** DF1 (Kvalitetsplan), DF3 (Brukonsept), DF4 (Konkurransegrunnlag), DF6 (Sluttdokumentasjon), DF7 (Arbeidsunderlag MMI400)
