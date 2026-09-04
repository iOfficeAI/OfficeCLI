# Deck industry theme pack (catalog 1.5.3)

CSBU WorkMate original vertical themes — **not** Dashi / AGPL clones. Token shape matches brand themes (`background` / `surface` / `text` / `mutedText` / `accent` / `fontFamily`).

| id | label | mode | accent | notes |
|----|-------|------|--------|-------|
| `industry-finance` | Industry Finance | light | `#0B6E4F` | Ledger / banking green + Georgia |
| `industry-consulting` | Industry Consulting | light | `#C2410C` | Briefcase warm stone + terracotta |
| `industry-tech` | Industry Tech | dark | `#6366F1` | Product / SaaS indigo night |
| `industry-education` | Industry Education | light | `#0284C7` | Campus sky (distinct from `education-warm`) |

Brand pack remains `csbu-workmate` / `csbu-workmate-night`. Catalog version **1.5.3** (18 themes). Remap with:

```bash
officecli deck theme-remap <spec> --to industry-finance --json
officecli deck theme-remap <spec> --to industry-tech --apply --write-report --json
```
