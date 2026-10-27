# Weekly Summary export

Exports the Revel demo Summary page for the current week to Excel and PDF. The week matches that page: the trailing 7 days ending yesterday in America/Chicago, kept inside the dates Supabase actually has. Day columns use that ending business date.

The Excel workbook has three sheets, each set to print on one page: share of week (pie), sales and labor, and the review leaderboard. The PDF has those same three pages. Files are named `Weekly_Summary_YYYY-MM-DD_to_YYYY-MM-DD`.

This script reads the Summary calculations from the sibling repo [hthsolutions/revel-demo](https://github.com/hthsolutions/revel-demo). Place the two folders next to each other, or set `REVEL_DEMO` to the revel-demo path.

## Run locally

```powershell
py -3 -m pip install -r requirements.txt
py -3 export_weekly_summary.py
py -3 export_weekly_summary.py --date 2026-10-26
```

`--date` is the business date the week ends on. `--output-dir` changes where the files are written. The default is this folder.

Credentials are the same ones the Summary app uses. Set `SUPABASE_URL` and `SUPABASE_KEY` in the environment, or keep them in `revel-demo/.streamlit/secrets.toml`. The anon key can sign in and still return no rows.

## GitHub Actions

`.github/workflows/weekly-summary.yml` runs the export when the send pdf Zap dispatches it, and when someone starts **Weekly summary** by hand. The job checks out this repo and `hthsolutions/revel-demo`, then uploads only the PDF as the `weekly-summary` artifact (kept 90 days). The Zap downloads that artifact, unzips it, and sends the PDF.

The workflow runs only after this repo is on GitHub. Add these repository secrets first. Use the Summary app key, not the anon key.

- `SUPABASE_URL`
- `SUPABASE_KEY`
