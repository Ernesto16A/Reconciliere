# Packaging Reconciliation

A browser-only tool for reconciling client vs. factory packaging records
(delivery number, box number, quantity), for both In and Out flows.

Everything runs client-side — Excel files are parsed in the browser
(SheetJS), nothing is uploaded to a server. Client profiles (column
mappings + box number mappings) are saved in the browser's `localStorage`,
so they persist between visits on the same device/browser, but don't sync
across devices or get backed up anywhere.

## Run it locally

```bash
npm install
npm run dev
```

Then open the printed local URL (usually http://localhost:5173).

## Deploy to GitHub Pages

This repo includes a GitHub Actions workflow
(`.github/workflows/deploy.yml`) that builds the app and publishes it to
GitHub Pages automatically.

1. Push this repo to GitHub.
2. In the repo, go to **Settings → Pages**.
3. Under **Build and deployment → Source**, choose **GitHub Actions**.
4. Push to the `main` branch (or re-run the workflow from the **Actions**
   tab). Once it finishes, the deployed URL will show up under
   **Settings → Pages** and in the workflow's summary.

No further configuration is needed — the Vite build uses a relative base
path so it works regardless of the repository name.

## Notes / limitations

- No backend — data isn't shared between people. Each person who opens the
  deployed link has their own separate `localStorage`, and profiles you
  create on one device won't show up on another.
- Uploaded files and reconciliation results are **not** saved anywhere —
  only client profiles (column mappings, box number mappings) persist.
- Column matching for both client and factory uploads relies on exact
  header names configured per client profile — there's no automatic
  guessing of column names.
