# Refresh proxy — Cloudflare Worker

The dashboard's ↻ Refresh button calls this Worker instead of the GitHub API
directly. The GitHub PAT lives **only** as a Worker secret — never in the repo
(GitHub auto-revokes PATs it finds in public repos) and never in a browser
(so the button works from any machine with zero setup).

```
Browser ──POST (no credentials)──▶ Cloudflare Worker ──PAT──▶ GitHub workflow_dispatch
```

## One-time setup (Cloudflare dashboard)

1. Cloudflare dashboard → **Workers & Pages** → **Create** → **Create Worker**.
2. Name it `hris-refresh` (any name works — the name decides the URL) → **Deploy**.
3. Click **Edit code**, replace the starter code with the contents of
   `worker.js` from this folder → **Deploy**.
4. Go to the Worker's **Settings** → **Variables and Secrets** → **Add**:
   - Type: **Secret**
   - Name: `HRIS_GITHUB_PAT`
   - Value: a GitHub PAT (classic) with `repo` + `workflow` scope
   → **Deploy**.
5. Note the Worker URL shown on its overview page, e.g.
   `https://hris-refresh.<your-subdomain>.workers.dev`, and set it as
   `REFRESH_PROXY_URL` in `generate_dashboard.py` (the Refresh button in
   `index.html` is generated from it).

## When the PAT expires

Generate a new PAT and update the `HRIS_GITHUB_PAT` secret in step 4.
Nothing else changes — no commits, no browser steps.

## Notes

- The Worker only accepts `POST` and only allows browser calls from
  `https://begb0037admin.github.io` (CORS + Origin check).
- The worst anyone could do by calling the URL directly is trigger a
  dashboard refresh run — the PAT itself is never exposed.
