// HRIS Dashboard — Refresh proxy
// ------------------------------------------------------------------
// Runs on Cloudflare Workers. Holds the GitHub PAT as a Worker secret
// (HRIS_GITHUB_PAT) so no token ever lives in the public repo or in a
// browser. The dashboard's Refresh button POSTs here with no
// credentials; this worker forwards the workflow_dispatch call to the
// GitHub API.
//
// Setup: see README.md in this folder.

const GITHUB_REPO = "begb0037admin/hris-dashboard";
const WORKFLOW_FILE = "update-dashboard.yml";
const BRANCH = "main";

// Only the GitHub Pages dashboard may call this from a browser.
const ALLOWED_ORIGIN = "https://begb0037admin.github.io";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function json(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
  });
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }
    if (request.method !== "POST") {
      return json(405, { ok: false, error: "Method not allowed — POST only." });
    }

    // Browsers always send Origin on cross-site POST; reject other sites.
    const origin = request.headers.get("Origin");
    if (origin && origin !== ALLOWED_ORIGIN) {
      return json(403, { ok: false, error: "Forbidden origin." });
    }

    if (!env.HRIS_GITHUB_PAT) {
      return json(500, {
        ok: false,
        error: "Worker secret HRIS_GITHUB_PAT is not set — add it under Settings → Variables and Secrets.",
      });
    }

    const resp = await fetch(
      `https://api.github.com/repos/${GITHUB_REPO}/actions/workflows/${WORKFLOW_FILE}/dispatches`,
      {
        method: "POST",
        headers: {
          "Authorization": "Bearer " + env.HRIS_GITHUB_PAT,
          "Accept": "application/vnd.github+json",
          "Content-Type": "application/json",
          // GitHub's API rejects requests without a User-Agent.
          "User-Agent": "hris-dashboard-refresh-worker",
        },
        body: JSON.stringify({ ref: BRANCH }),
      }
    );

    if (resp.status === 204) {
      return json(200, { ok: true });
    }

    const text = await resp.text();
    return json(502, {
      ok: false,
      status: resp.status,
      error: `GitHub API ${resp.status}: ${text.slice(0, 200)}`,
    });
  },
};
