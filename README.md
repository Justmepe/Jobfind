# jobwatch

Automated job aggregator that streams new listings to a Discord channel.

It pulls from **stable, officially-public sources** and posts rich embeds via a
Discord webhook:

| Source | Endpoint | Notes |
|---|---|---|
| RemoteOK | `remoteok.com/api` (JSON) | Broad remote roles across fields |
| WeWorkRemotely | `weworkremotely.com/remote-jobs.rss` | Broad remote roles |
| Remotive | `remotive.com/api/remote-jobs` (JSON) | Broad remote roles, many categories |
| Greenhouse | `boards-api.greenhouse.io/...` | Per-company; officially public & stable |
| Lever | `api.lever.co/v0/postings/...` | Per-company; officially public & stable |
| Ashby | `api.ashbyhq.com/posting-api/...` | Per-company; reaches most AI-training labs & startups |
| Adzuna | `api.adzuna.com/...` | **Broad incl. US ONSITE EHS.** Needs a free API key (see below) |

**Pre-configured company boards** (verified live): Greenhouse — `scaleai`, `snorkelai`,
`turing`, `labelbox`, `invisibletech`; Lever — `cority`, `appen`; Ashby — `mercor`,
`handshake`, `listenlabs`. These carry EHS-software-hybrid, data-analyst, and
domain **AI-Trainer** roles. Add/remove tokens freely.

> **Why not LinkedIn?** The `jobs-guest` endpoint is undocumented, rotates its
> HTML classes, and rate-limits/IP-blocks scrapers — it breaks within weeks and
> risks getting your IP flagged. The sources above are meant to be read
> programmatically, so they don't break.

## How it works

1. **Fetch** — query each enabled source.
2. **Filter** — keep jobs whose title/company/tags match any of your `keywords`
   (and, optionally, whose location matches `location_filter`).
3. **Deduplicate** — compare against `seen_jobs.json` so you never get the same
   job twice.
4. **Notify** — post new jobs to Discord as batched embeds (up to 10 per
   message), honoring Discord's rate limits.

## Setup (local)

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Create your config:
   ```bash
   cp config.example.json config.json
   ```
   Edit `config.json` — set your `keywords`, optional `location_filter`, and add
   any Greenhouse/Lever company board tokens you care about (see below).
3. Create a **Discord webhook**: Server Settings → Integrations → Webhooks →
   New Webhook → copy the URL.
4. Provide the webhook as an environment variable (preferred — keeps it out of
   files):
   ```bash
   # PowerShell
   $env:DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/…"
   # bash
   export DISCORD_WEBHOOK_URL="https://discord.com/api/webhooks/…"
   ```
5. **Seed once** so you aren't flooded with every currently-open role:
   ```bash
   python jobwatch.py --seed
   ```
6. From now on, run normally to get only *new* matches:
   ```bash
   python jobwatch.py
   ```

### Commands

| Command | Effect |
|---|---|
| `python jobwatch.py` | Fetch, filter, notify on new jobs, save state |
| `python jobwatch.py --seed` | Mark all current matches as seen, **no** notifications |
| `python jobwatch.py --dry-run` | Print what *would* send; no Discord posts, no state change |

## Config reference

```json
{
  "keywords": ["EHS", "environmental health", "safety", "compliance"],
  "location_filter": ["remote", "united states", "us"],
  "sources": {
    "remoteok": true,
    "weworkremotely": true,
    "greenhouse_boards": ["stripe", "airbnb"],
    "lever_boards": ["netflix"]
  }
}
```

- **keywords** — case-insensitive substring match against title + company +
  tags. Empty list = match everything.
- **location_filter** — case-insensitive substring match against the job's
  location. Empty list = no location filtering. Jobs with an unknown/blank
  location are always kept.
- **max_age_days** — drop any posting older than this many days (default `3`
  for strict first-mover recency). Set to `0` to disable the age filter. Every
  source reports its post date, shown as a "Posted" field on each alert.
- **keep_undated** — when `false` (default), a listing whose post date can't be
  parsed is dropped (strict recency). Set `true` to keep undated listings — do
  this if a source stops providing dates and you'd rather see them than lose
  them.
- **greenhouse_boards / lever_boards** — the company "token" from its careers
  URL. Examples:
  - Greenhouse: `boards.greenhouse.io/`**`stripe`** → use `"stripe"`.
  - Lever: `jobs.lever.co/`**`netflix`** → use `"netflix"`.
  - Add the companies you actually want to track; this is where EHS/compliance
    roles at specific employers will surface reliably.

- **ashby_boards** — the company token from `jobs.ashbyhq.com/`**`token`**. This
  is where most AI-training / data-labeling firms and AI startups post.
- **adzuna** — the broadest source and the best one for **US onsite EHS** roles
  (Indeed no longer offers a public feed). Get a free `app_id` + `app_key` at
  <https://developer.adzuna.com> (2-minute signup), paste them into the `adzuna`
  block, and edit `queries` to the search terms you want. Leave the keys blank
  to skip this source.

The webhook URL is read from `DISCORD_WEBHOOK_URL` first, then
`discord_webhook_url` in config.json. **Prefer the env var** so the secret never
lands in a file.

## Setup (GitHub Actions — hands-free, no server)

The included workflow `.github/workflows/job_check.yml` runs every hour and
**commits the `seen_jobs.json` state back to the repo** so deduplication
survives across runs (a plain Actions job would otherwise start clean each time
and re-notify everything).

1. Push this folder to a GitHub repo. For CI, put your real filters in
   **`config.example.json`** (the workflow copies it to `config.json` at
   runtime).
2. In the repo: Settings → Secrets and variables → Actions → **New repository
   secret**:
   - Name: `DISCORD_WEBHOOK_URL`
   - Value: your webhook URL
3. (Recommended) Run `python jobwatch.py --seed` locally first and commit the
   resulting `seen_jobs.json` (`git add -f seen_jobs.json`) so the first cloud
   run doesn't post every open role.
4. The workflow triggers hourly, or manually via the **Actions** tab →
   *Job Check* → *Run workflow*.

Change the cadence by editing the `cron` line (`0 * * * *` = hourly). Note
GitHub cron is UTC and scheduled runs can be delayed a few minutes under load.

## Files

| File | Purpose | Committed? |
|---|---|---|
| `jobwatch.py` | The pipeline | yes |
| `config.example.json` | Template + CI config | yes |
| `config.json` | Your local config (may hold webhook) | no (gitignored) |
| `seen_jobs.json` | Dedup state | local: no · CI: force-committed |
| `.github/workflows/job_check.yml` | Hourly scheduler | yes |
