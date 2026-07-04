#!/usr/bin/env python3
"""
jobwatch — automated job aggregator that streams new listings to Discord.

Pulls from stable, officially-public sources (RemoteOK JSON API, WeWorkRemotely
RSS, and any Greenhouse / Lever company boards you list), filters by keyword,
deduplicates against a local state file, and posts rich embeds to a Discord
webhook.

Usage:
    python jobwatch.py            # fetch, filter, notify on new jobs
    python jobwatch.py --seed     # mark everything currently listed as "seen"
                                   #   WITHOUT notifying (run once on first setup
                                   #   so you don't get flooded)
    python jobwatch.py --dry-run  # print what would be sent; no Discord posts

Config lives in config.json (see config.example.json). The Discord webhook URL
is read from the DISCORD_WEBHOOK_URL environment variable first, then config.
"""

import argparse
import email.utils
import html
import json
import os
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from urllib.parse import quote

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(HERE, "config.json")
SEEN_JOBS_FILE = os.path.join(HERE, "seen_jobs.json")

# Polite, identifiable UA. RemoteOK and others reject blank/obvious-bot agents.
USER_AGENT = "jobwatch/1.0 (+https://github.com/; personal job alert bot)"
REQUEST_TIMEOUT = 20
DISCORD_EMBED_COLOR = 3447003  # deep blue
MAX_EMBEDS_PER_MESSAGE = 10  # Discord hard limit
MAX_NOTIFICATIONS_PER_RUN = 40  # safety valve against a flood


# --------------------------------------------------------------------------- #
# Config + state
# --------------------------------------------------------------------------- #
def load_config():
    if not os.path.exists(CONFIG_FILE):
        sys.exit(
            f"No config found at {CONFIG_FILE}.\n"
            "Copy config.example.json to config.json and edit it."
        )
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    cfg.setdefault("keywords", [])
    cfg.setdefault("location_filter", [])
    cfg.setdefault("max_age_days", 14)  # drop postings older than this; 0 = no limit
    cfg.setdefault("keep_undated", True)  # keep jobs whose post date is unknown
    cfg.setdefault("sources", {})
    cfg["sources"].setdefault("remoteok", True)
    cfg["sources"].setdefault("weworkremotely", True)
    cfg["sources"].setdefault("remotive", True)
    cfg["sources"].setdefault("greenhouse_boards", [])
    cfg["sources"].setdefault("lever_boards", [])
    cfg["sources"].setdefault("ashby_boards", [])
    cfg["sources"].setdefault("adzuna", {})
    cfg["sources"].setdefault("amazon", {})

    webhook = os.environ.get("DISCORD_WEBHOOK_URL") or cfg.get("discord_webhook_url")
    cfg["_webhook"] = webhook
    return cfg


def load_seen():
    if os.path.exists(SEEN_JOBS_FILE):
        with open(SEEN_JOBS_FILE, "r", encoding="utf-8") as f:
            try:
                return set(json.load(f))
            except json.JSONDecodeError:
                return set()
    return set()


def save_seen(seen_ids):
    with open(SEEN_JOBS_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(seen_ids), f, indent=0)


APPLICATIONS_QUEUE_FILE = os.path.join(HERE, "applications_queue.json")


def queue_for_application(jobs, keep=100):
    """Append newly-notified jobs so apply.py can draft cover letters for them.

    Stores a compact record per job; keeps the most recent `keep` entries.
    Never blocks the pipeline — any error here is non-fatal.
    """
    try:
        existing = []
        if os.path.exists(APPLICATIONS_QUEUE_FILE):
            with open(APPLICATIONS_QUEUE_FILE, "r", encoding="utf-8") as f:
                existing = json.load(f)
        seen_ids = {j.get("id") for j in existing}
        for j in jobs:
            if j["id"] in seen_ids:
                continue
            existing.append(
                {
                    "id": j["id"],
                    "title": j["title"],
                    "company": j["company"],
                    "url": j.get("url", ""),
                    "source": j["source"],
                }
            )
        with open(APPLICATIONS_QUEUE_FILE, "w", encoding="utf-8") as f:
            json.dump(existing[-keep:], f, indent=2)
    except Exception as e:
        print(f"[queue] non-fatal error: {e}", file=sys.stderr)


# --------------------------------------------------------------------------- #
# Normalized job shape:
#   {"id", "title", "company", "url", "location", "source", "tags"}
# --------------------------------------------------------------------------- #
def _get(url, **kwargs):
    headers = {"User-Agent": USER_AGENT, "Accept": kwargs.pop("accept", "*/*")}
    return requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT, **kwargs)


def _iso_ts(s):
    """Parse an ISO-8601 string to a UTC epoch timestamp; None on failure."""
    if not s:
        return None
    try:
        s = s.strip()
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except Exception:
        return None


def _rfc822_ts(s):
    """Parse an RFC-822 date (RSS pubDate) to a UTC epoch timestamp."""
    if not s:
        return None
    try:
        dt = email.utils.parsedate_to_datetime(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except Exception:
        return None


def fetch_remoteok():
    """RemoteOK public JSON API. First array element is a metadata/legal notice."""
    jobs = []
    try:
        r = _get("https://remoteok.com/api", accept="application/json")
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"[remoteok] error: {e}", file=sys.stderr)
        return jobs

    for item in data:
        if not isinstance(item, dict) or "id" not in item or "position" not in item:
            continue  # skips the leading legal-notice object
        jobs.append(
            {
                "id": f"remoteok:{item.get('id')}",
                "title": (item.get("position") or "").strip(),
                "company": (item.get("company") or "").strip(),
                "url": item.get("url") or item.get("apply_url") or "",
                "location": (item.get("location") or "Remote").strip() or "Remote",
                "source": "RemoteOK",
                "posted_ts": item.get("epoch")
                if isinstance(item.get("epoch"), (int, float))
                else _iso_ts(item.get("date")),
                "tags": [str(t) for t in (item.get("tags") or [])],
            }
        )
    return jobs


def fetch_weworkremotely():
    """WeWorkRemotely main RSS feed. Titles look like 'Company: Role Title'."""
    jobs = []
    try:
        r = _get("https://weworkremotely.com/remote-jobs.rss", accept="application/rss+xml")
        r.raise_for_status()
        root = ET.fromstring(r.content)
    except Exception as e:
        print(f"[weworkremotely] error: {e}", file=sys.stderr)
        return jobs

    for item in root.iter("item"):
        title_raw = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        guid = (item.findtext("guid") or link).strip()
        region = (item.findtext("region") or "").strip()
        pub_ts = _rfc822_ts(item.findtext("pubDate"))
        if not title_raw or not link:
            continue
        if ":" in title_raw:
            company, title = title_raw.split(":", 1)
            company, title = company.strip(), title.strip()
        else:
            company, title = "", title_raw
        jobs.append(
            {
                "id": f"wwr:{guid}",
                "title": html.unescape(title),
                "company": html.unescape(company),
                "url": link,
                "location": region or "Remote",
                "source": "WeWorkRemotely",
                "posted_ts": pub_ts,
                "tags": [],
            }
        )
    return jobs


def fetch_greenhouse(board):
    """Greenhouse public board API: officially public, very stable."""
    jobs = []
    url = f"https://boards-api.greenhouse.io/v1/boards/{quote(board)}/jobs"
    try:
        r = _get(url, accept="application/json")
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"[greenhouse:{board}] error: {e}", file=sys.stderr)
        return jobs

    for item in data.get("jobs", []):
        jobs.append(
            {
                "id": f"greenhouse:{board}:{item.get('id')}",
                "title": (item.get("title") or "").strip(),
                "company": board,
                "url": item.get("absolute_url") or "",
                "location": ((item.get("location") or {}).get("name") or "").strip(),
                "source": f"Greenhouse ({board})",
                "posted_ts": _iso_ts(item.get("updated_at") or item.get("first_published")),
                "tags": [],
            }
        )
    return jobs


def fetch_lever(board):
    """Lever public postings API: officially public JSON."""
    jobs = []
    url = f"https://api.lever.co/v0/postings/{quote(board)}?mode=json"
    try:
        r = _get(url, accept="application/json")
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"[lever:{board}] error: {e}", file=sys.stderr)
        return jobs

    for item in data:
        cats = item.get("categories") or {}
        created = item.get("createdAt")
        jobs.append(
            {
                "id": f"lever:{board}:{item.get('id')}",
                "title": (item.get("text") or "").strip(),
                "company": board,
                "url": item.get("hostedUrl") or item.get("applyUrl") or "",
                "location": (cats.get("location") or "").strip(),
                "source": f"Lever ({board})",
                "posted_ts": created / 1000.0
                if isinstance(created, (int, float))
                else None,
                "tags": [t for t in [cats.get("team"), cats.get("commitment")] if t],
            }
        )
    return jobs


def fetch_ashby(board):
    """Ashby public job-board API. Used by many AI labs & tech startups."""
    jobs = []
    url = f"https://api.ashbyhq.com/posting-api/job-board/{quote(board)}"
    try:
        r = _get(url, accept="application/json")
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"[ashby:{board}] error: {e}", file=sys.stderr)
        return jobs

    for item in data.get("jobs", []):
        loc = (item.get("location") or "").strip()
        if item.get("isRemote") and "remote" not in loc.lower():
            loc = (loc + " (Remote)").strip()
        tags = [
            t
            for t in [item.get("department"), item.get("team"), item.get("employmentType")]
            if t
        ]
        jobs.append(
            {
                "id": f"ashby:{board}:{item.get('id')}",
                "title": (item.get("title") or "").strip(),
                "company": board,
                "url": item.get("jobUrl") or item.get("applyUrl") or "",
                "location": loc,
                "source": f"Ashby ({board})",
                "posted_ts": _iso_ts(item.get("publishedAt")),
                "tags": tags,
            }
        )
    return jobs


def fetch_remotive():
    """Remotive public API — broad remote listings across many categories."""
    jobs = []
    try:
        r = _get("https://remotive.com/api/remote-jobs?limit=200", accept="application/json")
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"[remotive] error: {e}", file=sys.stderr)
        return jobs

    for item in data.get("jobs", []):
        jobs.append(
            {
                "id": f"remotive:{item.get('id')}",
                "title": (item.get("title") or "").strip(),
                "company": (item.get("company_name") or "").strip(),
                "url": item.get("url") or "",
                "location": (item.get("candidate_required_location") or "Remote").strip(),
                "source": "Remotive",
                "posted_ts": _iso_ts(item.get("publication_date")),
                "tags": [item.get("category")] if item.get("category") else [],
            }
        )
    return jobs


def _strip_tags(s):
    import re

    return re.sub(r"<[^>]+>", "", s or "").strip()


def fetch_adzuna(adzuna_cfg):
    """Adzuna API — broad coverage incl. US ONSITE EHS/compliance roles.

    Requires a free app_id/app_key from https://developer.adzuna.com (2-min
    signup). Config shape:
        "adzuna": {
          "app_id": "...", "app_key": "...", "country": "us",
          "queries": ["EHS", "environmental health safety", "safety coordinator"]
        }
    """
    jobs = []
    app_id, app_key = adzuna_cfg.get("app_id"), adzuna_cfg.get("app_key")
    if not app_id or not app_key:
        return jobs
    country = adzuna_cfg.get("country", "us")
    max_days = adzuna_cfg.get("max_days_old", 30)
    for q in adzuna_cfg.get("queries", []):
        url = (
            f"https://api.adzuna.com/v1/api/jobs/{country}/search/1"
            f"?app_id={quote(app_id)}&app_key={quote(app_key)}"
            f"&results_per_page=50&what={quote(q)}&max_days_old={max_days}"
            f"&sort_by=date&content-type=application/json"
        )
        try:
            r = _get(url, accept="application/json")
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"[adzuna:{q}] error: {e}", file=sys.stderr)
            continue
        for item in data.get("results", []):
            jobs.append(
                {
                    "id": f"adzuna:{item.get('id')}",
                    "title": _strip_tags(item.get("title")),
                    "company": _strip_tags((item.get("company") or {}).get("display_name")),
                    "url": item.get("redirect_url") or "",
                    "location": ((item.get("location") or {}).get("display_name") or "").strip(),
                    "source": "Adzuna",
                    "posted_ts": _iso_ts(item.get("created")),
                    "tags": [
                        c
                        for c in [
                            item.get("contract_time"),
                            item.get("contract_type"),
                            (item.get("category") or {}).get("label"),
                        ]
                        if c
                    ],
                }
            )
        time.sleep(0.4)
    return jobs


def _amazon_ts(s):
    """Amazon posts dates as ISO or 'Month DD, YYYY'; return epoch or None."""
    ts = _iso_ts(s)
    if ts is not None:
        return ts
    if not s:
        return None
    try:
        dt = datetime.strptime(s.strip(), "%B %d, %Y").replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except Exception:
        return None


def fetch_amazon(amazon_cfg):
    """Amazon.jobs public search API — targets Amazon Global Safety (EHS) roles.

    Config shape:
        "amazon": { "enabled": true,
                    "queries": ["EHS", "environmental health safety", "safety"] }
    Mostly onsite roles (so the remote-only location filter trims heavily), but
    catches Amazon's remote WHS/safety-data/analytics postings.
    """
    jobs = []
    if not amazon_cfg or not amazon_cfg.get("enabled"):
        return jobs
    for q in amazon_cfg.get("queries", ["EHS"]):
        url = (
            "https://www.amazon.jobs/en/search.json?radius=24km&sort=recent"
            f"&result_limit=100&base_query={quote(q)}"
        )
        try:
            r = _get(url, accept="application/json")
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"[amazon:{q}] error: {e}", file=sys.stderr)
            continue
        for it in data.get("jobs", []):
            path = it.get("job_path") or ""
            team = it.get("team")
            team_label = team.get("label") if isinstance(team, dict) else None
            jobs.append(
                {
                    "id": f"amazon:{it.get('id_icims') or it.get('id')}",
                    "title": (it.get("title") or "").strip(),
                    "company": "Amazon",
                    "url": ("https://www.amazon.jobs" + path) if path else "",
                    "location": (it.get("normalized_location") or "").strip(),
                    "source": "Amazon",
                    "posted_ts": _amazon_ts(it.get("posted_date")),
                    "tags": [t for t in [it.get("job_schedule_type"), team_label] if t],
                }
            )
        time.sleep(0.5)
    return jobs


def gather_jobs(cfg):
    src = cfg["sources"]
    jobs = []
    if src.get("remoteok"):
        jobs += fetch_remoteok()
    if src.get("weworkremotely"):
        jobs += fetch_weworkremotely()
    if src.get("remotive"):
        jobs += fetch_remotive()
    for board in src.get("greenhouse_boards", []):
        jobs += fetch_greenhouse(board)
        time.sleep(0.4)
    for board in src.get("lever_boards", []):
        jobs += fetch_lever(board)
        time.sleep(0.4)
    for board in src.get("ashby_boards", []):
        jobs += fetch_ashby(board)
        time.sleep(0.4)
    if src.get("adzuna", {}).get("app_id"):
        jobs += fetch_adzuna(src["adzuna"])
    if src.get("amazon", {}).get("enabled"):
        jobs += fetch_amazon(src["amazon"])
    return jobs


# --------------------------------------------------------------------------- #
# Filtering
# --------------------------------------------------------------------------- #
# Terms that identify contract / freelance / consulting / 1099 arrangements.
CONTRACT_TERMS = [
    "contract",
    "contractor",
    "consultant",
    "consulting",
    "1099",
    "hourly",
    "freelance",
    "c2c",
    "corp-to-corp",
    "temporary",
    "temp-to-hire",
]


def is_contract(job):
    """True if the role looks like a contract/freelance/consulting engagement."""
    hay = " ".join(
        [job.get("title", ""), " ".join(job.get("tags", []))]
    ).lower()
    return any(t in hay for t in CONTRACT_TERMS)


def within_age(job, max_age_days, keep_undated=True):
    """True if the job is recent enough to keep."""
    if not max_age_days:
        return True
    ts = job.get("posted_ts")
    if ts is None:
        return keep_undated
    age_days = (time.time() - ts) / 86400.0
    return age_days <= max_age_days


def age_label(job):
    ts = job.get("posted_ts")
    if ts is None:
        return "date unknown"
    days = int((time.time() - ts) // 86400)
    if days <= 0:
        return "today"
    if days == 1:
        return "1 day ago"
    return f"{days} days ago"


def matches(job, keywords, location_filter):
    if keywords:
        haystack = " ".join(
            [job["title"], job["company"], " ".join(job.get("tags", []))]
        ).lower()
        if not any(kw.lower() in haystack for kw in keywords):
            return False
    if location_filter:
        loc = job.get("location", "").lower()
        # A role counts as remote if the location OR the title says so (many
        # aggregators report a geographic location but flag "Remote" in title).
        combined = loc + " " + job.get("title", "").lower()
        # Guard against negations like "Not a Remote Position" / "no remote".
        negations = ("not a remote", "not remote", "no remote", "non-remote",
                     "onsite only", "on-site only", "in-office only")
        if any(n in combined for n in negations):
            return False
        has_term = any(term.lower() in combined for term in location_filter)
        # Keep if location is unknown/blank, or a wanted term appears.
        if loc and not has_term:
            return False
    return True


# --------------------------------------------------------------------------- #
# Discord
# --------------------------------------------------------------------------- #
def build_embed(job):
    contract = is_contract(job)
    desc = f"**Company:** {job['company'] or 'Unknown'}"
    if contract:
        desc += "\n💼 **Looks like a contract / consulting role**"
    fields = []
    if job.get("location"):
        fields.append({"name": "Location", "value": job["location"][:100], "inline": True})
    fields.append({"name": "Posted", "value": age_label(job), "inline": True})
    if job.get("url"):
        fields.append({"name": "Action", "value": f"[Apply Now]({job['url']})", "inline": True})
    tags = job.get("tags") or []
    if tags:
        fields.append(
            {"name": "Tags", "value": ", ".join(tags[:6])[:200], "inline": False}
        )
    title = (job["title"] or "New role")[:248]
    return {
        "title": ("💼 " + title) if contract else title,
        "description": desc,
        "url": job.get("url") or None,
        "color": 15105570 if contract else DISCORD_EMBED_COLOR,  # amber for contract
        "fields": fields,
        "footer": {"text": f"via {job['source']}"},
    }


def post_embeds(webhook, embeds):
    """POST a batch (<=10) of embeds, honoring Discord rate limits."""
    for attempt in range(5):
        resp = requests.post(webhook, json={"embeds": embeds}, timeout=REQUEST_TIMEOUT)
        if resp.status_code in (200, 204):
            return True
        if resp.status_code == 429:
            retry_after = 1.0
            try:
                retry_after = float(resp.json().get("retry_after", 1.0))
            except Exception:
                pass
            print(f"[discord] rate limited; waiting {retry_after:.1f}s", file=sys.stderr)
            time.sleep(retry_after + 0.25)
            continue
        print(f"[discord] error {resp.status_code}: {resp.text[:200]}", file=sys.stderr)
        return False
    return False


def notify(webhook, jobs, dry_run=False):
    for i in range(0, len(jobs), MAX_EMBEDS_PER_MESSAGE):
        batch = jobs[i : i + MAX_EMBEDS_PER_MESSAGE]
        embeds = [build_embed(j) for j in batch]
        if dry_run:
            for j in batch:
                mark = " [CONTRACT]" if is_contract(j) else ""
                print(
                    f"  WOULD SEND:{mark} [{j['source']}] {j['title']} @ {j['company']} "
                    f"({age_label(j)})"
                )
            continue
        post_embeds(webhook, embeds)
        time.sleep(1.0)  # gentle pacing between batches


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    parser = argparse.ArgumentParser(description="Stream new job listings to Discord.")
    parser.add_argument(
        "--seed",
        action="store_true",
        help="Mark all current listings as seen without notifying (first-run setup).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be sent; do not post to Discord.",
    )
    args = parser.parse_args()

    cfg = load_config()
    webhook = cfg["_webhook"]
    if not args.dry_run and not args.seed and not webhook:
        sys.exit(
            "No Discord webhook. Set DISCORD_WEBHOOK_URL env var or "
            "discord_webhook_url in config.json."
        )

    seen = load_seen()
    all_jobs = gather_jobs(cfg)

    # De-duplicate within this run by id (sources can overlap).
    by_id = {}
    for j in all_jobs:
        by_id.setdefault(j["id"], j)
    all_jobs = list(by_id.values())

    # Collapse the same role posted under multiple locations/ids (e.g. one
    # Adzuna listing repeated per city) by (title, company).
    by_role = {}
    for j in all_jobs:
        key = (j["title"].strip().lower(), (j["company"] or "").strip().lower())
        by_role.setdefault(key, j)
    all_jobs = list(by_role.values())

    kw_matched = [
        j for j in all_jobs if matches(j, cfg["keywords"], cfg["location_filter"])
    ]
    matched = [
        j
        for j in kw_matched
        if within_age(j, cfg["max_age_days"], cfg["keep_undated"])
    ]
    aged_out = len(kw_matched) - len(matched)
    new_jobs = [j for j in matched if j["id"] not in seen]

    age_note = (
        f" (dropped {aged_out} older than {cfg['max_age_days']}d)" if aged_out else ""
    )
    print(
        f"Fetched {len(all_jobs)} jobs | {len(kw_matched)} match keywords | "
        f"{len(matched)} recent{age_note} | {len(new_jobs)} new"
    )

    if args.seed:
        for j in matched:
            seen.add(j["id"])
        save_seen(seen)
        print(f"Seeded {len(matched)} listings as seen. No notifications sent.")
        return

    if not new_jobs:
        print("No new matching jobs. Done.")
        return

    to_send = new_jobs[:MAX_NOTIFICATIONS_PER_RUN]
    if len(new_jobs) > MAX_NOTIFICATIONS_PER_RUN:
        print(
            f"Capping this run at {MAX_NOTIFICATIONS_PER_RUN} notifications "
            f"({len(new_jobs) - MAX_NOTIFICATIONS_PER_RUN} more will send next run)."
        )

    notify(webhook, to_send, dry_run=args.dry_run)

    if not args.dry_run:
        for j in to_send:
            seen.add(j["id"])
        save_seen(seen)
        queue_for_application(to_send)
        print(f"Sent {len(to_send)} notifications; state saved.")
    else:
        print(f"Dry run complete ({len(new_jobs)} new jobs would have been sent).")


if __name__ == "__main__":
    main()
