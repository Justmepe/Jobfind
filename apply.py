#!/usr/bin/env python3
"""
apply.py — generate a tailored cover-letter .docx for a job, ready to review & send.

This is the "draft for review" helper: it does NOT send anything to employers.
It writes a customized cover letter to the applications/ folder (and can email
the package to YOU for review when SMTP is configured in profile.json).

Tailoring is template-based (no API): it detects the job's flavor from the
title and selects the most relevant achievements from profile.json.

Usage:
    python apply.py --title "Senior Systems Administrator, EHS Systems" \
                    --company "Acme" --url "https://..."
    python apply.py --from-queue          # generate for every job the pipeline
                                          #   queued in applications_queue.json
    python apply.py --from-queue --email  # also email the packages to yourself
"""

import argparse
import json
import os
import re
import smtplib
import sys
from datetime import date
from email.message import EmailMessage

import requests
from docx import Document
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_FILE = os.path.join(HERE, "profile.json")
QUEUE_FILE = os.path.join(HERE, "applications_queue.json")
OUT_DIR = os.path.join(HERE, "applications")

# Map a detected job flavor to which achievement buckets to feature, in order.
FLAVOR_BUCKETS = {
    "ai": ["ai", "data", "powerplatform"],
    "systems": ["systems", "powerplatform", "data"],
    "data": ["data", "powerplatform", "ehs"],
    "powerplatform": ["powerplatform", "data", "systems"],
    "ehs": ["ehs", "powerplatform", "data"],
}

FLAVOR_PATTERNS = [
    ("ai", r"\b(ai|a\.i|ml|machine learning|llm|prompt|generative|annotat|rlhf|model)\b"),
    ("systems", r"\b(systems? admin|administrator|sharepoint|power platform|implementation)\b"),
    ("data", r"\b(data analyst|business intelligence|power bi|analytics|reporting)\b"),
    ("powerplatform", r"\b(power apps|power automate|low.?code|automation)\b"),
    ("ehs", r"\b(ehs|hse|safety|environmental|occupational|compliance|regulatory)\b"),
]

FLAVOR_FRAMING = {
    "ai": "the analytical rigor of a chemist with hands-on experience building the data pipelines and structured schemas that AI systems depend on",
    "systems": "deep, hands-on ownership of enterprise low-code systems (SharePoint, Power Apps, Power Automate) in a regulated, multi-facility environment",
    "data": "the ability to turn messy operational data into clean, auditable business-intelligence assets that leaders actually use",
    "powerplatform": "practical, production-tested Power Platform engineering delivered against real regulatory and operational requirements",
    "ehs": "a working command of OSHA, PSM, and EPA frameworks combined with the technical skill to digitize and automate them",
}


def load_profile():
    with open(PROFILE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def detect_flavor(title):
    t = (title or "").lower()
    for flavor, pat in FLAVOR_PATTERNS:
        if re.search(pat, t):
            return flavor
    return "ehs"


def _slug(s):
    return re.sub(r"[^a-zA-Z0-9]+", "-", (s or "").strip()).strip("-")[:60] or "role"


def build_cover_letter(profile, title, company):
    flavor = detect_flavor(title)
    buckets = FLAVOR_BUCKETS[flavor]
    picks = []
    for b in buckets:
        items = profile["achievements"].get(b, [])
        if items:
            picks.append(items[0])
        if len(picks) >= 3:
            break

    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    # Header — name + contact
    h = doc.add_paragraph()
    r = h.add_run(profile["name"])
    r.bold = True
    r.font.size = Pt(15)
    contact_bits = [profile.get("location"), profile.get("email"),
                    profile.get("phone"), profile.get("linkedin")]
    contact = "  |  ".join([c for c in contact_bits if c])
    doc.add_paragraph(contact)
    doc.add_paragraph(date.today().strftime("%B %d, %Y"))
    doc.add_paragraph()

    company = company or "your team"
    doc.add_paragraph(f"Dear {company} Hiring Team,")

    intro = (
        f"I am writing to apply for the {title} position at {company}. "
        f"As {profile['current_role']}, I am {profile['one_liner']}. "
        f"What I would bring to this role is {FLAVOR_FRAMING[flavor]}."
    )
    doc.add_paragraph(intro)

    doc.add_paragraph(
        "A few examples of the work I have contributed to that map directly to this role:"
    )
    for p in picks:
        b = doc.add_paragraph(style="List Bullet")
        # Capitalize first letter of the bullet.
        b.add_run(p[0].upper() + p[1:] + ".")

    close = (
        f"I am excited by the opportunity at {company} and would welcome the chance to "
        f"discuss how my mix of regulatory knowledge, systems-building, and data skills "
        f"can contribute. Thank you for your time and consideration."
    )
    doc.add_paragraph(close)
    doc.add_paragraph()
    doc.add_paragraph("Sincerely,")
    doc.add_paragraph(profile["name"])

    os.makedirs(OUT_DIR, exist_ok=True)
    fname = f"CoverLetter_{_slug(company)}_{_slug(title)}.docx"
    path = os.path.join(OUT_DIR, fname)
    doc.save(path)
    return path, flavor


def email_package(profile, jobs_done):
    """Email the generated packages to yourself for review (optional)."""
    smtp = profile.get("smtp") or {}
    if not smtp.get("app_password"):
        print("  (skip email: no smtp.app_password in profile.json)")
        return
    msg = EmailMessage()
    msg["Subject"] = f"[jobwatch] {len(jobs_done)} application drafts ready to review"
    msg["From"] = profile["email"]
    msg["To"] = profile["email"]
    lines = ["Your tailored cover-letter drafts are attached. Review, then send:\n"]
    for j in jobs_done:
        lines.append(f"- {j['title']} @ {j['company']}\n  {j.get('url','')}")
    msg.set_content("\n".join(lines))
    for j in jobs_done:
        with open(j["path"], "rb") as f:
            data = f.read()
        msg.add_attachment(
            data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(j["path"]),
        )
    # Attach the resume too, if present.
    resume = os.path.join(os.path.dirname(HERE), profile.get("resume_file", ""))
    if profile.get("resume_file") and os.path.exists(resume):
        with open(resume, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename=profile["resume_file"],
            )
    host = smtp.get("host", "smtp.gmail.com")
    port = int(smtp.get("port", 465))
    with smtplib.SMTP_SSL(host, port) as s:
        s.login(profile["email"], smtp["app_password"])
        s.send_message(msg)
    print(f"  emailed {len(jobs_done)} draft(s) to {profile['email']}")


def notify_discord(profile, jobs_done):
    """Post a heads-up to Discord that drafts are waiting in the inbox."""
    cfg_file = os.path.join(HERE, "config.json")
    webhook = os.environ.get("DISCORD_WEBHOOK_URL")
    if not webhook and os.path.exists(cfg_file):
        try:
            webhook = json.load(open(cfg_file, encoding="utf-8")).get("discord_webhook_url")
        except Exception:
            webhook = None
    if not webhook:
        return
    lines = "\n".join(f"• {j['title']} @ {j['company']}" for j in jobs_done[:15])
    embed = {
        "title": f"📄 {len(jobs_done)} application draft(s) ready to review",
        "description": f"Emailed to **{profile['email']}** — review and send:\n{lines}",
        "color": 15844367,
    }
    try:
        requests.post(webhook, json={"embeds": [embed]}, timeout=20)
    except Exception as e:
        print(f"  (discord notify failed: {e})")


def main():
    ap = argparse.ArgumentParser(description="Generate tailored cover-letter drafts.")
    ap.add_argument("--title")
    ap.add_argument("--company")
    ap.add_argument("--url", default="")
    ap.add_argument("--from-queue", action="store_true",
                    help="Generate for every job in applications_queue.json")
    ap.add_argument("--email", action="store_true",
                    help="Also email the packages to yourself (needs smtp in profile.json)")
    args = ap.parse_args()

    profile = load_profile()

    if args.from_queue:
        if not os.path.exists(QUEUE_FILE):
            sys.exit("No applications_queue.json yet — run the pipeline first.")
        with open(QUEUE_FILE, "r", encoding="utf-8") as f:
            jobs = json.load(f)
    elif args.title:
        jobs = [{"title": args.title, "company": args.company or "", "url": args.url}]
    else:
        sys.exit("Provide --title/--company or --from-queue.")

    done = []
    for j in jobs:
        path, flavor = build_cover_letter(profile, j["title"], j.get("company"))
        print(f"[{flavor:12}] {os.path.basename(path)}")
        done.append({**j, "path": path})

    if args.email and done:
        email_package(profile, done)
        notify_discord(profile, done)
        # Clear the queue after a successful email run so drafts aren't re-sent.
        if args.from_queue and (profile.get("smtp") or {}).get("app_password"):
            with open(QUEUE_FILE, "w", encoding="utf-8") as f:
                json.dump([], f)
            print("  queue cleared")

    print(f"\nGenerated {len(done)} cover letter(s) in {OUT_DIR}")


if __name__ == "__main__":
    main()
