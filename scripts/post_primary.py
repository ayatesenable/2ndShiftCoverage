#!/usr/bin/env python3
"""
post_primary.py — announce TOMORROW's second-shift coverage to a Microsoft Teams chat.

Reads the calendar's `published` Firebase node (the name-resolved, id-free feed the app
writes on every save/open), looks up tomorrow (America/New_York), and posts an Adaptive
Card to a Power Automate "Workflows" webhook. Runs on a short cron so it can also RE-POST
an update if tomorrow's coverage changes after the first announcement — dedup is handled by
a tiny state node in Firebase, so rapid edits between polls collapse into at most one post.

Stdlib only (urllib + zoneinfo) — no pip install needed in CI.

Configuration (all via environment / GitHub secrets; sensible defaults baked in):
  TEAMS_WEBHOOK_URL   (required)  Power Automate Workflows webhook URL. Store as a secret.
  FIREBASE_DB_URL     (optional)  RTDB base URL. Default matches the deployed calendar.
  FIREBASE_BASE_PATH  (optional)  Parent path of the calendar data. Default "second-shift-calendar".
  FIREBASE_TOKEN      (optional)  Appended as ?auth=... for reads/writes IF your DB rules
                                  require auth. Leave unset if the published path is world-readable.
  ANNOUNCE_TZ         (optional)  IANA tz for "tomorrow". Default "America/New_York".
  POST_HOLIDAYS       (optional)  "1" to also post on no-coverage-needed days. Default off.
  DRY_RUN             (optional)  "1" prints the payload instead of posting (safe testing).
  DATE_OVERRIDE       (optional)  "YYYY-MM-DD" to target a specific day (manual testing).
"""

import json
import os
import sys
import urllib.request
import urllib.error
from datetime import datetime, timedelta

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover  (Python <3.9)
    print("Python 3.9+ required (zoneinfo).", file=sys.stderr)
    sys.exit(1)

# ---- config -----------------------------------------------------------------
DB_URL    = os.environ.get("FIREBASE_DB_URL", "https://secondshift-calendar-default-rtdb.firebaseio.com").rstrip("/")
BASE_PATH = os.environ.get("FIREBASE_BASE_PATH", "second-shift-calendar").strip("/")
TOKEN     = os.environ.get("FIREBASE_TOKEN", "").strip()
WEBHOOK   = os.environ.get("TEAMS_WEBHOOK_URL", "").strip()
TZ        = os.environ.get("ANNOUNCE_TZ", "America/New_York")
POST_HOLIDAYS = os.environ.get("POST_HOLIDAYS", "") == "1"
DRY_RUN   = os.environ.get("DRY_RUN", "") == "1"
DATE_OVERRIDE = os.environ.get("DATE_OVERRIDE", "").strip()

WEEKDAY_FULL = {"Mon": "Monday", "Tue": "Tuesday", "Wed": "Wednesday",
                "Thu": "Thursday", "Fri": "Friday", "Sat": "Saturday", "Sun": "Sunday"}
MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ---- firebase REST ----------------------------------------------------------
def _auth(url):
    return f"{url}?auth={TOKEN}" if TOKEN else url

def fb_get(path):
    url = _auth(f"{DB_URL}/{path}.json")
    with urllib.request.urlopen(url, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))

def fb_put(path, value):
    url = _auth(f"{DB_URL}/{path}.json")
    body = json.dumps(value).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="PUT",
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return r.status


# ---- date helpers -----------------------------------------------------------
def target_date():
    if DATE_OVERRIDE:
        return DATE_OVERRIDE
    now = datetime.now(ZoneInfo(TZ))
    return (now + timedelta(days=1)).date().isoformat()

def pretty(ds, dow):
    y, m, d = map(int, ds.split("-"))
    return f"{WEEKDAY_FULL.get(dow, dow)}, {MONTHS[m]} {d}"


# ---- message rendering ------------------------------------------------------
def signature(entry):
    """Stable fingerprint of what we'd announce — used to detect changes."""
    if entry is None:
        return "missing"
    t = entry.get("type")
    st = entry.get("status")
    if st == "no-coverage-needed":
        return "holiday"
    if t == "saturday":
        names = [a.get("name") for a in (entry.get("available") or [])]
        return "sat:" + ",".join(sorted(n for n in names if n))
    parts = []
    for sh in sorted((entry.get("shifts") or {}).keys()):
        parts.append(f"{sh}:{(entry['shifts'][sh] or {}).get('primary') or '-'}")
    return f"{t}:{st}:" + "|".join(parts)

def should_post(entry):
    """Only speak up when there's something worth saying about tomorrow."""
    if entry is None:
        return False
    st = entry.get("status")
    t = entry.get("type")
    if st == "no-coverage-needed":
        return POST_HOLIDAYS
    if t == "saturday":
        return bool(entry.get("available"))          # sign-ups exist
    if t == "sunday":
        return st != "no-coverage"                   # someone claimed it
    return True                                      # weekday: always (covered OR uncovered alert)

def render(entry, updated=False):
    """Return (adaptive_card_dict, plain_text_fallback)."""
    ds = entry["date"]; dow = entry.get("dow", "")
    title = pretty(ds, dow)
    tag = "  ·  UPDATED" if updated else ""
    t = entry.get("type"); st = entry.get("status")
    body = [{"type": "TextBlock", "size": "Large", "weight": "Bolder",
             "text": f"Second-shift coverage — {title}{tag}", "wrap": True}]
    lines = []
    alert = False

    if st == "no-coverage-needed":
        body.append({"type": "TextBlock", "wrap": True, "spacing": "Small",
                     "text": "No coverage needed (holiday / shutdown)."})
        lines.append("No coverage needed (holiday / shutdown).")

    elif t == "saturday":
        avail = entry.get("available") or []
        body.append({"type": "TextBlock", "wrap": True, "spacing": "Small",
                     "text": "Saturday open availability:"})
        facts = []
        for a in avail:
            win = f"{a.get('fromLabel','')}\u2013{a.get('toLabel','')}"
            facts.append({"title": a.get("name", "?"), "value": win})
            lines.append(f"{a.get('name','?')}: {win}")
        if facts:
            body.append({"type": "FactSet", "facts": facts})

    else:  # weekday or sunday — shifts map
        shifts = entry.get("shifts") or {}
        facts = []
        for sh in sorted(shifts.keys()):
            s = shifts[sh] or {}
            label = f"{s.get('label', sh)} shift"
            p = s.get("primary")
            if p:
                lock = "  \U0001F512" if s.get("locked") else ""
                facts.append({"title": label, "value": f"\u2605 {p}{lock}"})
                lines.append(f"{label}: {p}")
            else:
                alert = True
                facts.append({"title": label, "value": "\u26A0 OPEN — no primary"})
                lines.append(f"{label}: OPEN — no primary")
        body.append({"type": "FactSet", "facts": facts})
        if alert:
            body.insert(1, {"type": "TextBlock", "wrap": True, "color": "Attention",
                            "weight": "Bolder", "spacing": "Small",
                            "text": "\u26A0 A shift is uncovered for tomorrow."})

    card = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard", "version": "1.4", "body": body,
            },
        }],
    }
    plain = f"Second-shift coverage — {title}{tag}: " + "; ".join(lines)
    return card, plain


def post_to_teams(card, plain):
    if DRY_RUN or not WEBHOOK:
        where = "DRY_RUN" if DRY_RUN else "no TEAMS_WEBHOOK_URL set"
        print(f"[{where}] would post:\n{plain}\n{json.dumps(card, indent=2)}")
        return True
    body = json.dumps(card).encode("utf-8")
    req = urllib.request.Request(WEBHOOK, data=body, method="POST",
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            print(f"Posted to Teams (HTTP {r.status}): {plain}")
            return True
    except urllib.error.HTTPError as e:
        print(f"Teams post failed HTTP {e.code}: {e.read().decode('utf-8', 'ignore')}", file=sys.stderr)
        return False


# ---- main -------------------------------------------------------------------
def main():
    if not WEBHOOK and not DRY_RUN:
        print("TEAMS_WEBHOOK_URL is not set. Set the secret, or run with DRY_RUN=1 to test.", file=sys.stderr)
        sys.exit(1)

    ds = target_date()
    try:
        entry = fb_get(f"{BASE_PATH}/published/days/{ds}")
    except Exception as e:  # noqa: BLE001
        print(f"Could not read published feed for {ds}: {e}", file=sys.stderr)
        sys.exit(1)

    if entry is None:
        print(f"No published entry for {ds} (outside current+next month, or feed not generated yet). Nothing to do.")
        return

    sig = signature(entry)
    speak = should_post(entry)

    # load dedup state: {date, sig, posted}
    try:
        state = fb_get(f"{BASE_PATH}/announce") or {}
    except Exception:  # noqa: BLE001
        state = {}
    prev_date = state.get("date")
    prev_sig = state.get("sig")
    prev_posted = bool(state.get("posted"))

    if prev_date != ds:
        do_post, updated = speak, False                      # first look at a new tomorrow
    elif prev_sig != sig:
        do_post, updated = (speak or prev_posted), True       # changed since last time (incl. releases)
    else:
        print(f"No change for {ds} (sig unchanged). Nothing to post.")
        return

    posted_ok = False
    if do_post:
        card, plain = render(entry, updated=updated)
        posted_ok = post_to_teams(card, plain)
    else:
        print(f"{ds}: nothing worth announcing (status={entry.get('status')}). Recording state, staying quiet.")

    # record state so we don't repeat (skip the write on a failed post so it retries next run)
    if not DRY_RUN and (posted_ok or not do_post):
        try:
            fb_put(f"{BASE_PATH}/announce",
                   {"date": ds, "sig": sig, "posted": bool(posted_ok),
                    "updatedAt": datetime.now(ZoneInfo(TZ)).isoformat()})
        except Exception as e:  # noqa: BLE001
            print(f"Warning: could not write announce state: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
