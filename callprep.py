#!/usr/bin/env python3
"""
callprep.py - customer case review from a ServiceNow export.

  python3 callprep.py acme              walk each case, type your status updates, build the page
  python3 callprep.py acme --no-update  rebuild the page from the updates you already have
  python3 callprep.py acme --by-priority  each section in plain priority order
                                        (default: split by whose move it is, then priority)
  python3 callprep.py                   pick an account from a list
  python3 callprep.py across            internal: everything waiting on us, every account

Exports go in exports/<account>/ (e.g. exports/acme/acme_tickets-10-8.xlsx).
The newest file is today's review; the one before it is what "new" and
"resolved" are measured against.

Your status updates live in notes/<account>.md - plain text, newest first,
safe to edit by hand. The page shows the top line for each case.

Output lands in out/<account>/ as .html (open, copy, paste into OneNote) and .csv.
Python 3 standard library only.
"""

import csv
import datetime as dt
import os
import re
import select
import shutil
import subprocess
import sys
import textwrap
import zipfile
from html import escape, unescape
from xml.etree import ElementTree as ET

try:
    import readline  # noqa: F401 - arrow keys and line editing in the update prompt
    HAVE_READLINE = True
except ImportError:
    HAVE_READLINE = False

ROOT = os.path.dirname(os.path.abspath(__file__))
EXPORT_DIR = os.path.join(ROOT, "exports")
NOTES_DIR = os.path.join(ROOT, "notes")
OUT_DIR = os.path.join(ROOT, "out")
EXTS = (".xlsx", ".xlsm", ".csv")


def set_root(path):
    """Point exports/, notes/ and out/ at another folder (the desktop app keeps its data there)."""
    global ROOT, EXPORT_DIR, NOTES_DIR, OUT_DIR
    ROOT = os.path.abspath(path)
    EXPORT_DIR = os.path.join(ROOT, "exports")
    NOTES_DIR = os.path.join(ROOT, "notes")
    OUT_DIR = os.path.join(ROOT, "out")


PROFILE_KEYS = {"vendor": "VENDOR", "jira_browse_url": "JIRA_BROWSE_URL", "default_type": "DEFAULT_TYPE",
                "types": "TYPES", "type_aliases": "TYPE_ALIASES", "type_columns": "TYPE_COLUMNS",
                "type_column_label": "TYPE_COLUMN_LABEL", "topics": "TOPICS",
                "self_hosted_products": "SELF_HOSTED_PRODUCTS", "version_label": "VERSION_LABEL",
                "show_tenant_on_page": "PAGE_SHOW_TENANT", "show_area_on_page": "PAGE_SHOW_AREA",
                "show_engineering_on_page": "PAGE_SHOW_ENGINEERING", "show_group_on_page": "SHOW_GROUP_ON_PAGE"}
PROFILE_NAME = None


def apply_profile(profile):
    """Use an organization profile (a dict, see profiles/example.json). Unknown keys are ignored."""
    global PROFILE_NAME
    g = globals()
    for key, var in PROFILE_KEYS.items():
        if key in profile:
            value = profile[key]
            if var == "TYPE_ALIASES":
                value = {k.lower(): v for k, v in value.items()}
            if var == "TYPE_COLUMNS":
                value = [" ".join(k.lower().split()) for k in value]
            g[var] = value
    PROFILE_NAME = profile.get("name") or "custom"


def load_profile(path=None):
    """Load the profile from path, $TAM_PROFILE, or profile.json next to this script. True if one was found."""
    import json
    for candidate in (path, os.environ.get("TAM_PROFILE"),
                      os.path.join(os.path.dirname(os.path.abspath(__file__)), "profile.json")):
        if candidate and os.path.isfile(candidate):
            with open(candidate, encoding="utf-8") as fh:
                apply_profile(json.load(fh))
            return True
    return False


# --- tunables ---------------------------------------------------------------
# Organization-specific settings (vendor name, case types, Jira, topics) come from a
# profile: profile.json next to this script, or the file named in $TAM_PROFILE.
# See profiles/example.json. The values below are the neutral defaults.
VENDOR = "Vendor"             # your company, as the customer sees it ("Action: Vendor")
DEFAULT_TYPE = "Support"      # a case with no request type is a Support ticket
# Page order and color for case types. Types not listed follow, alphabetically.
TYPES = {
    "Support": "#175cd3",
}
OTHER_TYPE_COLOR = "#475467"
# Values that should land in an existing section. A support tier ("Platinum Support",
# "Premium Support", ...) is a Support case, not a section of its own.
TYPE_ALIASES = {}
TYPE_TILES_MAX = 4            # more case types than this: one summary line instead of a box each
SHOW_GROUP_ON_PAGE = False    # assignment group is internal; Terminal and CSV always show it
AGING_DAYS = 60               # internal prep: open longer than this
AGE_WARN_DAYS = 30            # page: amber triangle on cases open this long
AGE_ALERT_DAYS = 90           # page: red triangle
AGE_CHART_MAX = 15            # page: bars in the case-age chart (oldest first)
IDLE_DAYS = 7                 # internal prep: ours, and no ServiceNow activity this long
FOLLOW_UP_SOON_DAYS = 2       # internal prep: follow-up date due within this many days
REASSIGN_WARN = 3             # internal prep: case reassigned this many times or more

# Optional columns shown on the customer page when the export has them.
# Everything else (tenant, follow-up, escalation, action status...) stays internal.
PAGE_SHOW_AREA = True         # Component / Sub-Component under the issue, and focus areas from them
PAGE_SHOW_ENGINEERING = True  # "Engineering" tag (with the Jira ID when known) when Has Jira? is yes
PAGE_SHOW_TENANT = False      # tenant name from Tenant URL, when the account has more than one
# Jira tickets link here (page, app). %s is the ticket, e.g. "https://acme.atlassian.net/browse/%s".
# Blank: tickets are shown without links.
JIRA_BROWSE_URL = ""

# A topic is called out on the page when it covers 3+ cases and 30%+ of the book.
# Used when the export has no Component / Sub-Component. Each entry is a list of
# word prefixes, so "cert" catches Cert, Certification(s). Example:
#   {"Certifications": ["cert", "recert", "campaign"], "Connectors": ["connector"]}
TOPICS = {}

# Products installed at the customer (versions matter more than components). Their cases
# show Product Version under the issue; for the others, Sub-Component is recommended.
SELF_HOSTED_PRODUCTS = []
TYPE_COLUMN_LABEL = "Request type"   # how the tip names the case-type column(s)
VERSION_LABEL = "product versions"

# Column names each field is read from, best first: a field takes the column that matches
# its earliest alias, wherever that column sits in the export, and a field whose columns
# are missing is left blank. Matching ignores case and extra spaces.
COLUMNS = {
    "number":      ["number", "case number", "case id", "case"],
    "priority":    ["highest priority", "priority", "case priority"],
    "description": ["short description", "subject", "summary", "case title", "description"],
    "details":     ["description", "detailed description"],
    "state":       ["state", "status", "case state"],
    "substate":    ["substate", "sub state", "sub-state"],
    "opened":      ["opened", "opened at", "created", "created on"],
    "updated":     ["updated", "last updated", "updated on", "updated at"],
    "contact":     ["customer contact", "contact"],
    "group":       ["assignment group", "group"],
    "account":     ["customer account", "account", "account name", "company"],
    "entitlement": ["entitlement", "support level", "name"],
    # "Name" is ambiguous in ServiceNow exports (owner, or a dot-walked entitlement
    # name) - map_columns() decides by looking at the values.
    "owner":       ["assigned to", "case owner", "owner", "name"],
    # optional extras
    "product":     ["product"],
    "component":   ["component"],
    "subcomponent": ["sub-component", "sub component", "subcomponent"],
    "version":     ["product version", "version"],
    "project":     ["project code", "name of project working on"],
    "has_jira":    ["has jira?", "has jira"],
    "bug":         ["pbi or bug number", "jira id", "jira ticket", "jira", "bug"],
    "bug_status":  ["bug status latest", "jira status"],
    "escalation":  ["escalation", "escalation level"],
    "escalated":   ["escalated?", "escalated", "active escalation", "active account escalation"],
    "action_status": ["action status"],
    "needs_attention": ["needs attention"],
    "follow_up":   ["follow up", "follow-up date", "follow up date", "follow up on"],
    "assignee_updated": ["assignee updated"],
    "non_assignee_updated": ["non-assignee updated"],
    "non_assignee_by": ["non-assignee updated by"],
    "tenant":      ["tenant url", "tenant", "tenant name"],
    "impact":      ["impact"],
    "urgency":     ["urgency"],
    "reassignments": ["reassignment count"],
    "sla_due":     ["sla due"],
}
# Case-type columns: every one exported is read, first non-blank value wins (blank = Support).
TYPE_COLUMNS = ["request type"]     # the profile lists your instance's request-type columns
# Comment columns: all are read and merged by date; an entry in several columns counts once.
JOURNAL_COLUMNS = ["comments and work notes", "work notes and comments", "additional comments",
                   "work notes", "comments", "activity"]
TIERS = {"platinum", "gold", "silver", "bronze", "premium", "premier", "standard",
         "enterprise", "basic", "essentials", "elite"}
CLOSED_WORDS = ("resolved", "closed", "cancel")
PRIORITY_RANK = {"P1": 1, "P2": 2, "P3": 3, "P4": 4, "P5": 5}
PRIORITY_COLORS = {"P1": ("#b42318", "#fee4e2"), "P2": ("#b54708", "#fef0c7"),
                   "P3": ("#175cd3", "#e0eaff"), "P4": ("#475467", "#eaecf0"),
                   "P5": ("#475467", "#eaecf0")}

# --- reading exports --------------------------------------------------------
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def _col_index(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    if not m:
        return None
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx(path):
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
    sheets = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    sheet = ET.fromstring(z.read(sheets[0]))
    rows = []
    for row in sheet.iter(NS + "row"):
        cells, pos = [], 0
        for c in row.findall(NS + "c"):
            i = _col_index(c.get("r"))
            if i is None:
                i = pos
            while len(cells) < i:
                cells.append("")
            t, v, inline = c.get("t"), c.find(NS + "v"), c.find(NS + "is")
            if inline is not None:
                val = "".join(x.text or "" for x in inline.iter(NS + "t"))
            elif v is None:
                val = ""
            elif t == "s":
                val = shared[int(v.text)]
            else:
                val = v.text or ""
            cells.append(val)
            pos = i + 1
        rows.append(cells)
    return rows


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return list(csv.reader(fh))


def parse_date(value):
    """xlsx exports carry Excel serial numbers; csv exports carry strings."""
    if not value:
        return None
    try:
        f = float(value)
        if 20000 < f < 90000:
            return EXCEL_EPOCH + dt.timedelta(days=f)
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y",
                "%m-%d-%Y %H:%M:%S", "%m-%d-%Y %H:%M", "%m-%d-%Y"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def nice_name(value):
    """'jane q public' -> 'Jane Q Public', "O'brien" -> "O'Brien"."""
    text = " ".join(str(value or "").split())
    if text.lower() in ("", "no contact", "none", "n/a", "-"):
        return ""

    def fix(word):
        if word.islower() or word.isupper():
            word = word.capitalize()
        return re.sub(r"(?<=['\-])([a-z])", lambda m: m.group(1).upper(), word)

    return " ".join(fix(w) for w in text.split(" "))


def norm_priority(value):
    m = re.search(r"([1-5])", str(value or ""))
    return "P%s" % m.group(1) if m else str(value or "").strip()


def _key(header):
    return " ".join(str(header).lower().split())


def is_type_column(key):
    return key in TYPE_COLUMNS


def map_columns(header, rows):
    """{field: column index}. Each field takes the column matching its earliest alias."""
    keys = [_key(h) for h in header]

    def values(i):
        return [str(r[i]).strip().lower() for r in rows if i < len(r) and str(r[i]).strip()]

    mapping = {}
    for field, aliases in COLUMNS.items():
        for alias in aliases:
            hits = [i for i, k in enumerate(keys) if k == alias]
            if field in ("owner", "entitlement") and alias == "name" and hits:
                vals = values(hits[0])
                tier_like = vals and sum(v in TIERS for v in vals) >= 0.8 * len(vals)
                if tier_like != (field == "entitlement"):
                    continue
            if hits:
                mapping[field] = hits[0]
                break
    if "priority" not in mapping:      # "Case Priority (Customer)" and the like
        hit = next((i for i, k in enumerate(keys) if "priority" in k), None)
        if hit is not None:
            mapping["priority"] = hit
    if "type" not in mapping:
        hit = next((i for i, k in enumerate(keys) if is_type_column(k)), None)
        if hit is not None:
            mapping["type"] = hit
    return mapping


def truthy(value):
    """'1', 'true', 'Yes' -> True; '0', 'false', 'No' -> False; blank -> None."""
    s = str(value or "").strip().lower()
    if not s:
        return None
    if s in ("0", "0.0", "false", "no", "n", "none", "normal"):
        return False
    return True


def tenant_name(value):
    """'https://acme-sb.example.com/ui/...' -> 'acme-sb' (first part of the host); anything else as is."""
    m = re.match(r"(?:https?://)?([a-z0-9-]+)\.[a-z0-9.-]+\.[a-z]{2,}(?:[/:?#]|$)", str(value or "").strip(), re.I)
    return m.group(1).lower() if m else " ".join(str(value or "").split())


def to_int(value):
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


def load_export(path):
    rows = read_xlsx(path) if path.lower().endswith((".xlsx", ".xlsm")) else read_csv(path)
    if not rows:
        raise SystemExit("Export is empty: %s" % path)
    header, body = [str(h).strip() for h in rows[0]], rows[1:]
    m = map_columns(header, body)
    if "number" not in m:
        raise SystemExit("No case number column in %s. Columns found: %s"
                         % (os.path.basename(path), ", ".join(header)))

    def get(raw, field):
        i = m.get(field)
        return str(raw[i]) if i is not None and i < len(raw) and raw[i] is not None else ""

    keys = [_key(h) for h in header]
    journal_cols = [i for i, k in enumerate(keys) if k in JOURNAL_COLUMNS]
    type_cols = sorted((i for i, k in enumerate(keys) if is_type_column(k)),
                       key=lambda i: TYPE_COLUMNS.index(keys[i]) if keys[i] in TYPE_COLUMNS else -1)

    def case_type(raw):
        for i in type_cols:
            value = " ".join(str(raw[i] if i < len(raw) else "").split())
            if value:
                return normalize_type(value)
        return DEFAULT_TYPE

    if journal_cols:
        m["journal"] = journal_cols[0]
    if type_cols:
        m["type"] = type_cols[0]

    def text(raw, field):
        return " ".join(get(raw, field).split())

    cases = []
    for raw in body:
        number = get(raw, "number").strip().upper()
        if not number:
            continue
        description = text(raw, "description")
        if m.get("description") == m.get("details") and len(description) > 140:   # only a long Description
            description = condense(description, 140)
        cases.append({
            "number": number,
            "priority": norm_priority(get(raw, "priority")),
            "description": description,
            "state": " ".join(get(raw, "state").split()),
            "substate": " ".join(get(raw, "substate").split()),
            "contact": nice_name(get(raw, "contact")),
            "owner": nice_name(get(raw, "owner")),
            "group": " ".join(get(raw, "group").split()),
            "account": " ".join(get(raw, "account").split()),
            "opened_dt": parse_date(get(raw, "opened")),
            "updated_dt": parse_date(get(raw, "updated")),
            "type": case_type(raw),
            "journal": "\n\n".join(str(raw[i]) for i in journal_cols if i < len(raw) and raw[i]),
            # optional extras: blank / None when the export doesn't have them
            "details": get(raw, "details").strip() if m.get("details") != m.get("description") else "",
            "product": text(raw, "product"),
            "component": text(raw, "component"),
            "subcomponent": text(raw, "subcomponent"),
            "version": text(raw, "version"),
            "project": text(raw, "project"),
            "has_jira": truthy(get(raw, "has_jira")),
            "bug": text(raw, "bug"),
            "bug_status": text(raw, "bug_status"),
            "escalation": text(raw, "escalation"),
            "escalated": truthy(get(raw, "escalated")),
            "action_status": text(raw, "action_status"),
            "needs_attention": truthy(get(raw, "needs_attention")),
            "follow_up_dt": parse_date(get(raw, "follow_up")),
            "assignee_updated_dt": parse_date(get(raw, "assignee_updated")),
            "non_assignee_updated_dt": parse_date(get(raw, "non_assignee_updated")),
            "non_assignee_by": nice_name(get(raw, "non_assignee_by")),
            "tenant": tenant_name(get(raw, "tenant")),
            "impact": text(raw, "impact"),
            "urgency": text(raw, "urgency"),
            "reassignments": to_int(get(raw, "reassignments")),
            "sla_due_dt": parse_date(get(raw, "sla_due")),
        })
    return cases, m


# --- comments and work notes -----------------------------------------------
JOURNAL_HEAD = re.compile(r"^(\d\d-\d\d-\d{4} \d\d:\d\d:\d\d) - (.+?) "
                          r"\((Work notes|Additional comments)\)\s*$", re.M)


EMAIL_HEADER = re.compile(r"^(From|Sent|Date|To|Cc|Bcc|Subject)\s*:", re.I)
BANNERS = ("this message is from an external sender",
           "this message came from outside your organization")
SYSTEM_AUTHORS = {"system", "admin administrator"}


def _plain(text):
    """Journal HTML -> one readable paragraph: no tags, email headers, banners or long links."""
    text = re.sub(r"\[/?code\]", "\n", text)
    text = re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", text, flags=re.S | re.I)
    text = re.sub(r"[#.][\w-]+\s*\{[^}]*\}", " ", text)        # stray CSS rules from email banners
    text = re.sub(r"<li[^>]*>", "\n\u2022 ", text, flags=re.I)
    text = re.sub(r"<(br|/p|/li|/div)[^>]*>", "\n", text, flags=re.I)
    text = unescape(re.sub(r"<[^>]+>", "", text)).replace("\u00a0", " ").replace("\u202f", " ")
    kept, body_chars = [], 0
    for line in text.split("\n"):
        line = " ".join(line.split())
        if not line or line.lower().rstrip(".") in BANNERS:
            continue
        if EMAIL_HEADER.match(line):
            if line.lower().startswith("from") and body_chars > 40:
                kept.append("[earlier email trimmed]")   # a quoted reply chain starts here
                break
            continue
        kept.append(line)
        body_chars += len(line)
    text = " ".join(kept)
    # Long links -> just the host, so they don't eat the whole line.
    return re.sub(r"https?://([^/\s]+)\S*", lambda m: m.group(1) + "/\u2026", text)


def parse_journal(raw):
    """Newest-first list of {when, author, kind, text}. kind is 'internal' or 'customer'."""
    entries, heads, seen = [], list(JOURNAL_HEAD.finditer(raw or "")), set()
    for i, m in enumerate(heads):
        body = raw[m.end(): heads[i + 1].start() if i + 1 < len(heads) else len(raw)]
        key = (m.group(1), m.group(2).strip(), m.group(3), " ".join(body.split())[:200])
        if key in seen:      # same entry from "Comments and Work notes" and "Additional comments"
            continue
        seen.add(key)
        entries.append({
            "when": parse_date(m.group(1)),
            "author": m.group(2).strip(),
            "kind": "internal" if m.group(3) == "Work notes" else "customer",
            "text": _plain(body),
        })
    return sorted(entries, key=lambda e: e["when"] or dt.datetime.min, reverse=True)


TICKET = re.compile(r"\b([A-Z][A-Z0-9]{1,14}-\d+)\b")
# Look like Jira keys but aren't: UTF-8, ISO-8601, SHA-256, CVE-2024...
NOT_JIRA = {"UTF", "ISO", "SHA", "AES", "RSA", "RFC", "TLS", "SSL", "PKCS", "CVE", "HTTP", "HTTPS", "IPV",
            "MD", "PT", "GMT", "UTC", "COVID", "WIN", "X", "CS", "CSM", "INC", "CHG", "TASK", "KB", "PRB", "RITM"}


def is_ticket(text):
    m = TICKET.fullmatch((text or "").strip())
    return bool(m) and m.group(1).split("-")[0] not in NOT_JIRA


def jira_status(entries):
    """('ENG-12345', 'Waiting for Engineering') from the newest Jira status change."""
    for e in entries:
        m = re.search(r"changed from:\s*(.+?)\s+to\s+(.+?)\s+Jira Issue URL:\s*([A-Z][A-Z0-9]+-\d+)",
                      e["text"])
        if m:
            return m.group(3), m.group(2)
    for e in entries:
        for m in TICKET.finditer(e["text"]):
            if m.group(1).split("-")[0] not in NOT_JIRA:
                return m.group(1), ""
    return None


def jira_of(c):
    """(ticket, status) for a case. Looks in a bug-number column, then Jira links in the raw comments
    (the cleaned text shortens links and would lose the ticket), then the comments themselves.
    ('', '') when Has Jira? is yes but no ticket number is in the export; None when no Jira at all."""
    if is_ticket(c.get("bug")):
        found = (c["bug"].strip(), c.get("bug_status") or "")
    else:
        journal = c.get("journal") or ""
        found = jira_status(parse_journal(journal))
        link = re.search(r"atlassian\.net/browse/([A-Z][A-Z0-9]{1,14}-\d+)", journal)
        if link and not (found and found[1]):          # a status-change message also carries the status
            found = (link.group(1), "")
    if found:
        return found
    return ("", "") if c.get("has_jira") else None


def jira_url(ticket):
    return JIRA_BROWSE_URL % ticket if ticket and is_ticket(ticket) and JIRA_BROWSE_URL else None


# --- finding exports --------------------------------------------------------
def account_from(filename):
    return re.split(r"[_\-. ]", os.path.basename(filename))[0].lower()


def _infer_year(month, day, ref):
    d = dt.date(ref.year, month, day)
    return d if d <= ref else dt.date(ref.year - 1, month, day)


def export_date(path):
    """Date the export was taken: from the filename (10-8, 2026-10-08), else file time."""
    return name_date(path) or dt.date.fromtimestamp(os.path.getmtime(path))


def name_date(path):
    """The date in an export's filename (10-8, 2026-10-08), or None."""
    stem = os.path.splitext(os.path.basename(path))[0]
    today = dt.date.today()
    try:
        m = re.search(r"(20\d\d)[-_.](\d{1,2})[-_.](\d{1,2})", stem)
        if m:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.search(r"(\d{1,2})[-_.](\d{1,2})(?:[-_.](\d{2,4}))?(?!.*\d)", stem)
        if m:
            month, day, year = int(m.group(1)), int(m.group(2)), m.group(3)
            if year:
                return dt.date(int(year) + (2000 if len(year) == 2 else 0), month, day)
            return _infer_year(month, day, today)
    except ValueError:
        pass
    return None


def _is_export(path):
    name = os.path.basename(path)
    return os.path.isfile(path) and name.lower().endswith(EXTS) and not name.startswith(("~", "."))


def exports_for(account):
    paths = []
    folder = os.path.join(EXPORT_DIR, account)
    if os.path.isdir(folder):
        paths += [os.path.join(folder, n) for n in os.listdir(folder)]
    if os.path.isdir(EXPORT_DIR):  # older flat layout: exports/acme_tickets-9-22.xlsx
        paths += [os.path.join(EXPORT_DIR, n) for n in os.listdir(EXPORT_DIR)
                  if account_from(n) == account]
    paths = [p for p in paths if _is_export(p)]
    return sorted(paths, key=lambda p: (export_date(p), os.path.getmtime(p)))


def list_accounts():
    if not os.path.isdir(EXPORT_DIR):
        return []
    found = set()
    for name in os.listdir(EXPORT_DIR):
        path = os.path.join(EXPORT_DIR, name)
        if os.path.isdir(path) and any(_is_export(os.path.join(path, n)) for n in os.listdir(path)):
            found.add(name.lower())
        elif _is_export(path):
            found.add(account_from(name))
    return sorted(found)


# --- status notes -----------------------------------------------------------
def notes_path(account):
    return os.path.join(NOTES_DIR, "%s.md" % account)


def load_notes(account):
    """{case number: [entries, newest first]}"""
    path = notes_path(account)
    if not os.path.exists(path):
        return {}
    notes, current = {}, None
    for line in open(path, encoding="utf-8"):
        m = re.match(r"^##\s+(\S+)\s*$", line)
        if m:
            current = m.group(1).upper()
            notes.setdefault(current, [])
        elif current and line.strip():
            notes[current].append(line.strip())
    return {k: v for k, v in notes.items() if v}


def save_notes(account, notes, title):
    os.makedirs(NOTES_DIR, exist_ok=True)
    with open(notes_path(account), "w", encoding="utf-8") as fh:
        fh.write("# %s - case status updates\n" % title)
        fh.write("# One '## CASENUMBER' block per case, newest line first.\n")
        fh.write("# The review page shows the top line. Edit freely.\n\n")
        for number in sorted(notes, reverse=True):
            if notes[number]:
                fh.write("## %s\n%s\n\n" % (number, "\n".join(notes[number])))


def split_entry(entry):
    """'10/8 - text' -> (date(2026,10,8), 'text'). Undated lines -> (None, line)."""
    m = re.match(r"^(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\s*[-–:]\s*(.*)$", entry)
    if not m:
        return None, entry
    try:
        month, day, year = int(m.group(1)), int(m.group(2)), m.group(3)
        if year:
            when = dt.date(int(year) + (2000 if len(year) == 2 else 0), month, day)
        else:
            when = _infer_year(month, day, dt.date.today())
    except ValueError:
        return None, entry
    return when, m.group(4)


# --- shaping the book -------------------------------------------------------
def is_closed(case):
    return any(w in case["state"].lower() for w in CLOSED_WORDS)


def side(case):
    """'customer' if the next move is theirs, 'vendor' if ours, None if unknown."""
    s = case["state"].lower()
    if not s:
        return None
    if "customer" in s or "awaiting info" in s or "awaiting caller" in s:
        return "customer"
    return "vendor"


def sort_key(case):
    return (PRIORITY_RANK.get(case["priority"], 9), case["opened_dt"] or dt.datetime.max)


def _known_type(name):
    """Case-insensitive match against TYPES, so 'Rule review' and 'Rule Review' are one section."""
    return next((t for t in TYPES if t.lower() == name.lower()), None)


def normalize_type(value):
    low = value.lower()
    if low in TYPE_ALIASES:
        return _known_type(TYPE_ALIASES[low]) or TYPE_ALIASES[low]
    words = low.split()
    if words[-1] == "support" and all(w in TIERS for w in words[:-1]):   # "Platinum Support"
        return DEFAULT_TYPE
    return _known_type(value) or value


def type_rank(name):
    known = list(TYPES)
    t = _known_type(name)
    return (known.index(t), "") if t else (len(known), name.lower())


def type_color(name):
    return TYPES.get(_known_type(name) or "", OTHER_TYPE_COLOR)


ACTION_RANK = {"customer": 0, "vendor": 1, None: 2}   # --by-action: customer's moves first


def action_sort_key(case):
    return (ACTION_RANK[side(case)],) + sort_key(case)


def grouped(cases, by_action=False):
    """[(type, cases)] in TYPES order. Each section sorted by priority then age,
    or with by_action, by whose move it is first."""
    key = action_sort_key if by_action else sort_key
    types = sorted({c["type"] for c in cases}, key=type_rank)
    return [(t, sorted((c for c in cases if c["type"] == t), key=key)) for t in types]


def display_name(account):
    return account.upper() if len(account) <= 4 else account.replace("-", " ").title()


def short_date(d):
    return "%s %d" % (d.strftime("%b"), d.day) if d else ""


def long_date(d):
    return "%s %d, %d" % (d.strftime("%B"), d.day, d.year)


def days_between(earlier, ref):
    return max(0, (ref - earlier.date()).days) if earlier else None


def _useful_component(value):
    return value if value and not value.lower().startswith(("other", "n/a", "none")) else ""


def area(c):
    """Where the issue is: 'Reporting · Exports' (Component / Sub-Component)
    or 'Acme Server 8.4p2' (Product Version, for self-hosted products)."""
    parts = [p for p in (_useful_component(c.get("component")), _useful_component(c.get("subcomponent"))) if p]
    if len(parts) == 2 and parts[1].lower().startswith(parts[0].lower()):
        parts = parts[1:]
    if c.get("version"):
        parts.append(c["version"])
    return " \u00b7 ".join(parts)


def is_self_hosted(cases):
    names = [n.lower() for n in SELF_HOSTED_PRODUCTS]
    return any(n in (c.get("product") or "").lower() or n in (c.get("version") or "").lower()
               for c in cases for n in names)


def engineering(c):
    """'Engineering · ENG-12345' when the case has a Jira ticket (Has Jira? or a bug number), else None."""
    if not (c.get("has_jira") or is_ticket(c.get("bug"))):
        return None
    ticket = (jira_of(c) or ("",))[0]
    return "Engineering" + (" \u00b7 " + ticket if ticket else "")


def _same_person(a, b):
    letters = lambda x: re.sub(r"[^a-z]", "", (x or "").lower())
    return bool(letters(a)) and letters(a) == letters(b)


def customer_replied(c):
    """When the customer contact updated the case after the owner last did, else None."""
    na, a = c.get("non_assignee_updated_dt"), c.get("assignee_updated_dt")
    if na and (not a or na > a) and _same_person(c.get("non_assignee_by"), c.get("contact")):
        return na
    return None


def case_flags(c, as_of):
    """Short internal flags for Terminal, the app and the prep list. Never on the page."""
    flags = []
    esc = (c.get("escalation") or "").lower()
    if c.get("escalated") or (esc and esc not in ("normal", "none")):
        flags.append("escalated" + (" (%s)" % c["escalation"] if esc and esc not in ("normal", "none") else ""))
    if c.get("needs_attention") or (c.get("action_status") or "").lower() == "needs attention":
        flags.append("needs attention")
    replied = customer_replied(c)
    if replied:
        flags.append("customer replied %s, no owner update since" % short_date(replied.date()))
    f = c.get("follow_up_dt")
    if f:
        if f.date() < as_of:
            flags.append("follow-up overdue (%s)" % short_date(f.date()))
        elif (f.date() - as_of).days <= FOLLOW_UP_SOON_DAYS:
            flags.append("follow-up due %s" % short_date(f.date()))
    sla = c.get("sla_due_dt")
    if sla and (sla.date() - as_of).days <= 1:
        flags.append("SLA due %s" % short_date(sla.date()))
    if (c.get("reassignments") or 0) >= REASSIGN_WARN:
        flags.append("reassigned %d times" % c["reassignments"])
    return flags


# What each recommended column adds, for the "add these to your view" tip.
RECOMMENDED = [
    ("Highest Priority", "priority", "priority tags"),
    ("State", "state", "whose move it is"),
    ("Opened", "opened", "case age"),
    ("Updated", "updated", "which cases need an update"),
    ("Assigned to", "owner", "owner names"),
    ("Customer Contact", "contact", "contact names"),
    ("Comments and Work notes", "journal", "latest comments"),
    ("Request type", "type", "page sections"),
    ("Component", "component", "product areas"),
    ("Sub-Component", "subcomponent", "product areas"),     # hosted / cloud products
    ("Product Version", "version", "product versions"),      # self-hosted products
    ("Has Jira?", "has_jira", "the Engineering tag"),
    ("Assignee Updated", "assignee_updated", "owner idle time"),
    ("Non-Assignee Updated", "non_assignee_updated", "customer-replied alerts"),
    ("Non-Assignee Updated By", "non_assignee_by", "customer-replied alerts"),
    ("Follow up", "follow_up", "follow-up reminders"),
    ("Action status", "action_status", "needs-attention flags"),
]


def missing_columns(mapping, cases=()):
    """[(column, what it adds)] for recommended columns this export doesn't have.
    Product Version is suggested for self-hosted products (see SELF_HOSTED_PRODUCTS),
    Sub-Component for the rest."""
    self_hosted = is_self_hosted(cases)
    skip = {"subcomponent"} if self_hosted else ({"version"} if SELF_HOSTED_PRODUCTS else set())
    labels = {"type": TYPE_COLUMN_LABEL, "version": "Product Version"}
    whys = {"version": VERSION_LABEL}
    return [(labels.get(field, col), whys.get(field, why)) for col, field, why in RECOMMENDED
            if field not in mapping and field not in skip]


def topic_callouts(cases):
    labels = [_useful_component(c.get("subcomponent")) or _useful_component(c.get("component")) for c in cases]
    if PAGE_SHOW_AREA and sum(1 for x in labels if x) >= 0.5 * len(cases):   # real product areas beat keywords
        counts = {}
        for label in labels:
            if label:
                counts[label] = counts.get(label, 0) + 1
        hits = [(n, t) for t, n in counts.items() if n >= 3 and n >= 0.3 * len(cases)]
        return sorted(hits, reverse=True)[:2]
    hits = []
    for topic, stems in TOPICS.items():
        n = 0
        for c in cases:
            words = re.findall(r"[a-z0-9]+", c["description"].lower())
            if any(w.startswith(s) for w in words for s in stems):
                n += 1
        if n >= 3 and n >= 0.3 * len(cases):
            hits.append((n, topic))
    return sorted(hits, reverse=True)[:2]


# --- the interactive pass ---------------------------------------------------
USE_COLOR = sys.stdout.isatty() and not os.environ.get("NO_COLOR")


def _c(code, text):
    return "\033[%sm%s\033[0m" % (code, text) if USE_COLOR else text


def _prompt(code, text):
    """Colored prompt for input(). The \\001/\\002 markers tell readline the color codes take no
    room; without them it miscounts the width and Ctrl+U / Ctrl+K leave text behind on long lines."""
    if USE_COLOR and HAVE_READLINE:
        return "\001\033[%sm\002%s\001\033[0m\002" % (code, text)
    return _c(code, text)


def ask(prompt):
    """input() that keeps a multi-line paste together. Lines already waiting right after Enter
    came from the same paste, so they join this answer instead of spilling into the next case."""
    parts = [input(prompt)]
    if os.name == "posix" and sys.stdin.isatty():
        while select.select([sys.stdin], [], [], 0.05)[0]:
            parts.append(input())
    lines = [p.strip() for p in parts if p.strip()]
    if len(lines) > 1:
        print(_c("2", "   (pasted %d lines, kept as one update)" % len(lines)))
    return " ".join(lines)


def count(n, word):
    return "%d %s%s" % (n, word, "" if n == 1 else "s")


# What Enter does on a case when you type nothing:
#   "latest" - copy the newest ServiceNow comment OR work note (whichever is newer)
#   "public" - copy only the newest customer-visible comment
#   "off"    - keep your last update as it is
AUTOFILL = "public"
AUTOFILL_CHARS = 280          # copied text is trimmed to about this length for the page
COPY_TAGS = {"customer": "[comment]", "internal": "[internal note]"}
GREETING = re.compile(r"^(hi|hello|hey|howdy|dear|good (morning|afternoon|evening))\b[^,.!:]{0,40}[,.!:]\s*",
                      re.I)


SIGN_OFF = re.compile(r"\b(thanks|thank you|many thanks|kind regards|best regards|warm regards|regards|"
                      r"best|cheers|sincerely)\s*[,!.]\s", re.I)


# ...and a bare closing like "Kind Regards Jon" at the very end.
SIGN_OFF_END = re.compile(r"\s(?i:kind regards|best regards|warm regards|regards|thanks|thank you|cheers|best)"
                          r"\s+[A-Z][\w.'-]*(\s+[A-Z][\w.'-]*){0,2}\s*$")


def is_copy(text):
    return text.startswith(tuple(COPY_TAGS.values()))


def page_text(text):
    """Note text as it appears on the page: copy tags removed."""
    for tag in COPY_TAGS.values():
        if text.startswith(tag):
            return text[len(tag):].strip()
    return text


def newest_entry(journal, kinds=("customer", "internal")):
    """Newest comment or work note written by a person (not System, not the intake form)."""
    for e in journal:
        if (e["kind"] in kinds and e["text"] and e["when"]
                and e["author"].lower() not in SYSTEM_AUTHORS
                and not e["text"].lower().startswith("are you a partner?")):
            return e
    return None


def condense(text, limit=AUTOFILL_CHARS):
    text = text.replace("[earlier email trimmed]", "").strip()
    text = GREETING.sub("", text, count=1)
    m = SIGN_OFF.search(text, 30) or SIGN_OFF_END.search(text, 30)   # drop "Thanks, Name | Title ..."
    if m:
        text = text[:m.start()].rstrip()
    if len(text) > limit:
        cut = text[:limit]
        stop = max(cut.rfind(". "), cut.rfind("? "), cut.rfind("! "))
        text = cut[:stop + 1] if stop > limit // 2 else cut.rsplit(" ", 1)[0] + " …"
    return text[:1].upper() + text[1:]


def autofill_line(case, notes):
    """The note line Enter would add, or None if your own update is already current."""
    if AUTOFILL == "off":
        return None
    kinds = ("customer",) if AUTOFILL == "public" else ("customer", "internal")
    e = newest_entry(parse_journal(case["journal"]), kinds)
    if not e:
        return None
    entries = notes.get(case["number"], [])
    if entries:
        when, _ = split_entry(entries[0])
        if when and when >= e["when"].date():
            return None
    line = "%d/%d - %s %s" % (e["when"].month, e["when"].day, COPY_TAGS[e["kind"]], condense(e["text"]))
    return None if line in entries else line


def needs_update(case, notes, prev_cases, state_was, as_of):
    entries = notes.get(case["number"])
    if not entries:
        return True
    when, text = split_entry(entries[0])
    if when and when >= as_of:          # already updated for this export
        return False
    if is_copy(text):                   # copied from ServiceNow: current unless there's something newer
        e = newest_entry(parse_journal(case["journal"]))
        return bool(e and when and e["when"].date() > when)
    if prev_cases is not None and case["number"] not in prev_cases:
        return True
    if case["number"] in state_was:
        return True
    return bool(when and case["updated_dt"] and case["updated_dt"].date() > when)


COMMENT_LINES = 4     # wrapped lines shown per comment before "..."


def latest_comments(journal):
    """Newest public comment and newest internal note written by a person, newest first."""
    picked = []
    for kind in ("customer", "internal"):
        e = next((e for e in journal if e["kind"] == kind and e["text"]
                  and e["author"].lower() not in SYSTEM_AUTHORS), None)
        if e:
            picked.append(e)
    return sorted(picked, key=lambda e: e["when"] or dt.datetime.min, reverse=True)


def _width():
    return max(60, min(110, shutil.get_terminal_size((100, 20)).columns)) - 13


def _wrapped(text, max_lines):
    lines = textwrap.wrap(text, _width()) or [""]
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1][:_width() - 2].rstrip() + " …"
    return lines


def print_comment(e):
    label = _c("33", "internal") if e["kind"] == "internal" else _c("36", "public  ")
    when = short_date(e["when"].date()) if e["when"] else "?"
    print("   %s %s" % (label, _c("2", "%s · %s" % (when, e["author"]))))
    for line in _wrapped(e["text"], COMMENT_LINES):
        print("            " + line)


def print_card(n, total, c, title, notes, prev_cases, state_was, as_of):
    tags = ""
    if prev_cases is not None and c["number"] not in prev_cases:
        tags = "  " + _c("1;32", "NEW")
    print(_c("2", "─" * 64))
    print(" %s  %s  %s  %s%s" % (_c("2", "%d/%d" % (n, total)), _c("1", c["number"]),
                                 c["priority"], _c("35", c["type"]), tags))
    print(" " + c["description"])

    state = c["state"]
    if c["number"] in state_was:
        state = "%s → %s" % (state_was[c["number"]] or "?", c["state"])
    if c["substate"]:
        state += " · " + c["substate"]
    if state:
        print("   %s %s" % (_c("2", "status  "), state))
    people = [p for p in (c["contact"] and "%s (%s)" % (c["contact"], title),
                          c["owner"] and "%s (%s)" % (c["owner"], VENDOR)) if p]
    if people:
        print("   %s %s" % (_c("2", "people  "), " · ".join(people)))
    if c["group"]:
        print("   %s %s" % (_c("2", "group   "), c["group"]))
    if c.get("project"):
        print("   %s %s" % (_c("2", "project "), c["project"]))
    where = " \u00b7 ".join(x for x in (area(c), c.get("tenant") and "tenant " + c["tenant"]) if x)
    if where:
        print("   %s %s" % (_c("2", "area    "), where))
    if c["opened_dt"]:
        age = days_between(c["opened_dt"], as_of)
        line = "%s · %d day%s" % (short_date(c["opened_dt"].date()), age, "" if age == 1 else "s")
        if c["updated_dt"]:
            line += " · last activity %s" % short_date(c["updated_dt"].date())
        print("   %s %s" % (_c("2", "opened  "), line))

    journal = parse_journal(c["journal"])
    jira = jira_of(c)
    if jira and jira[0]:
        print("   %s %s%s  %s" % (_c("2", "jira    "), jira[0], " · " + jira[1] if jira[1] else "",
                                _c("2", jira_url(jira[0]) or "")))
    elif jira:
        print("   %s %s" % (_c("2", "jira    "), "yes (ticket number not in the export)"))
    flags = case_flags(c, as_of)
    if flags:
        print("   %s %s" % (_c("2", "flags   "), _c("33", " \u00b7 ".join(flags))))
    for e in latest_comments(journal):
        print_comment(e)

    last = notes.get(c["number"], [None])[0]
    print("   %s %s" % (_c("1", "yours   "), last if last else _c("2", "(none yet)")))
    fill = autofill_line(c, notes)
    if fill:
        print("   %s %s" % (_c("2", "enter   "), _c("2", "copies " + fill)))
    return fill


def run_updates(account, title, ordered, notes, prev_cases, state_was, as_of):
    changed = [c for c in ordered if needs_update(c, notes, prev_cases, state_was, as_of)]
    print()
    if changed and len(changed) < len(ordered):
        answer = input("Update the %s with activity since your last update, or [a]ll %d? [Enter = %d] "
                       % (count(len(changed), "case"), len(ordered), len(changed))).strip().lower()
        todo = ordered if answer.startswith("a") else changed
    elif changed:
        todo = ordered
    else:
        answer = input("No activity since your last updates. Review all %d anyway? [y/N] "
                       % len(ordered)).strip().lower()
        todo = ordered if answer.startswith("y") else []
    if not todo:
        return

    if AUTOFILL == "off":
        print(_c("2", "\nType the update and press Enter.  Enter alone keeps the last one.  q = done."))
    else:
        print(_c("2", "\nType the update and press Enter.  Enter alone copies the newest ServiceNow entry"))
        print(_c("2", "(shown on the 'enter' line).  s = skip, keep as is.  q = done."))
    print(_c("2", "Comments and work notes only reach the page if you copy them.\n"))
    today = dt.date.today()
    stamp = "%d/%d" % (today.month, today.day)
    for n, c in enumerate(todo, 1):
        fill = print_card(n, len(todo), c, title, notes, prev_cases, state_was, as_of)
        try:
            text = ask(_prompt("1;36", "   update   › "))
        except (EOFError, KeyboardInterrupt):
            print("\nStopped. Updates so far are saved.")
            return
        if text.lower() == "q":
            return
        if text.lower() == "s":
            continue
        line = "%s - %s" % (stamp, text) if text else fill
        if line:
            notes.setdefault(c["number"], []).insert(0, line)
            save_notes(account, notes, title)


# --- page -------------------------------------------------------------------
# OneNote drops <style> blocks on paste, so styling is inline throughout.
F = "font-family:Calibri,'Segoe UI',Arial,sans-serif;"
INK, MUTED, FAINT, LINE, NAVY = "#101828", "#667085", "#98a2b3", "#d0d5dd", "#1f3864"
TH = F + ("font-size:10pt;font-weight:bold;color:#ffffff;background:%s;text-align:left;"
          "padding:6px 9px;border:1px solid %s;" % (NAVY, NAVY))
TD = F + "font-size:10.5pt;color:%s;padding:7px 9px;border:1px solid %s;vertical-align:top;" % (INK, LINE)
CUSTOMER_ACCENT, VENDOR_ACCENT, GOOD = "#b54708", "#175cd3", "#067647"
DASH = '<span style="color:%s;">&mdash;</span>' % FAINT


def pill(text, fg, bg):
    return ('<span style="%sfont-size:8.5pt;font-weight:bold;color:%s;background:%s;'
            'padding:1px 7px;border-radius:9px;white-space:nowrap;">%s</span>'
            % (F, fg, bg, escape(text)))


def tile(value, label, color):
    return ('<td style="%spadding:10px 18px;border:1px solid %s;background:#f9fafb;">'
            '<div style="%sfont-size:24pt;font-weight:bold;color:%s;line-height:1.1;">%d</div>'
            '<div style="%sfont-size:9.5pt;color:%s;white-space:nowrap;">%s</div></td>'
            % (F, LINE, F, color, value, F, MUTED, escape(label)))


WARN, CRITICAL, CALM = "#fab219", "#d03b3b", "#b9c0cb"   # status steps from the dataviz palette


def age_color(days):
    if days is None or days < AGE_WARN_DAYS:
        return None
    return CRITICAL if days >= AGE_ALERT_DAYS else WARN


def age_icon(days):
    """Colored triangle for 30+ / 90+ days; the day count beside it carries the meaning too."""
    color = age_color(days)
    return '<span style="color:%s;">&#9650;</span> ' % color if color else ""


def days_text(days):
    return "today" if days == 0 else "%d day%s" % (days, "" if days == 1 else "s")


def age_chart(cases, as_of):
    """One bar per case, oldest first. Built from table cells so it survives a OneNote paste."""
    aged = sorted(((days_between(c["opened_dt"], as_of), c) for c in cases if c["opened_dt"]),
                  key=lambda x: (-x[0], x[1]["number"]))
    if len(aged) < 2:
        return ""
    shown, rest = aged[:AGE_CHART_MAX], aged[AGE_CHART_MAX:]
    longest = max(shown[0][0], 1)
    days = sorted(d for d, _ in aged)
    mid = len(days) // 2
    median = days[mid] if len(days) % 2 else (days[mid - 1] + days[mid]) // 2
    over = sum(1 for d in days if d >= AGE_WARN_DAYS)

    h = ['<div style="%sfont-size:13pt;font-weight:bold;color:%s;margin:18px 0 2px;">How long cases have been open</div>'
         % (F, NAVY),
         '<div style="%sfont-size:9.5pt;color:%s;margin-bottom:6px;">Median %s &middot; %d of %d open %d+ days '
         '&nbsp;&nbsp;<span style="color:%s;">&#9650;</span> %d+ days &nbsp;'
         '<span style="color:%s;">&#9650;</span> %d+ days</div>'
         % (F, MUTED, days_text(median), over, len(aged), AGE_WARN_DAYS,
            WARN, AGE_WARN_DAYS, CRITICAL, AGE_ALERT_DAYS),
         '<table style="border-collapse:collapse;width:100%;max-width:900px;">']
    for d, c in shown:
        desc = c["description"] if len(c["description"]) <= 44 else c["description"][:43].rstrip() + "&hellip;"
        tip = escape("%s: %s, open %s (since %s)" % (c["number"], c["description"], days_text(d),
                                                     long_date(c["opened_dt"].date())))
        pct = max(1, round(82.0 * d / longest))      # leave room for the label at the bar tip
        h.append('<tr title="%s">'
                 '<td style="%sfont-size:9.5pt;color:%s;padding:3px 10px 3px 0;width:38%%;white-space:nowrap;">'
                 '<b>%s</b> <span style="color:%s;">%s</span></td>'
                 '<td style="padding:3px 0;border-left:1px solid %s;">'
                 '<table style="border-collapse:collapse;width:100%%;"><tr>'
                 '<td style="width:%d%%;height:12px;padding:0;background:%s;border-radius:0 4px 4px 0;'
                 'font-size:1pt;line-height:1px;">&nbsp;</td>'
                 '<td style="%sfont-size:9.5pt;color:%s;padding:0 0 0 6px;white-space:nowrap;'
                 'font-variant-numeric:tabular-nums;">%s%s</td>'
                 '</tr></table></td></tr>'
                 % (tip, F, INK, escape(c["number"]), MUTED, desc if "&hellip;" in desc else escape(desc),
                    LINE, pct, age_color(d) or CALM, F, INK, age_icon(d), days_text(d)))
    h.append("</table>")
    if rest:
        h.append('<div style="%sfont-size:9pt;color:%s;margin-top:4px;">+ %d more open %s or less</div>'
                 % (F, MUTED, len(rest), days_text(rest[0][0])))
    return "\n".join(h)


def latest_update_html(entries):
    if not entries:
        return DASH
    when, text = split_entry(entries[0])
    text = page_text(text)
    prefix = ('<span style="color:%s;font-size:9pt;">%s &middot; </span>' % (MUTED, short_date(when))
              if when else "")
    return prefix + escape(text)


def who_html(c, title, states_known):
    """Next step (when State is exported) and the people on the case."""
    lines = []
    s = side(c) if states_known else None
    if s == "customer":
        lines.append(pill("Action: %s" % title, CUSTOMER_ACCENT, "#fef0c7"))
    elif s == "vendor":
        lines.append(pill(VENDOR, VENDOR_ACCENT, "#e0eaff"))
    first, second = (c["owner"], c["contact"]) if s == "vendor" else (c["contact"], c["owner"])
    if first:
        lines.append(escape(first))
    if second:
        lines.append('<span style="color:%s;font-size:9pt;">%s</span>' % (MUTED, escape(second)))
    if SHOW_GROUP_ON_PAGE and c["group"]:
        lines.append('<span style="color:%s;font-size:8.5pt;">%s</span>' % (FAINT, escape(c["group"])))
    return "<br>".join(lines) or DASH


def action_divider(label, count, color, ncols):
    return ('<tr><td colspan="%d" style="%sfont-size:9pt;font-weight:bold;color:%s;background:#f2f4f7;'
            'padding:4px 9px;border:1px solid %s;">%s <span style="color:%s;font-weight:normal;">(%d)</span>'
            '</td></tr>' % (ncols, F, color, LINE, escape(label), MUTED, count))


def issue_html(c, show_tenant=False):
    """Short description, then the product area / tenant and an Engineering tag when known."""
    extra = []
    if PAGE_SHOW_AREA and area(c):
        extra.append(escape(area(c)))
    if show_tenant and c.get("tenant"):
        extra.append("Tenant " + escape(c["tenant"]))
    html = escape(c["description"])
    if extra:
        html += '<br><span style="color:%s;font-size:8.5pt;">%s</span>' % (MUTED, " &middot; ".join(extra))
    eng = engineering(c) if PAGE_SHOW_ENGINEERING else None
    if eng:
        url = jira_url((jira_of(c) or ("",))[0])
        tag = pill(eng, "#5925dc", "#f4f3ff")
        html += '<br>' + ('<a href="%s" style="text-decoration:none;">%s</a>' % (escape(url), tag) if url else tag)
    return html


def case_table(cases, notes, title, states_known, show_who, as_of, by_action=False, show_tenant=False):
    cols = [("Case", "12%"), ("Priority", "7%"), ("Issue", "31%"), ("Latest update", "34%")]
    if show_who:
        cols.append(("Next step" if states_known else "Contacts", "16%"))
    out = ['<table style="border-collapse:collapse;width:100%%;margin:4px 0 14px;"><tr>%s</tr>'
           % "".join('<th style="%swidth:%s;">%s</th>' % (TH, w, label) for label, w in cols)]
    labels = {"customer": ("Action: %s" % title, CUSTOMER_ACCENT), "vendor": (VENDOR, VENDOR_ACCENT),
              None: ("No state", MUTED)}
    current = "start"
    for i, c in enumerate(cases):
        if by_action and states_known and side(c) != current:
            current = side(c)
            label, color = labels[current]
            out.append(action_divider(label, sum(1 for x in cases if side(x) == current), color, len(cols)))
        bg = "background:#f9fafb;" if i % 2 else "background:#ffffff;"
        fg, pbg = PRIORITY_COLORS.get(c["priority"], (MUTED, "#eaecf0"))
        age = days_between(c["opened_dt"], as_of)
        opened = ('<br><span style="color:%s;font-size:8.5pt;white-space:nowrap;">Open %s</span>'
                  % (MUTED, days_text(age)) if age is not None else "")
        cells = ['<b>%s</b>%s' % (escape(c["number"]), opened),
                 pill(c["priority"] or "?", fg, pbg),
                 issue_html(c, show_tenant),
                 latest_update_html(notes.get(c["number"]))]
        if show_who:
            cells.append(who_html(c, title, states_known))
        out.append("<tr>%s</tr>" % "".join(
            '<td style="%s%s%s">%s</td>' % (TD, bg, "text-align:center;" if j == 1 else "", v)
            for j, v in enumerate(cells)))
    out.append("</table>")
    return "\n".join(out)


def section_heading(text, count, color):
    return ('<div style="%sfont-size:13pt;font-weight:bold;color:%s;margin:18px 0 2px;">%s '
            '<span style="color:%s;font-weight:normal;">(%d)</span></div>'
            % (F, color, escape(text), MUTED, count))


def render_html(title, subtitle, as_of, prev_date, groups, resolved, notes, by_action=False):
    open_cases = [c for _, cs in groups for c in cs]
    states_known = any(c["state"] for c in open_cases)
    show_who = states_known or any(c["contact"] or c["owner"] or (SHOW_GROUP_ON_PAGE and c["group"])
                                   for c in open_cases)
    h = ['<!doctype html><html><head><meta charset="utf-8">'
         '<title>%s Case Review</title></head>' % escape(title),
         '<body style="background:#ffffff;margin:28px;%s">' % F,
         '<div style="%smax-width:1100px;">' % F]

    h.append('<div style="%sfont-size:20pt;font-weight:bold;color:%s;">%s &middot; Support Case Review</div>'
             % (F, NAVY, escape(title)))
    sub = long_date(as_of) + (" &middot; " + escape(subtitle) if subtitle else "")
    h.append('<div style="%sfont-size:10.5pt;color:%s;margin-bottom:6px;">%s</div>' % (F, MUTED, sub))

    tiles = [tile(len(open_cases), "Open cases", NAVY)]
    if 1 < len(groups) <= TYPE_TILES_MAX:
        tiles += [tile(len(cs), name, type_color(name)) for name, cs in groups]
    if states_known:
        tiles.append(tile(sum(1 for c in open_cases if side(c) == "customer"),
                          "Action needed from %s" % title, CUSTOMER_ACCENT))
    aged = sum(1 for c in open_cases if (days_between(c["opened_dt"], as_of) or 0) >= AGE_WARN_DAYS)
    if any(c["opened_dt"] for c in open_cases):
        tiles.append(tile(aged, "Open %d+ days" % AGE_WARN_DAYS, "#b54708" if aged else MUTED))
    if prev_date:
        tiles.append(tile(len(resolved), "Resolved since %s" % short_date(prev_date), GOOD))
    h.append('<table style="border-collapse:collapse;margin:10px 0 8px;"><tr>%s</tr></table>' % "".join(tiles))
    if len(groups) > TYPE_TILES_MAX:   # too many types for boxes: one line instead
        h.append('<div style="%sfont-size:10pt;color:%s;margin:0 0 6px;">%s</div>' % (F, MUTED, " &middot; ".join(
            '<span style="color:%s;">&#9632;</span> %d %s' % (type_color(n), len(cs), escape(n)) for n, cs in groups)))

    for n, topic in topic_callouts(open_cases):
        h.append('<div style="%sfont-size:10.5pt;color:%s;background:#f2f4f7;border-left:4px solid %s;'
                 'padding:6px 10px;margin:6px 0;"><b>Focus area: %s</b> &mdash; %d of %d open cases.</div>'
                 % (F, INK, NAVY, escape(topic), n, len(open_cases)))

    h.append(age_chart(open_cases, as_of))

    show_tenant = PAGE_SHOW_TENANT and len({c.get("tenant") for c in open_cases if c.get("tenant")}) > 1
    for name, cases in groups:
        h.append(section_heading(name, len(cases), type_color(name)))
        h.append(case_table(cases, notes, title, states_known, show_who, as_of, by_action, show_tenant))

    if resolved:
        h.append(section_heading("Resolved since %s" % short_date(prev_date), len(resolved), GOOD))
        h.append('<ul style="%sfont-size:10.5pt;color:%s;margin-top:4px;list-style:none;padding-left:4px;">'
                 % (F, INK))
        for c in sorted(resolved, key=lambda c: (type_rank(c["type"]), c["number"])):
            kind = (' <span style="color:%s;font-size:9pt;">(%s)</span>' % (MUTED, escape(c["type"]))
                    if len({r["type"] for r in resolved} | {g for g, _ in groups}) > 1 else "")
            h.append('<li><span style="color:%s;">&#10003;</span> <b>%s</b> &mdash; %s%s</li>'
                     % (GOOD, escape(c["number"]), escape(c["description"]), kind))
        h.append("</ul>")

    h.append('<div style="%sfont-size:13pt;font-weight:bold;color:%s;margin:18px 0 4px;">'
             'Questions / concerns during the call</div>' % (F, NAVY))
    h.append('<ol style="%sfont-size:10.5pt;"><li>&nbsp;</li><li>&nbsp;</li></ol>' % F)
    h.append('<div style="%sfont-size:8.5pt;color:%s;margin-top:20px;">Source: ServiceNow, %s.</div>'
             % (F, FAINT, long_date(as_of)))
    h.append("</div></body></html>")
    return "\n".join(h)


def write_csv(path, title, groups, notes, as_of):
    """Same content as the page. Comments and work notes are deliberately left out."""
    def next_step(c):
        s = side(c)
        return "Action: %s" % title if s == "customer" else VENDOR if s == "vendor" else ""

    every = [c for _, cs in groups for c in cs]
    has_area = any(area(c) for c in every)
    has_eng = any(engineering(c) for c in every)
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Type", "Case", "Priority", "Issue", "Latest update", "Update date",
                    "Next step", "%s contact" % title, "%s owner" % VENDOR, "Assignment group", "Opened", "Days open"]
                   + (["Area"] if PAGE_SHOW_AREA and has_area else [])
                   + (["Engineering"] if PAGE_SHOW_ENGINEERING and has_eng else []))
        for _, cases in groups:
            for c in cases:
                when, text = split_entry(notes[c["number"]][0]) if notes.get(c["number"]) else (None, "")
                text = page_text(text)
                w.writerow([c["type"], c["number"], c["priority"], c["description"], text,
                            when.isoformat() if when else "", next_step(c), c["contact"], c["owner"], c["group"],
                            c["opened_dt"].date().isoformat() if c["opened_dt"] else "",
                            days_between(c["opened_dt"], as_of) if c["opened_dt"] else ""]
                           + ([area(c)] if PAGE_SHOW_AREA and has_area else [])
                           + ([engineering(c) or ""] if PAGE_SHOW_ENGINEERING and has_eng else []))


def render_text(title, as_of, prev_date, groups, resolved, notes, new_ids, by_action):
    """Plain-text version for tools that drop tables (Gainsight notes, email, Slack)."""
    open_cases = [c for _, cs in groups for c in cs]
    states_known = any(c["state"] for c in open_cases)
    out = ["%s · Support Case Review · %s" % (title, long_date(as_of)), ""]

    summary = ["%d open case%s" % (len(open_cases), "" if len(open_cases) == 1 else "s")]
    if states_known:
        summary.append("%d action needed from %s" % (sum(1 for c in open_cases if side(c) == "customer"), title))
    aged = sorted(((days_between(c["opened_dt"], as_of), c) for c in open_cases if c["opened_dt"]),
                  key=lambda x: -x[0])
    old = [(d, c) for d, c in aged if d >= AGE_WARN_DAYS]
    if aged:
        summary.append("%d open %d+ days" % (len(old), AGE_WARN_DAYS))
    if prev_date:
        summary.append("%d resolved since %s" % (len(resolved), short_date(prev_date)))
    out += [" · ".join(summary), ""]
    if old:
        out.append("Open %d+ days: %s" % (AGE_WARN_DAYS, ", ".join(
            "%s (%s)" % (c["number"], days_text(d)) for d, c in old)))
        out.append("")

    labels = {"customer": "Action: %s" % title, "vendor": VENDOR, None: "No state"}
    for name, cases in groups:
        out.append("%s (%d)" % (name.upper(), len(cases)))
        current = "start"
        for c in cases:
            if by_action and states_known and side(c) != current:
                current = side(c)
                out += ["", "%s (%d)" % (labels[current], sum(1 for x in cases if side(x) == current))]
            age = days_between(c["opened_dt"], as_of)
            new = " [NEW]" if c["number"] in new_ids else ""
            out.append("• %s · %s · %s%s" % (c["number"], c["priority"] or "?", c["description"], new))
            detail = []
            if age is not None:
                detail.append("Open %s" % days_text(age))
            if PAGE_SHOW_AREA and area(c):
                detail.append(area(c))
            if PAGE_SHOW_ENGINEERING and engineering(c):
                detail.append(engineering(c))
            s = side(c) if states_known else None
            who = (c["owner"] or c["contact"]) if s == "vendor" else (c["contact"] or c["owner"])
            if s:
                detail.append("Next: %s%s" % (title if s == "customer" else VENDOR, " (%s)" % who if who else ""))
            elif who:
                detail.append(who)
            if detail:
                out.append("   " + " · ".join(detail))
            entries = notes.get(c["number"])
            if entries:
                when, text = split_entry(entries[0])
                out.append("   Update%s: %s" % (" " + short_date(when) if when else "", page_text(text)))
        out.append("")

    if resolved:
        out.append("RESOLVED SINCE %s (%d)" % (short_date(prev_date).upper(), len(resolved)))
        for c in sorted(resolved, key=lambda c: c["number"]):
            out.append("✓ %s · %s" % (c["number"], c["description"]))
        out.append("")
    out += ["QUESTIONS / CONCERNS DURING THE CALL", "1.", "2.", ""]
    return "\n".join(out)


def copy_to_clipboard(text):
    """True if the text is now on the clipboard (macOS pbcopy, Windows clip)."""
    try:
        if sys.platform == "darwin":
            env = dict(os.environ, LANG="en_US.UTF-8")
            subprocess.run(["pbcopy"], input=text.encode("utf-8"), check=True, env=env)
            return True
        if os.name == "nt":
            subprocess.run(["clip"], input=text.encode("utf-16-le"), check=True,
                           creationflags=0x08000000)   # CREATE_NO_WINDOW: no console flash from the app
            return True
    except (OSError, subprocess.CalledProcessError):
        pass
    return False


def open_file(path):
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        elif os.name == "nt":
            os.startfile(path)  # noqa
        else:
            subprocess.run(["xdg-open", path], check=False)
    except OSError:
        pass


# --- commands ---------------------------------------------------------------
def load_review(account, export=None, by_action=True):
    """Everything a review needs, as a dict. Shared by the command line and the desktop app."""
    files = exports_for(account)
    if export:
        export = os.path.abspath(export)
        if export not in files:
            files = sorted(files + [export], key=lambda p: (export_date(p), os.path.getmtime(p)))
    if not files:
        raise SystemExit("No exports for '%s'. Put them in exports/%s/" % (account, account))
    current = export or files[-1]
    earlier = [p for p in files if export_date(p) < export_date(current)]
    previous = earlier[-1] if earlier else None

    cases, mapping = load_export(current)
    open_cases = [c for c in cases if not is_closed(c)]
    open_ids = {c["number"] for c in open_cases}
    resolved = [c for c in cases if is_closed(c)]

    prev_cases, prev_date = None, None
    if previous:
        prev_date = export_date(previous)
        prev_cases = {c["number"]: c for c in load_export(previous)[0] if not is_closed(c)}
        seen = {c["number"] for c in resolved}
        resolved += [c for n, c in prev_cases.items() if n not in open_ids and n not in seen]
    new_ids = {n for n in open_ids if prev_cases is not None and n not in prev_cases}
    state_was = {c["number"]: prev_cases[c["number"]]["state"] for c in open_cases
                 if prev_cases and c["number"] in prev_cases and c["state"]
                 and prev_cases[c["number"]]["state"]
                 and prev_cases[c["number"]]["state"] != c["state"]}

    as_of = min(export_date(current), dt.date.today())
    title = display_name(account)
    accounts = [c["account"] for c in cases if c["account"]]
    subtitle = max(set(accounts), key=accounts.count) if accounts else ""
    groups = grouped(open_cases, by_action)
    ordered = [c for _, cs in groups for c in cs]
    return {"account": account, "current": current, "previous": previous, "mapping": mapping,
            "open_cases": open_cases, "resolved": resolved, "prev_cases": prev_cases,
            "prev_date": prev_date, "new_ids": new_ids, "state_was": state_was, "as_of": as_of,
            "title": title, "subtitle": subtitle, "groups": groups, "ordered": ordered,
            "by_action": by_action}


def write_outputs(r, notes):
    """Write the page, CSV and text for a loaded review. Returns (base path, text)."""
    account, as_of = r["account"], r["as_of"]
    folder = os.path.join(OUT_DIR, account)
    os.makedirs(folder, exist_ok=True)
    base = os.path.join(folder, "%s-%s%s" % (account, as_of.isoformat(),
                                             "" if r["by_action"] else "-by-priority"))
    with open(base + ".html", "w", encoding="utf-8") as fh:
        fh.write(render_html(r["title"], r["subtitle"], as_of, r["prev_date"], r["groups"],
                             r["resolved"], notes, r["by_action"]))
    write_csv(base + ".csv", r["title"], r["groups"], notes, as_of)
    text = render_text(r["title"], as_of, r["prev_date"], r["groups"], r["resolved"], notes,
                       r["new_ids"], r["by_action"])
    with open(base + ".txt", "w", encoding="utf-8") as fh:
        fh.write(text)
    return base, text


def prep_list(r, notes):
    """'Before the call' reminders: internal only, never on the page."""
    ordered, as_of = r["ordered"], r["as_of"]
    prep = []
    for c in ordered:
        age = days_between(c["opened_dt"], as_of)
        idle = days_between(c["updated_dt"], as_of)
        if c["priority"] == "P1":
            prep.append("%s is P1" % c["number"])
        if age is not None and age >= AGING_DAYS:
            prep.append("%s open %d days (%s)" % (c["number"], age, c["state"] or "no state"))
        elif c.get("assignee_updated_dt") and side(c) != "customer":
            owner_idle = days_between(c["assignee_updated_dt"], as_of)
            if owner_idle >= IDLE_DAYS:
                prep.append("%s no update from the owner in %d days" % (c["number"], owner_idle))
        elif idle is not None and idle >= IDLE_DAYS and side(c) != "customer":
            prep.append("%s no ServiceNow activity in %d days" % (c["number"], idle))
        flags = case_flags(c, as_of)
        if flags:
            prep.append("%s %s" % (c["number"], " \u00b7 ".join(flags)))
    internal = [c["number"] for c in ordered if notes.get(c["number"])
                and split_entry(notes[c["number"]][0])[1].startswith(COPY_TAGS["internal"])]
    if internal:
        prep.append("page shows a copied INTERNAL note, check before presenting: %s" % ", ".join(internal))
    missing = [c["number"] for c in ordered if not notes.get(c["number"])]
    if missing:
        prep.append("no status update yet: %s" % ", ".join(missing))
    return prep


def review(account, export=None, update=True, open_page=True, by_action=True):
    r = load_review(account, export, by_action)
    title, open_cases, previous = r["title"], r["open_cases"], r["previous"]
    notes = load_notes(account)

    print("%s — %s (%s%s)"
          % (title, count(len(open_cases), "open case"), os.path.basename(r["current"]),
             ", compared with %s" % os.path.basename(previous) if previous else ""))
    if r["prev_cases"] is not None:
        print("  %d new · %d resolved · %d changed status"
              % (len(r["new_ids"]), len(r["resolved"]), len(r["state_was"])))
    print("  " + " · ".join("%d %s" % (len(cs), name) for name, cs in r["groups"]))
    if "state" not in r["mapping"]:
        print("  No State column in this export, so the page can't show whose move it is.")
    missing = missing_columns(r["mapping"], r["open_cases"])
    if missing:
        print(_c("2", "  Tip: add %s to your ServiceNow view for %s. See docs/servicenow-view.md."
                 % (", ".join(col for col, _ in missing[:4]) + (" and %d more" % (len(missing) - 4) if len(missing) > 4 else ""),
                    ", ".join(dict.fromkeys(why for _, why in missing[:4])))))

    if update:
        run_updates(account, title, r["ordered"], notes, r["prev_cases"], r["state_was"], r["as_of"])
        notes = load_notes(account)

    base, text = write_outputs(r, notes)
    print("\nPage  → %s" % os.path.relpath(base + ".html", ROOT))
    print("Sheet → %s" % os.path.relpath(base + ".csv", ROOT))
    print("Text  → %s" % os.path.relpath(base + ".txt", ROOT))
    if copy_to_clipboard(text):
        print("        (copied to the clipboard - paste it into Gainsight)")

    prep = prep_list(r, notes)
    if prep:
        print("\nBefore the call (not on the page):")
        for line in prep:
            print("  • " + line)

    if open_page and sys.stdout.isatty():
        open_file(base + ".html")


def pick_account():
    accounts = list_accounts()
    if not accounts:
        raise SystemExit("No exports yet. Put one in exports/<account>/ and run again.")
    for i, a in enumerate(accounts, 1):
        print("  %2d  %s" % (i, display_name(a)))
    choice = input("Account? ").strip().lower()
    if choice.isdigit() and 1 <= int(choice) <= len(accounts):
        return accounts[int(choice) - 1]
    if choice in accounts:
        return choice
    raise SystemExit("Unknown account: %s" % choice)


def across(open_page=True):
    """Internal view: everything waiting on us, every account, oldest first."""
    rows = []
    for account in list_accounts():
        files = exports_for(account)
        if not files:
            continue
        cases, _ = load_export(files[-1])
        as_of = min(export_date(files[-1]), dt.date.today())
        for c in cases:
            if not is_closed(c) and side(c) != "customer":
                rows.append((account, c, as_of))
    if not rows:
        raise SystemExit("No exports found in exports/")
    rows.sort(key=lambda r: -(days_between(r[1]["opened_dt"], r[2]) or 0))
    today = dt.date.today()
    h = ['<!doctype html><html><head><meta charset="utf-8"><title>Waiting on us</title></head>',
         '<body style="background:#ffffff;margin:28px;%s">' % F,
         '<div style="%sfont-size:20pt;font-weight:bold;color:%s;">Waiting on %s &mdash; all accounts</div>'
         % (F, NAVY, VENDOR),
         '<div style="%sfont-size:10.5pt;color:%s;margin-bottom:10px;">%s &middot; %s across %s '
         '&middot; oldest first</div>' % (F, MUTED, long_date(today), count(len(rows), "case"),
                                           count(len({r[0] for r in rows}), "account")),
         '<table style="border-collapse:collapse;width:100%;"><tr>'
         + "".join('<th style="%s">%s</th>' % (TH, x)
                   for x in ["Account", "Case", "Type", "Priority", "Age", "Idle", "Issue", "Owner"]) + "</tr>"]
    for i, (account, c, as_of) in enumerate(rows):
        age, idle = days_between(c["opened_dt"], as_of), days_between(c["updated_dt"], as_of)
        hot = (age or 0) >= AGING_DAYS or (idle or 0) >= IDLE_DAYS
        bg = "background:#fef3f2;" if hot else ("background:#f9fafb;" if i % 2 else "")
        fg, pbg = PRIORITY_COLORS.get(c["priority"], (MUTED, "#eaecf0"))
        h.append("<tr>" + "".join('<td style="%s%s">%s</td>' % (TD, bg, v) for v in [
            "<b>%s</b>" % escape(display_name(account)), escape(c["number"]),
            '<span style="color:%s;">%s</span>' % (type_color(c["type"]), escape(c["type"])),
            pill(c["priority"] or "?", fg, pbg),
            "%sd" % age if age is not None else "?", "%sd" % idle if idle is not None else "?",
            escape(c["description"]),
            (escape(c["owner"]) or "&mdash;") + ('<br><span style="color:%s;font-size:9pt;">%s</span>'
                                                 % (MUTED, escape(c["group"])) if c["group"] else "")])
            + "</tr>")
    h.append("</table></body></html>")
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, "waiting-on-us-%s.html" % today.isoformat())
    open(out, "w", encoding="utf-8").write("\n".join(h))
    print("%s waiting on %s across %s → %s"
          % (count(len(rows), "case"), VENDOR, count(len({r[0] for r in rows}), "account"),
             os.path.relpath(out, ROOT)))
    if open_page and sys.stdout.isatty():
        open_file(out)


load_profile()


def main(argv):
    flags = {a for a in argv if a.startswith("--")}
    args = [a for a in argv if not a.startswith("--")]
    if flags & {"--help", "-h"} or (args and args[0] in ("help", "-h")):
        print(__doc__)
        return
    open_page = "--no-open" not in flags
    if args and args[0] == "across":
        across(open_page)
        return
    if args and args[0] == "build":  # older command name
        args = args[1:]
    export = None
    if args and _is_export(args[0]):
        export = args[0]
        parent = os.path.basename(os.path.dirname(os.path.abspath(export)))
        account = parent.lower() if parent.lower() != "exports" else account_from(export)
    elif args:
        account = args[0].lower()
    else:
        account = pick_account()
    review(account, export, update="--no-update" not in flags, open_page=open_page,
           by_action="--by-priority" not in flags)


if __name__ == "__main__":
    try:
        main(sys.argv[1:])
    except KeyboardInterrupt:
        print()
