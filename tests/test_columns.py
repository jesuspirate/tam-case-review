"""
Any mix of export columns must still give a sensible review.

Builds a fake ServiceNow export (made-up cases, no customer data) and runs the whole
review - page, CSV, text, prep list, app data - for many column combinations:
the bare minimum, each column left out in turn, random subsets, shuffled order,
odd capitalisation, unknown extra columns, and the traps (a "Case" or long
"Description" column, the same comments in several columns).

    python3 tests/test_columns.py              (stdlib only)
    python3 tests/test_columns.py some.xlsx    also run the combinations on a real export
"""

import csv
import datetime as dt
import os
import random
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "app"))
import callprep as cp   # noqa: E402

cp.copy_to_clipboard = lambda text: False     # tests never touch the clipboard
cp.open_file = lambda path: None

# Tests run with the neutral example profile, whatever profile.json you have locally.
import json   # noqa: E402
EXAMPLE_PROFILE = json.load(open(os.path.join(os.path.dirname(HERE), "profiles", "example.json")))

TODAY = dt.date.today()


def stamp(days_ago, hour=10):
    d = dt.datetime.combine(TODAY - dt.timedelta(days=days_ago), dt.time(hour, 0, 0))
    return d.strftime("%Y-%m-%d %H:%M:%S")


def journal(entries):
    """ServiceNow journal text: newest first, 'MM-DD-YYYY HH:MM:SS - Author (Kind)' headers."""
    out = []
    for days_ago, author, kind, text in entries:
        d = dt.datetime.combine(TODAY - dt.timedelta(days=days_ago), dt.time(9, 30))
        out.append("%s - %s (%s)\n%s\n" % (d.strftime("%m-%d-%Y %H:%M:%S"), author,
                                             "Work notes" if kind == "internal" else "Additional comments", text))
    return "\n".join(out)


CASES = [
    # number, prio, desc, state, substate, opened, updated, contact, owner, type, ps type, component, sub, jira, action,
    # follow-up, assignee upd, non-assignee upd, non-assignee by, version, tenant, comments
    ("CS0000101", "1 - Critical", "Certification campaign stuck in staging", "Pending Acme", "Pending Internal",
     40, 1, "pat customer", "Sam Owner", "", "", "Identity Governance", "Certifications", "1", "Needs attention",
     1, 5, 1, "Pat Customer", "", "https://acme.cloud.example.com/ui", [
         (1, "Pat Customer", "customer", "Any update? We are blocked on the certification."),
         (5, "Sam Owner", "internal", "Escalated to engineering in ENG-12345."),
         (9, "Sam Owner", "customer", "Hi Pat, we are looking into it. Thanks, Sam")]),
    ("CS0000102", "P2", "Certification reminders not sent", "Pending Customer", "Awaiting Info",
     95, 3, "Lee Customer", "Ana Owner", "", "", "Identity Governance", "Certifications", "0", "Blocked by customer",
     12, 3, 6, "Lee Customer", "", "https://acme-sb.cloud.example.com", [
         (3, "Ana Owner", "customer", "Please send the HAR file."),
         (6, "Lee Customer", "customer", "Here is what we see.")]),
    ("CS0000103", "3", "Certification report export missing columns", "Pending Acme", "In Progress",
     10, 2, "Pat Customer", "Sam Owner", "Ask an Expert", "", "Identity Governance", "Certifications", "0", "",
     "", 2, 4, "Pat Customer", "", "https://acme.cloud.example.com", []),
    ("CS0000104", "P4", "Rule review for connector rules", "New", "",
     2, 0, "", "", "Health Check", "", "Connectors & Integrations", "Rules", "", "",
     "", "", "", "", "", "", [(0, "System", "internal", "Case created")]),
    ("CS0000105", "P3", "Upgrade planning 8.3 to 8.4", "Pending Acme", "In Progress",
     200, 20, "Kim Customer", "Raj Owner", "", "Professional Services", "Other Server", "", "1", "Needs attention",
     "", 20, 15, "Kim Customer", "Acme Server 8.3p3", "", [(15, "Kim Customer", "customer", "Following up on the plan.")]),
    ("CS0000106", "P2", "Old resolved case", "Resolved", "",
     60, 30, "Kim Customer", "Raj Owner", "", "Platinum Support", "", "", "0", "",
     "", 30, 31, "Kim Customer", "Acme Server 8.4p2", "", []),
]

# header -> value builder; every column name is one the script should understand (or ignore safely)
COLUMNS = {
    "Number": lambda c: c[0],
    "Highest Priority": lambda c: c[1],
    "Short description": lambda c: c[2],
    "State": lambda c: c[3],
    "Substate": lambda c: c[4],
    "Opened": lambda c: stamp(c[5]),
    "Updated": lambda c: stamp(c[6]),
    "Customer Contact": lambda c: c[7],
    "Assigned to": lambda c: c[8],
    "Assignment group": lambda c: "Support-Test",
    "Customer Account": lambda c: "Acme Corp",
    "Success Package Request Type": lambda c: c[9],
    "Professional Services Type": lambda c: c[10],
    "Component": lambda c: c[11],
    "Sub-Component": lambda c: c[12],
    "Has Jira?": lambda c: c[13],
    "Action status": lambda c: c[14],
    "Follow up": lambda c: stamp(-c[15]) if c[15] != "" else "",
    "Assignee Updated": lambda c: stamp(c[16], 8) if c[16] != "" else "",
    "Non-Assignee Updated": lambda c: stamp(c[17], 12) if c[17] != "" else "",
    "Non-Assignee Updated By": lambda c: c[18],
    "Product Version": lambda c: c[19],
    "Tenant URL": lambda c: c[20],
    "Comments and Work notes": lambda c: journal(c[21]),
    "Description": lambda c: "Long description. " * 40 + c[2],
    "Escalation": lambda c: "Normal",
    "Reassignment count": lambda c: "3.0" if c[0] == "CS0000101" else "0.0",
    "Impact": lambda c: "Widespread",
    "Urgency": lambda c: "High",
    "Unknown Custom Field": lambda c: "whatever",
}
ALL = list(COLUMNS)


def write_export(path, headers, transform=None, extra=None, first=None):
    """CSV export with the given columns. transform renames headers; extra / first add
    raw (header, value) columns after / before them."""
    extra, first = extra or [], first or []
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow([h for h, _ in first] + [transform(h) if transform else h for h in headers] + [h for h, _ in extra])
        for c in CASES:
            w.writerow([f(c) for _, f in first] + [COLUMNS[h](c) for h in headers] + [f(c) for _, f in extra])


class ColumnCombinations(unittest.TestCase):
    def setUp(self):
        cp.apply_profile(EXAMPLE_PROFILE)
        self.root = tempfile.mkdtemp(prefix="callprep-test-")
        cp.set_root(self.root)
        os.makedirs(os.path.join(self.root, "exports", "acme"))

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def run_review(self, headers, transform=None, extra=None, previous=None, first=None):
        folder = os.path.join(self.root, "exports", "acme")
        for name in os.listdir(folder):
            os.remove(os.path.join(folder, name))
        if previous:
            write_export(os.path.join(folder, "acme_tickets-%s.csv" % (TODAY - dt.timedelta(days=7))), previous)
        write_export(os.path.join(folder, "acme_tickets-%s.csv" % TODAY), headers, transform, extra, first)
        r = cp.load_review("acme")
        notes = {"CS0000101": ["%d/%d - We are on it." % (TODAY.month, TODAY.day)]}
        base, text = cp.write_outputs(r, notes)
        html = open(base + ".html", encoding="utf-8").read()
        prep = cp.prep_list(r, notes)
        cp.missing_columns(r["mapping"], r["open_cases"])
        for c in r["ordered"]:        # what Terminal and the app show per case
            cp.case_flags(c, r["as_of"])
            cp.area(c)
            cp.engineering(c)
        self.check_app(headers)
        return r, html, text, prep

    def check_app(self, headers):
        try:
            import tam_app
        except ImportError:
            return
        data = tam_app.review_json("acme")
        self.assertTrue(all("flags" in c for c in data["cases"]))

    def assert_sane(self, r, html, text):
        numbers = [c["number"] for c in r["ordered"]]
        self.assertTrue(numbers, "no cases read")
        for n in numbers:
            self.assertIn(n, html)
            self.assertIn(n, text)
        for c in r["ordered"]:
            self.assertLessEqual(len(c["description"]), 145, "a long Description leaked into the issue column")
        self.assertFalse("Long description. " * 12 in html, "the full long Description reached the page")

    def test_everything(self):
        r, html, text, prep = self.run_review(ALL)
        self.assert_sane(r, html, text)
        self.assertIn("Identity Governance \u00b7 Certifications", html)
        self.assertIn("Engineering \u00b7 ENG-12345", html)          # Has Jira? + id from the comments
        self.assertIn("Acme Server 8.3p3", html)                             # self-hosted version as the area
        self.assertNotIn("Other Server", html)                                  # vague components hidden
        self.assertIn("Focus area: Certifications", html)
        self.assertNotIn("acme-sb", html)                                    # tenant stays internal by default
        self.assertNotIn("customer replied", html)                           # flags stay internal
        joined = "\n".join(prep)
        self.assertIn("CS0000101 needs attention", joined)
        self.assertIn("customer replied", joined)
        self.assertIn("reassigned 3 times", joined)

    def test_bare_minimum(self):
        for headers in (["Number"], ["Number", "Short description"], ["Short description", "Number"]):
            r, html, text, _ = self.run_review(headers)
            self.assertEqual(len(r["ordered"]), len(CASES))
            self.assertIn("CS0000101", html)

    def test_each_column_left_out(self):
        for skip in ALL:
            if skip == "Number":
                continue
            with self.subTest(left_out=skip):
                r, html, text, _ = self.run_review([h for h in ALL if h != skip])
                self.assert_sane(r, html, text)

    def test_random_subsets_and_order(self):
        rnd = random.Random(42)
        others = [h for h in ALL if h != "Number"]
        for i in range(60):
            headers = ["Number"] + rnd.sample(others, rnd.randint(0, len(others)))
            rnd.shuffle(headers)
            with self.subTest(i=i, headers=headers):
                r, html, text, _ = self.run_review(headers)
                self.assert_sane(r, html, text)

    def test_header_case_and_spacing(self):
        r, html, text, _ = self.run_review(ALL, transform=lambda h: "  " + h.upper().replace(" ", "  ") + " ")
        self.assert_sane(r, html, text)
        self.assertIn("Identity Governance \u00b7 Certifications", html)

    def test_only_long_description(self):
        headers = [h for h in ALL if h != "Short description"]
        r, html, text, _ = self.run_review(headers)
        self.assert_sane(r, html, text)

    def test_case_column_before_number(self):
        # "Case" ("CS0000101 - text") is the first column and "Number" comes later: Number must win
        first = [("Case", lambda c: "%s - %s" % (c[0], c[2]))]
        r, html, text, _ = self.run_review(["Short description", "State", "Number"], first=first)
        self.assertTrue(all(len(c["number"]) == 9 for c in r["ordered"]))

    def test_comments_in_several_columns(self):
        extra = [("Additional comments", lambda c: journal([e for e in c[21] if e[2] == "customer"])),
                 ("Work notes", lambda c: journal([e for e in c[21] if e[2] == "internal"]))]
        r, html, text, _ = self.run_review(ALL, extra=extra)
        c = next(x for x in r["ordered"] if x["number"] == "CS0000101")
        entries = cp.parse_journal(c["journal"])
        self.assertEqual(len(entries), 3, "an entry in several columns must count once")

    def test_old_priority_and_name_columns(self):
        rename = {"Highest Priority": "Priority", "Assigned to": "Name"}
        r, html, text, _ = self.run_review([h for h in ALL], transform=lambda h: rename.get(h, h))
        self.assertEqual(next(c for c in r["ordered"] if c["number"] == "CS0000105")["owner"], "Raj Owner")

    def test_compared_with_older_export_of_other_columns(self):
        r, html, text, _ = self.run_review(ALL, previous=["Number", "Short description", "State"])
        self.assert_sane(r, html, text)


    def test_jira_links(self):
        # a Jira link inside a comment (the display text shortens links, the ticket must survive)
        CASES.append(("CS0000107", "P3", "Linked by URL", "Pending Acme", "", 5, 1, "Pat Customer", "Sam Owner",
                      "", "", "", "", "1", "", "", "", "", "", "", "", [
                          (1, "Sam Owner", "internal", "Raised https://acme.atlassian.net/browse/ENG-55555 today")]))
        CASES.append(("CS0000108", "P3", "Has Jira but no ticket in export", "Pending Acme", "", 5, 1, "", "",
                      "", "", "", "", "1", "", "", "", "", "", "", "", [(1, "Sam Owner", "internal", "Uses UTF-8 and ISO-8601")]))
        try:
            r, html, text, _ = self.run_review(ALL)
        finally:
            del CASES[-2:]
        self.assertIn('href="https://acme.atlassian.net/browse/ENG-12345"', html)
        self.assertIn('href="https://acme.atlassian.net/browse/ENG-55555"', html)
        c = next(x for x in r["ordered"] if x["number"] == "CS0000108")
        self.assertEqual(cp.jira_of(c), ("", ""))          # UTF-8 / ISO-8601 are not tickets
        self.assertEqual(cp.engineering(c), "Engineering")
        self.assertIsNone(cp.jira_url("UTF-8"))


class ContractUsage(unittest.TestCase):
    """Limited case types counted per contract year, across every saved export."""

    def setUp(self):
        cp.apply_profile(EXAMPLE_PROFILE)          # Ask an Expert: 80, Go Live Support: 2 per year
        self.root = tempfile.mkdtemp(prefix="callprep-contract-")
        cp.set_root(self.root)
        self.folder = os.path.join(self.root, "exports", "acme")
        os.makedirs(self.folder)

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def export(self, day, rows):
        """rows: (number, type, state, opened days before `day`)."""
        path = os.path.join(self.folder, "acme_tickets-%s.csv" % day)
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["Number", "Short description", "State", "Opened", "Success Package Request Type"])
            for number, kind, state, ago in rows:
                opened = dt.datetime.combine(day - dt.timedelta(days=ago), dt.time(9))
                w.writerow([number, "Case " + number, state, opened.strftime("%Y-%m-%d %H:%M:%S"), kind])

    def test_contract_year(self):
        start = dt.date(2025, 3, 1)
        self.assertIsNone(cp.contract_year(start, dt.date(2025, 2, 28)))
        self.assertEqual(cp.contract_year(start, dt.date(2026, 2, 28)), (dt.date(2025, 3, 1), dt.date(2026, 2, 28)))
        self.assertEqual(cp.contract_year(start, dt.date(2026, 3, 1)), (dt.date(2026, 3, 1), dt.date(2027, 2, 28)))

    def test_counts_across_exports_and_this_year_only(self):
        today = TODAY
        start = today - dt.timedelta(days=100)
        # earlier export: an Ask an Expert case that closes (drops out) before the next export
        self.export(today - dt.timedelta(days=14), [("CS0000201", "Ask an Expert", "Pending Acme", 20),
                                                    ("CS0000202", "Go Live Support", "Pending Acme", 5)])
        self.export(today, [("CS0000202", "Go Live Support", "Pending Acme", 19),
                            ("CS0000203", "Ask an Expert", "Pending Customer", 3),
                            ("CS0000204", "Ask an Expert", "Pending Acme", 150),     # before this contract year
                            ("CS0000205", "", "Pending Acme", 2)])                   # Support: unlimited, not counted
        cp.save_contract("acme", {"start": start.isoformat()})
        u = cp.contract_usage("acme", today)
        rows = {r["type"]: r for r in u["rows"]}
        self.assertEqual(rows["Ask an Expert"]["cases"], ["CS0000201", "CS0000203"])
        self.assertEqual(rows["Go Live Support"]["used"], 1)
        self.assertFalse(u["exact"])                      # open cases only: a minimum
        self.assertEqual(cp.usage_text(rows["Ask an Expert"], u["exact"]), "at least 2 of 80 used")
        self.assertNotIn("Support", rows)

    def test_exact_when_export_has_closed_cases_and_no_resolved_flood(self):
        today = TODAY
        self.export(today - dt.timedelta(days=7), [("CS0000301", "Go Live Support", "Pending Acme", 30)])
        self.export(today, [("CS0000301", "Go Live Support", "Resolved", 37),         # closed since last export
                            ("CS0000302", "Go Live Support", "Closed", 60),           # closed long ago
                            ("CS0000303", "Go Live Support", "Pending Acme", 2)])
        cp.save_contract("acme", {"start": (today - dt.timedelta(days=200)).isoformat()})
        r = cp.load_review("acme")
        row = next(x for x in r["usage"]["rows"] if x["type"] == "Go Live Support")
        self.assertTrue(r["usage"]["exact"])
        self.assertEqual(row["used"], 3)
        self.assertEqual(cp.usage_text(row, True), "3 of 2 used, 1 over")
        self.assertEqual([c["number"] for c in r["resolved"]], ["CS0000301"])   # not the old closed case
        base, text = cp.write_outputs(r, {})
        html = open(base + ".html", encoding="utf-8").read()
        self.assertIn("Contract year", html)
        self.assertIn("3 of 2 used, 1 over", text)
        self.assertTrue(any("Go Live Support 3 of 2 used" in p for p in cp.prep_list(r, {})))

    def test_no_contract_no_usage(self):
        self.export(TODAY, [("CS0000401", "Ask an Expert", "Pending Acme", 3)])
        self.assertIsNone(cp.load_review("acme")["usage"])

def run_on_real_export(path, rounds=40):
    """Random column subsets of a real export (kept local: nothing is written outside a temp folder)."""
    rows = cp.read_xlsx(path) if path.lower().endswith((".xlsx", ".xlsm")) else cp.read_csv(path)
    header, body = rows[0], rows[1:]
    num = next(i for i, h in enumerate(header) if cp._key(h) == "number")
    rnd = random.Random(7)
    root = tempfile.mkdtemp(prefix="callprep-real-")
    cp.set_root(root)
    folder = os.path.join(root, "exports", "real")
    os.makedirs(folder)
    try:
        for i in range(rounds):
            keep = sorted({num} | set(rnd.sample(range(len(header)), rnd.randint(0, len(header)))))
            rnd.shuffle(keep)
            out = os.path.join(folder, "real_tickets-%s.csv" % TODAY)
            with open(out, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow([header[k] for k in keep])
                for row in body:
                    w.writerow([row[k] if k < len(row) else "" for k in keep])
            r = cp.load_review("real")
            base, text = cp.write_outputs(r, {})
            cp.prep_list(r, {})
            for c in r["ordered"]:
                cp.case_flags(c, r["as_of"])
        print("real export: %d random column mixes OK (%d columns, %d rows)" % (rounds, len(header), len(body)))
    finally:
        shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    real = [a for a in sys.argv[1:] if not a.startswith("-")]
    for path in real:
        run_on_real_export(path)
    unittest.main(argv=[sys.argv[0]] + [a for a in sys.argv[1:] if a.startswith("-")], verbosity=1)
