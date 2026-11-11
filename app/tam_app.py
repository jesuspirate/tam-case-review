#!/usr/bin/env python3
"""
TAM Case Review - desktop app around callprep.py.

Runs a small server on 127.0.0.1 and shows the UI (ui.html) in an app window
(pywebview). If the window library isn't available it opens the default browser
instead. All review logic lives in callprep.py; this file only moves data
between it and the UI.
"""

import datetime as dt
import json
import os
import re
import secrets
import subprocess
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE = getattr(sys, "_MEIPASS", HERE)          # PyInstaller unpacks data files here
sys.path.insert(0, os.path.dirname(HERE))        # dev: callprep.py sits one folder up
import callprep as cp                            # noqa: E402

APP_NAME = "TAM Case Review"
VERSION = "1.0"


# --- settings ---------------------------------------------------------------
def config_dir():
    if sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    elif os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.path.expanduser("~/.config")
    return os.path.join(base, APP_NAME)


CONFIG = os.path.join(config_dir(), "config.json")


def load_profile():
    """Organization profile: $TAM_PROFILE, then one saved in the app's settings folder,
    then the one bundled into the app at build time, then profile.json beside callprep.py."""
    for path in (os.environ.get("TAM_PROFILE"), os.path.join(config_dir(), "profile.json"),
                 os.path.join(BUNDLE, "profile.json"), os.path.join(os.path.dirname(HERE), "profile.json")):
        if path and cp.load_profile(path):
            return path
    return None


def default_data_dir():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    legacy = os.path.join(docs, "tam-toolkit")          # the command-line setup from the README
    if os.path.isdir(os.path.join(legacy, "exports")):
        return legacy
    return os.path.join(docs, APP_NAME)


def load_config():
    try:
        with open(CONFIG, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def save_config(cfg):
    os.makedirs(os.path.dirname(CONFIG), exist_ok=True)
    with open(CONFIG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2)


def use_data_dir(path):
    path = os.path.abspath(os.path.expanduser(path))
    for sub in ("exports", "notes", "out"):
        os.makedirs(os.path.join(path, sub), exist_ok=True)
    cp.set_root(path)
    return path


# --- shaping data for the UI -------------------------------------------------
def iso(d):
    return d.isoformat() if d else None


def account_summary(account):
    files = cp.exports_for(account)
    latest = files[-1] if files else None
    open_count = None
    if latest:
        try:
            cases, _ = cp.load_export(latest)
            open_count = sum(1 for c in cases if not cp.is_closed(c))
        except (SystemExit, Exception):        # an unreadable export shouldn't hide the account
            open_count = None
    return {"id": account, "title": cp.display_name(account), "exports": len(files),
            "latest": os.path.basename(latest) if latest else None,
            "latest_date": iso(cp.export_date(latest)) if latest else None,
            "open": open_count}


def note_line(e):
    """The notes line 'Use this' saves for a ServiceNow comment (same format as Enter in Terminal)."""
    return "%d/%d - %s %s" % (e["when"].month, e["when"].day, cp.COPY_TAGS[e["kind"]],
                              cp.condense(e["text"]))


def entry_json(line):
    when, text = cp.split_entry(line)
    return {"raw": line, "date": iso(when), "text": cp.page_text(text), "copied": cp.is_copy(text),
            "internal": text.startswith(cp.COPY_TAGS["internal"])}


def case_json(c, r, notes):
    as_of = r["as_of"]
    journal = cp.parse_journal(c["journal"])
    jira = cp.jira_of(c)
    comments = []
    for e in cp.latest_comments(journal):
        comments.append({"kind": e["kind"], "when": iso(e["when"].date()) if e["when"] else None,
                         "author": e["author"], "text": e["text"],
                         "line": note_line(e) if e["when"] and e["text"] else None})
    entries = notes.get(c["number"], [])
    today = dt.date.today()
    top_when = cp.split_entry(entries[0])[0] if entries else None
    return {
        "number": c["number"], "priority": c["priority"], "type": c["type"],
        "type_color": cp.type_color(c["type"]), "description": c["description"],
        "state": c["state"], "state_was": r["state_was"].get(c["number"]),
        "substate": c["substate"], "contact": c["contact"], "owner": c["owner"],
        "group": c["group"], "opened": iso(c["opened_dt"].date()) if c["opened_dt"] else None,
        "age": cp.days_between(c["opened_dt"], as_of),
        "last_activity": iso(c["updated_dt"].date()) if c["updated_dt"] else None,
        "jira": list(jira) if jira else None, "jira_url": cp.jira_url(jira[0]) if jira else None,
        "comments": comments,
        "notes": [entry_json(x) for x in entries],
        "fill": cp.autofill_line(c, notes),
        "needs": cp.needs_update(c, notes, r["prev_cases"], r["state_was"], as_of),
        "updated_today": bool(top_when and top_when >= today),
        "new": r["prev_cases"] is not None and c["number"] not in r["prev_cases"],
        "side": cp.side(c),
        # optional columns (blank when the export doesn't have them)
        "area": cp.area(c), "tenant": c.get("tenant") or "", "project": c.get("project") or "",
        "engineering": cp.engineering(c), "flags": cp.case_flags(c, as_of),
        "details": c.get("details") or "", "impact": c.get("impact") or "", "urgency": c.get("urgency") or "",
        "action_status": c.get("action_status") or "",
        "follow_up": iso(c["follow_up_dt"].date()) if c.get("follow_up_dt") else None,
    }


def usage_json(u):
    if not u:
        return None
    return {"first": iso(u["first"]), "last": iso(u["last"]), "exact": u["exact"],
            "rows": [dict(row, text=cp.usage_text(row, u["exact"])) for row in u["rows"]]}


def set_contract(account, start):
    """Save (or with a blank start, remove) an account's contract start date."""
    account = safe_account(account)
    if not start:
        cp.save_contract(account, None)
        return {"ok": True}
    try:
        dt.date.fromisoformat(start)
    except ValueError:
        raise ValueError("The contract start date doesn't look right: %s" % start)
    contract = dict(cp.load_contracts().get(account) or {}, start=start)
    cp.save_contract(account, contract)
    return {"ok": True}


def review_json(account, by_priority=False):
    r = cp.load_review(account, by_action=not by_priority)
    notes = cp.load_notes(account)
    return {
        "account": account, "title": r["title"], "subtitle": r["subtitle"],
        "as_of": iso(r["as_of"]), "current": os.path.basename(r["current"]),
        "previous": os.path.basename(r["previous"]) if r["previous"] else None,
        "prev_date": iso(r["prev_date"]),
        "counts": {"open": len(r["open_cases"]), "new": len(r["new_ids"]),
                   "resolved": len(r["resolved"]), "changed": len(r["state_was"]),
                   "compared": r["prev_cases"] is not None},
        "groups": [{"type": t, "color": cp.type_color(t), "count": len(cs)} for t, cs in r["groups"]],
        "has_state": "state" in r["mapping"],
        "missing": [{"column": col, "why": why} for col, why in cp.missing_columns(r["mapping"], r["open_cases"])],
        "contract": cp.load_contracts().get(account),
        "usage": usage_json(r.get("usage")),
        "cases": [case_json(c, r, notes) for c in r["ordered"]],
    }


def safe_account(name):
    name = re.sub(r"[^a-z0-9-]+", "-", (name or "").strip().lower()).strip("-")
    if not name:
        raise ValueError("Give the account a name (letters and numbers).")
    return name


def import_export(src_name, data, account, date_text=None):
    """Save an export as exports/<account>/<account>_tickets-YYYY-MM-DD.ext and check it reads."""
    account = safe_account(account)
    ext = os.path.splitext(src_name)[1].lower()
    if ext not in cp.EXTS:
        raise ValueError("That file is %s. Export from ServiceNow as Excel (.xlsx) or CSV."
                         % (ext or "not a spreadsheet"))
    when = None
    if date_text:
        try:
            when = dt.date.fromisoformat(date_text)
        except ValueError:
            raise ValueError("The export date doesn't look right: %s" % date_text)
    when = when or cp.name_date(src_name) or dt.date.today()
    folder = os.path.join(cp.EXPORT_DIR, account)
    existed = os.path.isdir(folder)
    os.makedirs(folder, exist_ok=True)
    for name in os.listdir(folder):              # the same file added twice: keep the one copy
        path = os.path.join(folder, name)
        if cp._is_export(path) and os.path.getsize(path) == len(data):
            with open(path, "rb") as fh:
                if fh.read() == data:
                    cases, _ = cp.load_export(path)
                    return {"account": account, "saved_as": name, "date": iso(cp.export_date(path)),
                            "cases": len(cases), "duplicate": True,
                            "open": sum(1 for c in cases if not cp.is_closed(c))}
    stem = "%s_tickets-%s" % (account, when.isoformat())
    dest, n = os.path.join(folder, stem + ext), 2
    while os.path.exists(dest):
        dest, n = os.path.join(folder, "%s (%d)%s" % (stem, n, ext)), n + 1
    tmp = os.path.join(folder, "~importing-" + os.path.basename(dest))   # "~" files are never read as exports
    with open(tmp, "wb") as fh:
        fh.write(data)
    try:
        cases, _ = cp.load_export(tmp)
    except BaseException as e:
        os.remove(tmp)
        if not existed and not os.listdir(folder):
            os.rmdir(folder)
        if isinstance(e, SystemExit):
            raise ValueError(str(e).replace(tmp, src_name).replace(os.path.basename(tmp), src_name))
        raise ValueError("Couldn't read %s. Is it a ServiceNow Excel export?" % src_name)
    os.replace(tmp, dest)
    return {"account": account, "saved_as": os.path.basename(dest), "date": when.isoformat(),
            "cases": len(cases), "open": sum(1 for c in cases if not cp.is_closed(c))}


def save_update(account, number, text=None, line=None):
    notes = cp.load_notes(account)
    if line is None:
        text = " ".join((text or "").split())
        if not text:
            raise ValueError("Type an update first.")
        today = dt.date.today()
        line = "%d/%d - %s" % (today.month, today.day, text)
    entries = notes.setdefault(number, [])
    if not entries or entries[0] != line:
        entries.insert(0, line)
    cp.save_notes(account, notes, cp.display_name(account))


def remove_latest(account, number):
    notes = cp.load_notes(account)
    if notes.get(number):
        notes[number].pop(0)
        cp.save_notes(account, notes, cp.display_name(account))


def build(account, by_priority=False, open_page=True):
    r = cp.load_review(account, by_action=not by_priority)
    notes = cp.load_notes(account)
    base, text = cp.write_outputs(r, notes)
    copied = cp.copy_to_clipboard(text)
    if open_page:
        cp.open_file(base + ".html")
    return {"html": base + ".html", "csv": base + ".csv", "txt": base + ".txt",
            "name": os.path.basename(base), "copied": copied, "prep": cp.prep_list(r, notes)}


def build_across():
    out = os.path.join(cp.OUT_DIR, "waiting-on-us-%s.html" % dt.date.today().isoformat())
    cp.across(open_page=False)
    cp.open_file(out)
    return {"html": out}


def inside_data(path):
    root = os.path.realpath(cp.ROOT)
    path = os.path.realpath(path)
    return path == root or path.startswith(root + os.sep)


# --- server -----------------------------------------------------------------
TOKEN = secrets.token_urlsafe(24)
LAST_SEEN = [time.time()]
WINDOW = [None]          # the pywebview window, for native dialogs


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        if os.environ.get("TAM_DEBUG"):
            sys.stderr.write("%s\n" % (fmt % args))

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _authorized(self):
        return secrets.compare_digest(self.headers.get("X-Token", ""), TOKEN)

    def do_GET(self):
        url = urlparse(self.path)
        if url.path == "/":
            q = parse_qs(url.query)
            if q.get("t", [""])[0] != TOKEN:
                return self._send(403, b"Forbidden", "text/plain")
            with open(os.path.join(BUNDLE, "ui.html"), "rb") as fh:
                html = fh.read().replace(b"__TOKEN__", TOKEN.encode())
            return self._send(200, html, "text/html; charset=utf-8")
        if not self._authorized():
            return self._send(403, {"error": "forbidden"})
        q = {k: v[0] for k, v in parse_qs(url.query).items()}
        self._api("GET", url.path, q)

    def do_POST(self):
        if not self._authorized():
            return self._send(403, {"error": "forbidden"})
        url = urlparse(self.path)
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        if url.path == "/api/import":
            q = {"name": unquote(self.headers.get("X-Filename", "")),
                 "account": unquote(self.headers.get("X-Account", "")),
                 "date": self.headers.get("X-Date", "")}
            return self._api("POST", url.path, q, raw)
        try:
            body = json.loads(raw or b"{}")
        except ValueError:
            return self._send(400, {"error": "bad request"})
        self._api("POST", url.path, body)

    def _api(self, method, path, q, raw=None):
        LAST_SEEN[0] = time.time()
        try:
            result = route(method, path, q, raw)
        except (ValueError, SystemExit) as e:
            return self._send(400, {"error": str(e)})
        except Exception as e:  # show it in the app rather than failing silently
            return self._send(500, {"error": "%s: %s" % (type(e).__name__, e)})
        if result is None:
            return self._send(404, {"error": "not found"})
        self._send(200, result)


def route(method, path, q, raw):
    if path == "/api/ping":
        return {"ok": True}
    if path == "/api/home":
        accounts = [account_summary(a) for a in cp.list_accounts()]
        return {"data_dir": cp.ROOT, "accounts": accounts, "version": VERSION,
                "native": WINDOW[0] is not None, "platform": sys.platform,
                "vendor": cp.VENDOR, "profile": cp.PROFILE_NAME,
                "allowances": cp.CONTRACT_ALLOWANCES}
    if path == "/api/review":
        return review_json(q["account"], q.get("by_priority") in ("1", True))
    if method != "POST":
        return None
    if path == "/api/import":
        return import_export(q["name"], raw, q["account"], q.get("date"))
    if path == "/api/import-path":
        src = q["path"]
        with open(src, "rb") as fh:
            data = fh.read()
        return import_export(os.path.basename(src), data, q["account"], q.get("date"))
    if path == "/api/contract":
        return set_contract(q["account"], q.get("start") or "")
    if path == "/api/note":
        save_update(q["account"], q["number"], q.get("text"), q.get("line"))
        return {"ok": True}
    if path == "/api/note-remove":
        remove_latest(q["account"], q["number"])
        return {"ok": True}
    if path == "/api/build":
        return build(q["account"], bool(q.get("by_priority")))
    if path == "/api/across":
        return build_across()
    if path == "/api/copy":
        p = q["path"]
        if not inside_data(p):
            raise ValueError("Not in the data folder.")
        with open(p, encoding="utf-8") as fh:
            return {"copied": cp.copy_to_clipboard(fh.read())}
    if path == "/api/open":
        target = q.get("path") or cp.ROOT
        if q.get("folder") in ("exports", "notes", "out"):
            target = os.path.join(cp.ROOT, q["folder"], q.get("account", ""))
        if not inside_data(target):
            raise ValueError("Not in the data folder.")
        if q.get("reveal") and sys.platform == "darwin":
            subprocess.run(["open", "-R", target], check=False)
        elif q.get("reveal") and os.name == "nt":
            subprocess.run('explorer /select,"%s"' % os.path.normpath(target), check=False)
        else:
            cp.open_file(target if os.path.exists(target) else cp.ROOT)
        return {"ok": True}
    if path == "/api/data-dir":
        cfg = load_config()
        cfg["data_dir"] = use_data_dir(q["path"])
        save_config(cfg)
        return {"data_dir": cp.ROOT}
    if path == "/api/open-url":
        url = q.get("url") or ""
        prefix = (cp.JIRA_BROWSE_URL or "").split("%s")[0]
        if not prefix or not url.startswith(prefix) or not cp.is_ticket(url[len(prefix):]):
            raise ValueError("Only Jira ticket links can be opened from here.")
        import webbrowser
        webbrowser.open(url)
        return {"ok": True}
    if path == "/api/pick":
        return pick(q.get("kind", "file"))
    return None


def pick(kind):
    """Native open dialog (app window only). Returns {"path": ...} or {"path": None}."""
    import webview
    win = WINDOW[0]
    if win is None:
        raise ValueError("Native dialogs need the app window.")
    if kind == "folder":
        res = win.create_file_dialog(webview.FileDialog.FOLDER, directory=cp.ROOT)
    else:
        res = win.create_file_dialog(
            webview.FileDialog.OPEN, allow_multiple=False,
            directory=os.path.join(os.path.expanduser("~"), "Downloads"),
            file_types=("ServiceNow export (*.xlsx;*.xlsm;*.csv)", "All files (*.*)"))
    if not res:
        return {"path": None}
    return {"path": res[0] if isinstance(res, (list, tuple)) else res}


def serve():
    server = ThreadingHTTPServer(("127.0.0.1", int(os.environ.get("TAM_PORT") or 0)), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server, "http://127.0.0.1:%d/?t=%s" % (server.server_address[1], TOKEN)


def main():
    load_profile()
    cfg = load_config()
    data_dir = os.environ.get("TAM_DATA_DIR") or cfg.get("data_dir") or default_data_dir()
    try:
        use_data_dir(data_dir)
    except OSError:
        use_data_dir(default_data_dir())
    server, url = serve()
    if "--browser" not in sys.argv:
        try:
            import webview
            WINDOW[0] = webview.create_window(APP_NAME, url, width=1240, height=860,
                                              min_size=(900, 620), text_select=True)
            webview.start()
            server.shutdown()
            return
        except Exception as e:  # no window library on this machine: fall back to the browser
            WINDOW[0] = None
            sys.stderr.write("App window unavailable (%s); opening the browser.\n" % e)
    if not os.environ.get("TAM_NO_OPEN"):
        import webbrowser
        webbrowser.open(url)
    print("%s is running at %s\nClose the browser tab to quit (or press Ctrl+C)." % (APP_NAME, url))
    try:
        while time.time() - LAST_SEEN[0] < 90:   # the page pings every 20s; quit once it's gone
            time.sleep(5)
    except KeyboardInterrupt:
        pass
    server.shutdown()


if __name__ == "__main__":
    main()
