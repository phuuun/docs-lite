#!/usr/bin/env python3
"""Documents: a tiny plain-text writing app. Files are .txt/.md under ~/Documents (the root is fixed)."""
import gi, json, pathlib, sys, urllib.parse
gi.require_version("Gtk", "3.0")
gi.require_version("WebKit2", "4.1")
from gi.repository import Gio, Gtk, WebKit2

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import docx_html

HERE = pathlib.Path(__file__).resolve().parent
ROOT = (pathlib.Path.home() / "Documents").resolve()
ROOT.mkdir(exist_ok=True)
EXT = (".txt", ".md")

def safe(rel):
    """Relative path -> absolute path, refusing anything that escapes ROOT."""
    p = (ROOT / rel).resolve()
    if p != ROOT and ROOT not in p.parents: raise ValueError("outside Documents")
    return p

def rel(p): return "" if p == ROOT else str(p.relative_to(ROOT))

def clean(n):
    n = "".join(c for c in n if c not in '/\\\0').strip().lstrip(".")
    if not n: raise ValueError("empty name")
    return n

def unique(d, name):
    p, i = d / name, 2
    while p.exists(): p = d / f"{pathlib.Path(name).stem} ({i}){pathlib.Path(name).suffix}"; i += 1
    return p

def import_file(path, d=ROOT):
    """Copy a .docx/.html/.txt into folder d as .txt (original untouched); returns its rel path."""
    p = pathlib.Path(path)
    ext = p.suffix.lower()
    if ext == ".docx": text = docx_html.html_to_text(docx_html.convert(p))
    elif ext in (".html", ".htm"): text = docx_html.html_to_text(p.read_text(errors="replace"))
    else: text = p.read_text(errors="replace")
    out = unique(d, clean(p.stem) + (ext if ext in EXT else ".txt"))
    out.write_text(text)
    return rel(out)

def pick(d):
    dlg = Gtk.FileChooserDialog(title="Import", parent=win, action=Gtk.FileChooserAction.OPEN)
    dlg.add_buttons("Cancel", Gtk.ResponseType.CANCEL, "Import", Gtk.ResponseType.OK)
    f = Gtk.FileFilter(); f.set_name("Documents (.docx .html .txt .md)")
    for e in ("docx", "html", "htm", "txt", "md"): f.add_pattern(f"*.{e}")
    dlg.add_filter(f)
    path = dlg.get_filename() if dlg.run() == Gtk.ResponseType.OK else None
    dlg.destroy()
    return import_file(path, d) if path else None

def handle(m):
    op = m["op"]
    p = safe(m.get("path", ""))
    if op == "ls":  # folders first, then files, hidden ones skipped
        items = [{"name": e.name, "dir": e.is_dir()} for e in p.iterdir()
                 if not e.name.startswith(".") and (e.is_dir() or e.suffix.lower() in EXT)]
        return sorted(items, key=lambda i: (not i["dir"], i["name"].lower()))
    if op == "load": return p.read_text(errors="replace")
    if op == "save":
        tmp = p.with_name(p.name + ".tmp")
        tmp.write_text(m["text"]); tmp.replace(p)
        return True
    if op == "new":  # p is the parent folder
        name = clean(m["name"])
        if m["dir"]:
            q = p / name; q.mkdir()
        else:
            if not name.lower().endswith(EXT): name += ".txt"
            q = p / name; q.touch(exist_ok=False)
        return rel(q)
    if op == "rename":
        q = p.with_name(clean(m["name"]))
        if p.is_file() and not q.name.lower().endswith(EXT): q = q.with_name(q.name + p.suffix)
        if q.exists(): raise FileExistsError(q.name)
        p.rename(q)
        return rel(q)
    if op == "delete":
        Gio.File.new_for_path(str(p)).trash(None)  # recoverable from the system trash
        return True
    if op == "import": return pick(p)

def on_msg(mgr, res):
    v = res.get_js_value() if hasattr(res, "get_js_value") else res
    m = json.loads(v.to_string())
    try: out = {"ok": handle(m)}
    except Exception as e: out = {"err": f"{type(e).__name__}: {e}"}
    view.evaluate_javascript(f"onPy({m['id']},{json.dumps(out)})", -1, None, None, None, None, None)

mgr = WebKit2.UserContentManager()
mgr.register_script_message_handler("bridge")
mgr.connect("script-message-received::bridge", on_msg)
view = WebKit2.WebView.new_with_user_content_manager(mgr)
view.get_settings().set_enable_developer_extras(False)
frag = ""
if len(sys.argv) > 1:
    try:
        a = pathlib.Path(sys.argv[1]).resolve()
        frag = "#" + urllib.parse.quote(rel(a) if a.suffix.lower() in EXT and ROOT in a.parents else import_file(a))
    except Exception as e: print("open failed:", e, file=sys.stderr)
view.load_uri((HERE / "editor.html").as_uri() + frag)

win = Gtk.Window(title="Documents")
win.set_default_size(1200, 800)
win.add(view)
win.connect("destroy", Gtk.main_quit)
win.show_all()
Gtk.main()
