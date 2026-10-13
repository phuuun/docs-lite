"""Minimal stdlib .docx -> HTML (paragraphs, runs, headings, lists, tables, links, images)."""
import base64, html, zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
MIME = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "svg": "image/svg+xml", "bmp": "image/bmp"}
HEAD = {"Title": "h1", "Heading1": "h2", "Heading2": "h3"}

def on(el):  # toggle property like <w:b/> or <w:b w:val="0"/>
    return el is not None and el.get(W + "val", "true") not in ("0", "false", "none")

def convert(path):
    z = zipfile.ZipFile(path)
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("word/_rels/document.xml.rels"))}
    lists = {}  # (numId, ilvl) -> 'ol' | 'ul'
    if "word/numbering.xml" in z.namelist():
        n = ET.fromstring(z.read("word/numbering.xml"))
        fmt = {(a.get(W + "abstractNumId"), l.get(W + "ilvl")): l.find(W + "numFmt").get(W + "val")
               for a in n.iter(W + "abstractNum") for l in a.iter(W + "lvl") if l.find(W + "numFmt") is not None}
        for m in n.iter(W + "num"):
            a = m.find(W + "abstractNumId").get(W + "val")
            for (aid, lvl), f in fmt.items():
                if aid == a:
                    lists[(m.get(W + "numId"), lvl)] = "ul" if f == "bullet" else "ol"

    def image(d):
        b = d.find(".//" + A + "blip")
        ext = d.find(".//" + WP + "extent")
        t = rels.get(b.get(R + "embed")) if b is not None else None
        if not t or "word/" + t not in z.namelist():
            return ""
        data = base64.b64encode(z.read("word/" + t)).decode()
        w = f' style="width:{int(ext.get("cx")) // 9525}px"' if ext is not None else ""
        return f'<img src="data:{MIME.get(t.rsplit(".", 1)[-1].lower(), "image/png")};base64,{data}"{w}>'

    def run(r):
        out = ""
        for c in r:
            if c.tag == W + "t": out += html.escape(c.text or "")
            elif c.tag == W + "tab": out += "&emsp;"
            elif c.tag == W + "br": out += "<br>"
            elif c.tag == W + "drawing": out += image(c)
        p = r.find(W + "rPr")
        if p is None or not out:
            return out
        css = ""
        col = p.find(W + "color"); sz = p.find(W + "sz")
        if col is not None and col.get(W + "val") not in (None, "auto"): css += f"color:#{col.get(W + 'val')};"
        if sz is not None: css += f"font-size:{int(sz.get(W + 'val')) / 2}pt;"
        if css: out = f'<span style="{css}">{out}</span>'
        for tag, el in (("b", "b"), ("i", "i"), ("u", "u"), ("s", "strike")):
            if on(p.find(W + el)): out = f"<{tag}>{out}</{tag}>"
        return out

    def para(p):
        txt = ""
        for c in p:
            if c.tag == W + "r": txt += run(c)
            elif c.tag == W + "hyperlink":
                inner = "".join(run(r) for r in c.iter(W + "r"))
                href = rels.get(c.get(R + "id"))
                txt += f'<a href="{html.escape(href)}">{inner}</a>' if href else inner
        return txt

    def blocks(parent):
        out, stack = "", []  # stack: open list tags
        def close(to=0):
            nonlocal out
            while len(stack) > to: out += f"</{stack.pop()}>"
        for el in parent:
            if el.tag == W + "p":
                pp = el.find(W + "pPr")
                st = pp.find(W + "pStyle") if pp is not None else None
                num = pp.find(W + "numPr") if pp is not None else None
                jc = pp.find(W + "jc") if pp is not None else None
                align = {"center": "center", "right": "right", "both": "justify"}.get(jc.get(W + "val")) if jc is not None else None
                style = f' style="text-align:{align}"' if align else ""
                txt = para(el)
                if num is not None and num.find(W + "numId") is not None and num.find(W + "numId").get(W + "val") != "0":
                    lvl = num.find(W + "ilvl"); lvl = lvl.get(W + "val") if lvl is not None else "0"
                    kind = lists.get((num.find(W + "numId").get(W + "val"), lvl), "ul")
                    depth = int(lvl) + 1
                    close(depth)
                    while len(stack) < depth: stack.append(kind); out += f"<{kind}>"
                    out += f"<li{style}>{txt}</li>"
                    continue
                close()
                tag = HEAD.get(st.get(W + "val"), "p") if st is not None else "p"
                if tag == "p" and st is not None and st.get(W + "val").startswith("Heading"): tag = "h4"
                out += f"<{tag}{style}>{txt or '<br>'}</{tag}>"
            elif el.tag == W + "tbl":
                close()
                out += "<table>"
                for tr in el.iter(W + "tr"):
                    out += "<tr>"
                    for tc in tr.findall(W + "tc"):
                        gs = tc.find(W + "tcPr/" + W + "gridSpan")
                        span = f' colspan="{gs.get(W + "val")}"' if gs is not None else ""
                        out += f"<td{span}>{blocks(tc)}</td>"
                    out += "</tr>"
                out += "</table><p><br></p>"
        close()
        return out

    return blocks(ET.fromstring(z.read("word/document.xml")).find(W + "body"))

if __name__ == "__main__":  # self-check: build a tiny docx and convert it
    import io, sys
    if len(sys.argv) > 1:
        print(convert(sys.argv[1])); sys.exit()
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    doc = (f'<w:document {ns}><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>'
           '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>'
           '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", '<Relationships xmlns="x"/>')
    out = convert(buf)
    assert "<h2>Hi</h2>" in out and "<b>bold</b>" in out and "<td><p>c</p></td>" in out, out
    print("ok")


def html_to_text(src):
    """HTML fragment -> plain text: blocks become paragraphs, list items '- ', images dropped."""
    from html.parser import HTMLParser
    class P(HTMLParser):
        def __init__(s): super().__init__(convert_charrefs=True); s.out, s.skip = [], 0
        def handle_starttag(s, t, a):
            if t in ("script", "style"): s.skip += 1
            elif t == "br": s.out.append("\n")
            elif t == "li": s.out.append("\n- ")
            elif t in ("td", "th"): s.out.append("\t")
            elif t in ("p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "tr", "ul", "ol", "table"): s.out.append("\n")
        def handle_endtag(s, t):
            if t in ("script", "style"): s.skip -= 1
            elif t in ("p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "tr", "ul", "ol", "table"): s.out.append("\n")
        def handle_data(s, d):
            if not s.skip: s.out.append(d)
    import re
    p = P(); p.feed(src)
    return re.sub(r"\n{3,}", "\n\n", "".join(p.out)).strip() + "\n"
