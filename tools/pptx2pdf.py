"""Render a pptxgenjs-built .pptx to PDF without PowerPoint or LibreOffice: every slide's shapes,
text, tables, lines and pictures are laid out as positioned HTML at their exact EMU geometry,
then printed with headless Chrome, one 13.333 x 7.5 in page per slide.
Calibri / Cambria are drawn with their metric-compatible twins Carlito / Caladea, so text wraps
as it does in PowerPoint. Covers the subset our build scripts emit (not a general converter).
Usage: python pptx2pdf.py deck.pptx out.pdf [--slides 1-10] [--png-dir DIR]"""
import base64, html, pathlib, re, subprocess, sys, tempfile, zipfile
from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
EMU_PX = 96 / 914400                      # 96 px per inch
FONT_MAP = {"Calibri": "'Carlito', Calibri, Arial, sans-serif", "Cambria": "'Caladea', Cambria, Georgia, serif",
            "Courier New": "'Courier New', Courier, monospace", "Arial": "Arial, sans-serif"}
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

src, out = pathlib.Path(sys.argv[1]).resolve(), pathlib.Path(sys.argv[2]).resolve()
out.parent.mkdir(parents=True, exist_ok=True)
args = sys.argv[3:]
only = None
if "--slides" in args:
    a, b = args[args.index("--slides") + 1].split("-"); only = range(int(a), int(b) + 1)
png_dir = pathlib.Path(args[args.index("--png-dir") + 1]).resolve() if "--png-dir" in args else None

z = zipfile.ZipFile(src)
X = lambda name: etree.fromstring(z.read(name))
theme = X("ppt/theme/theme1.xml")
SCHEME = {}
for slot in ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]:
    el = theme.find(f".//a:clrScheme/a:{slot}", NS)
    c = el.find("a:srgbClr", NS); c = c if c is not None else el.find("a:sysClr", NS)
    SCHEME[slot] = c.get("val") if c.get("val") and len(c.get("val")) == 6 else c.get("lastClr", "000000")
SCHEME.update({"tx1": SCHEME["dk1"], "bg1": SCHEME["lt1"], "tx2": SCHEME["dk2"], "bg2": SCHEME["lt2"]})
MAJOR = theme.find(".//a:majorFont/a:latin", NS).get("typeface")
MINOR = theme.find(".//a:minorFont/a:latin", NS).get("typeface")
pres = X("ppt/presentation.xml")
sz = pres.find("p:sldSz", NS); SW, SH = int(sz.get("cx")) * EMU_PX, int(sz.get("cy")) * EMU_PX

def rels(part):
    p = pathlib.PurePosixPath(part); rp = f"{p.parent}/_rels/{p.name}.rels"
    if rp not in z.namelist(): return {}
    out = {}
    for r in X(rp):
        t = r.get("Target"); out[r.get("Id")] = (r.get("Type").rsplit("/", 1)[-1], str(pathlib.PurePosixPath(p.parent, t)) if not t.startswith("http") else t)
    return {k: (ty, re.sub(r"[^/]+/\.\./", "", v)) for k, (ty, v) in out.items()}

def color(el):
    """hex colour from an element holding srgbClr/schemeClr (or None)"""
    if el is None: return None
    c = el.find("a:srgbClr", NS)
    if c is not None: return "#" + c.get("val")
    c = el.find("a:schemeClr", NS)
    if c is not None: return "#" + SCHEME.get(c.get("val"), "000000")
    return None

def emu(v): return int(v) * EMU_PX
def esc(t): return html.escape(t, quote=False)

def run_css(rpr, defaults):
    d = dict(defaults)
    if rpr is not None:
        if rpr.get("sz"): d["size"] = int(rpr.get("sz")) / 100
        if rpr.get("b") is not None: d["bold"] = rpr.get("b") == "1"
        if rpr.get("i") is not None: d["italic"] = rpr.get("i") == "1"
        if rpr.get("u") not in (None, "none"): d["underline"] = True
        if rpr.get("spc"): d["spc"] = int(rpr.get("spc")) / 100
        c = color(rpr.find("a:solidFill", NS))
        if c: d["color"] = c
        lat = rpr.find("a:latin", NS)
        if lat is not None and lat.get("typeface"):
            tf = lat.get("typeface"); d["font"] = MAJOR if tf == "+mj-lt" else MINOR if tf == "+mn-lt" else tf
    css = f"font-size:{d['size']}pt;color:{d.get('color', '#000')};font-family:{FONT_MAP.get(d['font'], repr(d['font']))};"
    css += "font-weight:700;" if d.get("bold") else "font-weight:400;"
    if d.get("italic"): css += "font-style:italic;"
    if d.get("underline"): css += "text-decoration:underline;"
    if d.get("spc"): css += f"letter-spacing:{d['spc']}pt;"
    return css, d

def paragraphs_html(txbody, defaults, slide_no):
    """txBody -> list of <p> html; defaults: dict(size,color,bold,font) inherited from lstStyle/placeholder"""
    lst = txbody.find("a:lstStyle/a:lvl1pPr", NS)
    base = dict(defaults)
    if lst is not None:
        dr = lst.find("a:defRPr", NS)
        if dr is not None: _, base = run_css(dr, base)
    out = []
    for p in txbody.findall("a:p", NS):
        ppr = p.find("a:pPr", NS)
        algn = (ppr.get("algn") if ppr is not None and ppr.get("algn") else (lst.get("algn") if lst is not None and lst.get("algn") else "l"))
        align = {"l": "left", "ctr": "center", "r": "right", "just": "justify"}.get(algn, "left")
        marL = emu(ppr.get("marL")) if ppr is not None and ppr.get("marL") else 0
        indent = emu(ppr.get("indent")) if ppr is not None and ppr.get("indent") else 0
        lh, after, before = 1.2, 0, 0
        if ppr is not None:
            ls = ppr.find("a:lnSpc/a:spcPct", NS)
            if ls is not None: lh = 1.2 * int(ls.get("val")) / 100000
            sa = ppr.find("a:spcAft/a:spcPts", NS)
            if sa is not None: after = int(sa.get("val")) / 100
            sb = ppr.find("a:spcBef/a:spcPts", NS)
            if sb is not None: before = int(sb.get("val")) / 100
        bullet = ppr is not None and ppr.find("a:buChar", NS) is not None
        spans, first_d = [], None
        for r in p:
            tag = etree.QName(r).localname
            if tag == "r":
                css, d = run_css(r.find("a:rPr", NS), base); first_d = first_d or d
                t = esc(r.findtext("a:t", default="", namespaces=NS))
                if r.find("a:rPr/a:hlinkClick", NS) is not None: css += "text-decoration:underline;"
                spans.append(f'<span style="{css}">{t}</span>')
            elif tag == "fld":
                css, d = run_css(r.find("a:rPr", NS), base); first_d = first_d or d
                t = str(slide_no) if r.get("type") == "slidenum" else esc(r.findtext("a:t", default="", namespaces=NS))
                spans.append(f'<span style="{css}">{t}</span>')
            elif tag == "br":
                spans.append("<br>")
        if first_d is None:
            er = p.find("a:endParaRPr", NS); _, first_d = run_css(er, base)
            spans.append("&#8203;")
        size = first_d["size"]
        style = f"margin:0;padding:0 0 0 {marL}px;text-indent:{indent}px;text-align:{align};line-height:{lh};margin-top:{before}pt;margin-bottom:{after}pt;font-size:{size}pt;"
        bul = ""
        if bullet:
            ch = ppr.find("a:buChar", NS).get("char")
            bcss, _ = run_css(None, first_d)
            bul = f'<span style="{bcss}display:inline-block;width:{-indent}px;text-indent:0">{esc(ch)}</span>'
        out.append(f'<p style="{style}">{bul}{"".join(spans)}</p>')
    return out

def shape_html(sp, slide_no, defaults_by_ph):
    """one p:sp -> (html div or None, svg fragment or None)"""
    xfrm = sp.find("p:spPr/a:xfrm", NS)
    ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if xfrm is None and ph is not None:
        lay = defaults_by_ph.get(ph.get("type"), {}); xfrm = lay.get("xfrm")
    if xfrm is None: return None, None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    x, y, w, h = emu(off.get("x")), emu(off.get("y")), emu(ext.get("cx")), emu(ext.get("cy"))
    spPr = sp.find("p:spPr", NS)
    geom = spPr.find("a:prstGeom", NS); prst = geom.get("prst") if geom is not None else "rect"
    if prst == "line":
        ln = spPr.find("a:ln", NS); c = color(ln.find("a:solidFill", NS)) or "#000"
        lw = max(1.0, emu(ln.get("w") or 9525))
        x1, y1, x2, y2 = x, y, x + w, y + h
        if xfrm.get("flipH") == "1": x1, x2 = x2, x1
        if xfrm.get("flipV") == "1": y1, y2 = y2, y1
        dash = ln.find("a:prstDash", NS); da = ' stroke-dasharray="6 4"' if dash is not None and dash.get("val") not in (None, "solid") else ""
        mk = f"m{c[1:]}"
        me = f' marker-end="url(#{mk})"' if ln.find("a:tailEnd", NS) is not None and ln.find("a:tailEnd", NS).get("type") not in (None, "none") else ""
        ms = f' marker-start="url(#{mk})"' if ln.find("a:headEnd", NS) is not None and ln.find("a:headEnd", NS).get("type") not in (None, "none") else ""
        return None, (c, f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" stroke="{c}" stroke-width="{lw:.2f}"{da}{me}{ms}/>')
    fill = color(spPr.find("a:solidFill", NS))
    ln = spPr.find("a:ln", NS); border = ""
    if ln is not None and ln.find("a:solidFill", NS) is not None:
        lc = color(ln.find("a:solidFill", NS)); lw = max(0.75, emu(ln.get("w") or 9525))
        dash = ln.find("a:prstDash", NS); style = "dashed" if dash is not None and dash.get("val") not in (None, "solid") else "solid"
        border = f"border:{lw:.2f}px {style} {lc};"
    radius = ""
    if prst == "roundRect":
        gd = geom.find("a:avLst/a:gd", NS); adj = int(gd.get("fmla").split()[-1]) / 100000 if gd is not None else 0.16667
        radius = f"border-radius:{adj * min(w, h):.2f}px;"
    elif prst == "ellipse": radius = "border-radius:50%;"
    body = sp.find("p:txBody", NS)
    inner, valign = "", "flex-start"
    pad = "0"
    if body is not None and "".join(body.itertext()).strip():
        bp = body.find("a:bodyPr", NS)
        ins = [emu(bp.get(k)) if bp is not None and bp.get(k) is not None else d for k, d in (("tIns", 4.8), ("rIns", 9.6), ("bIns", 4.8), ("lIns", 9.6))]
        pad = f"{ins[0]:.2f}px {ins[1]:.2f}px {ins[2]:.2f}px {ins[3]:.2f}px"
        anchor = (bp.get("anchor") if bp is not None else None) or "t"
        valign = {"t": "flex-start", "ctr": "center", "b": "flex-end"}.get(anchor, "flex-start")
        defaults = {"size": 18, "color": "#" + SCHEME["tx1"], "bold": False, "font": MINOR}
        if ph is not None:
            lay = defaults_by_ph.get(ph.get("type"), {})
            defaults.update(lay.get("defaults", {}))
        inner = "".join(paragraphs_html(body, defaults, slide_no))
    bg = f"background:{fill};" if fill else ""
    div = (f'<div style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;width:{w:.2f}px;height:{h:.2f}px;box-sizing:border-box;{bg}{border}{radius}'
           f'padding:{pad};display:flex;flex-direction:column;justify-content:{valign};overflow:visible">'
           f'<div>{inner}</div></div>')
    return div, None

def table_html(gf, slide_no):
    xfrm = gf.find("p:xfrm", NS); off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    x, y, w = emu(off.get("x")), emu(off.get("y")), emu(ext.get("cx"))
    tbl = gf.find(".//a:tbl", NS)
    cols = [emu(c.get("w")) for c in tbl.findall("a:tblGrid/a:gridCol", NS)]
    rows = []
    for tr in tbl.findall("a:tr", NS):
        cells = []
        for tc in tr.findall("a:tc", NS):
            pr = tc.find("a:tcPr", NS)
            m = [emu(pr.get(k)) if pr is not None and pr.get(k) else 6 for k in ("marT", "marR", "marB", "marL")]
            fill = color(pr.find("a:solidFill", NS)) if pr is not None else None
            bd = ""
            for side, tag in (("left", "lnL"), ("right", "lnR"), ("top", "lnT"), ("bottom", "lnB")):
                l = pr.find(f"a:{tag}", NS) if pr is not None else None
                if l is not None and l.find("a:solidFill", NS) is not None:
                    bd += f"border-{side}:{max(0.75, emu(l.get('w') or 6350)):.2f}px solid {color(l.find('a:solidFill', NS))};"
            va = {"t": "top", "ctr": "middle", "b": "bottom"}.get(pr.get("anchor") if pr is not None else "t", "top")
            defaults = {"size": 18, "color": "#" + SCHEME["tx1"], "bold": False, "font": MINOR}
            inner = "".join(paragraphs_html(tc.find("a:txBody", NS), defaults, slide_no))
            cells.append(f'<td style="padding:{m[0]:.2f}px {m[1]:.2f}px {m[2]:.2f}px {m[3]:.2f}px;{bd}{"background:" + fill + ";" if fill else ""}vertical-align:{va};overflow:hidden">{inner}</td>')
        rows.append(f'<tr style="height:{emu(tr.get("h")):.2f}px">{"".join(cells)}</tr>')
    cg = "".join(f'<col style="width:{c:.2f}px">' for c in cols)
    return (f'<table style="position:absolute;left:{x:.2f}px;top:{y:.2f}px;width:{w:.2f}px;border-collapse:collapse;table-layout:fixed">'
            f'<colgroup>{cg}</colgroup>{"".join(rows)}</table>')

def picture_html(pic, part_rels, tmp):
    xfrm = pic.find("p:spPr/a:xfrm", NS); off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    rid = pic.find("p:blipFill/a:blip", NS).get(f"{{{NS['r']}}}embed")
    target = part_rels[rid][1]; data = z.read(target)
    f = tmp / pathlib.PurePosixPath(target).name; f.write_bytes(data)
    return (f'<img src="{f.as_uri()}" style="position:absolute;left:{emu(off.get("x")):.2f}px;top:{emu(off.get("y")):.2f}px;'
            f'width:{emu(ext.get("cx")):.2f}px;height:{emu(ext.get("cy")):.2f}px;object-fit:fill">')

def tree_html(tree, part, slide_no, defaults_by_ph, tmp, skip_ph=False, skip_sldnum=False):
    divs, lines = [], []
    prels = rels(part)
    for el in tree:
        tag = etree.QName(el).localname
        if tag == "sp":
            ph = el.find("p:nvSpPr/p:nvPr/p:ph", NS)
            if ph is not None and skip_ph and ph.get("type") != "sldNum": continue
            if ph is not None and ph.get("type") == "sldNum" and skip_sldnum: continue
            d, l = shape_html(el, slide_no, defaults_by_ph)
            if d: divs.append(d)
            if l: lines.append(l)
        elif tag == "pic": divs.append(picture_html(el, prels, tmp))
        elif tag == "graphicFrame" and el.find(".//a:tbl", NS) is not None: divs.append(table_html(el, slide_no))
    return divs, lines

def layout_defaults(layout):
    """placeholder type -> {xfrm, defaults} from a layout"""
    out = {}
    for sp in layout.findall(".//p:cSld/p:spTree/p:sp", NS):
        ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is None: continue
        d = {"size": 18, "color": "#" + SCHEME["tx1"], "bold": False, "font": MAJOR if ph.get("type") == "title" else MINOR}
        dr = sp.find("p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr", NS)
        if dr is not None: _, d = run_css(dr, d)
        out[ph.get("type")] = {"xfrm": sp.find("p:spPr/a:xfrm", NS), "defaults": d}
    return out

order = [rels("ppt/presentation.xml")[s.get(f"{{{NS['r']}}}id")][1] for s in pres.findall("p:sldIdLst/p:sldId", NS)]
tmp = pathlib.Path(tempfile.mkdtemp(prefix="pptx2pdf-"))
pages = []
for i, part in enumerate(order, 1):
    if only and i not in only: continue
    slide = X(part)
    lay_part = next(v for t, v in rels(part).values() if t == "slideLayout"); layout = X(lay_part)
    bg = color(slide.find("p:cSld/p:bg/p:bgPr/a:solidFill", NS)) or color(layout.find("p:cSld/p:bg/p:bgPr/a:solidFill", NS)) or "#" + SCHEME["lt1"]
    ldef = layout_defaults(layout)
    sl_tree = slide.find("p:cSld/p:spTree", NS)
    has_num = any(f.get("type") == "slidenum" for f in sl_tree.iter(f"{{{NS['a']}}}fld"))
    ld, ll = tree_html(layout.find("p:cSld/p:spTree", NS), lay_part, i, ldef, tmp, skip_ph=True, skip_sldnum=has_num)
    sd, sl = tree_html(sl_tree, part, i, ldef, tmp)
    allines = ll + sl
    markers = "".join(f'<marker id="m{c[1:]}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="{c}"/></marker>' for c in sorted({c for c, _ in allines}))
    svg = f'<svg style="position:absolute;left:0;top:0" width="{SW:.0f}" height="{SH:.0f}"><defs>{markers}</defs>{"".join(l for _, l in allines)}</svg>'
    # layout shapes, slide shapes, then connectors on top (as in PowerPoint's z-order for these decks: lines are added after boxes)
    pages.append(f'<section style="background:{bg}">{"".join(ld)}{"".join(sd)}{svg}</section>')

css = f"""@page{{size:{SW / 96:.3f}in {SH / 96:.3f}in;margin:0}}html,body{{margin:0;padding:0}}
section{{width:{SW:.2f}px;height:{SH:.2f}px;position:relative;overflow:hidden;page-break-after:always;break-after:page}}
section:last-child{{page-break-after:auto}}p{{white-space:pre-wrap;overflow-wrap:break-word}}td p{{white-space:normal}}"""
fonts = '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=block">'
page = tmp / "deck.html"
page.write_text(f"<!doctype html><html><head><meta charset='utf-8'>{fonts}<style>{css}</style></head><body>{''.join(pages)}</body></html>")
subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=15000",
                f"--print-to-pdf={out}", page.as_uri()], check=True, capture_output=True)
print("wrote", out, f"({len(pages)} slides, {SW / 96:.2f} x {SH / 96:.2f} in)")
if png_dir:
    import pymupdf
    png_dir.mkdir(parents=True, exist_ok=True)
    for n, pg in enumerate(pymupdf.open(out), 1):
        pg.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5)).save(png_dir / f"slide-{n:02d}.png")
    print("pngs in", png_dir)
