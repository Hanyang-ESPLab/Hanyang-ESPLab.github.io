"""Parse the cached Google Sites pages (_cache/gs/*.html) into structured data with image URLs."""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from gs_seq import seq  # noqa: E402

GS = Path(__file__).parent / "_cache" / "gs"
NAME = re.compile(r"^(.+?) \((.+?)\)\s*🏠?$")
DATE = re.compile(r"^(20\d\d)년\s*(\d{1,2})월(?:\s*(\d{1,2})일)?")
CATS = {"Workshop": "event", "Patent": "patent", "New research project": "project", "Top Conference": "paper",
        "International Cooperation": "event", "Award": "award", "News": "event", "Exhibition": "event"}

def tokens(name):
    toks = seq(GS / f"{name}.html")
    # the first image on every page is the shared header banner
    if toks and toks[0][0] == "img":
        toks = toks[1:]
    skip = {"Ph.D. Course", "M.S. Course", "Alumni", "Undergraduate researcher", "Undergraduate researchers", "Undergraduate Researcher"}
    return [t for t in toks if not (t[0] == "txt" and (t[1] in skip or re.fullmatch(r"20\d\d", t[1])))]

def interests(line):
    return [x.strip() for x in line.replace("Research Interests", "").split("- ") if x.strip()] if line.startswith("Research Interests") else []

def members(page, group):
    out, cur, section, pending_img = [], None, "", None
    for kind, v in tokens(page):
        if kind == "img":
            pending_img = v
            continue
        if v in ("Full-Time Students", "Part-Time Students", "Ph.D. Alumni", "Master's Alumni"):
            section = v
            continue
        m = NAME.match(v)
        if m and re.search(r"[가-힣]", m.group(2)):
            cur = {"en": m.group(1).strip(), "ko": m.group(2).replace(" ", ""), "img": pending_img, "section": section,
                   "group": group, "role": "", "org": [], "period": [], "email": [], "interests": []}
            pending_img = None
            out.append(cur)
            continue
        if not cur:
            continue
        if v.startswith("Research Interests"):
            cur["interests"] = interests(v)
        elif "@" in v:
            cur["email"].append(v)
        elif v == "Affiliation:":
            pass
        elif re.match(r"^(M\.S\.|Ph\.D\.|M\.S\./Ph\.D\.) \d{4}", v):
            cur["period"] = re.findall(r"(M\.S\./Ph\.D\.|M\.S\.|Ph\.D\.) (\d{4}\.\d{2}-\d{4}\.\d{2})", v)
        elif re.search(r"Student", v):
            cur["role"] = v
        else:
            cur["org"] += [x.strip() + (")" if "(" in x else "") for x in v.split(")") if x.strip()]
    # Google Sites uses one shared silhouette for members without a photo
    from collections import Counter
    c = Counter(p["img"] for p in out if p["img"])
    for p in out:
        p["ph"] = bool(p["img"] and c[p["img"]] > 1)
    return out

def professor():
    t = [v for k, v in tokens("professor") if k == "txt"]
    img = next((v for k, v in tokens("professor") if k == "img"), None)
    def block(start, end):
        i = t.index(start) + 1
        j = t.index(end) if end else len(t)
        return t[i:j]
    def split_awd(s):
        m = re.match(r"^(.*?)\((.*)\)$", s)
        return {"en": m.group(1).strip(), "ko": m.group(2).strip()} if m else {"en": s, "ko": s}
    contact = next(x for x in t if x.startswith("Tel:"))
    fax = re.search(r"Fax: ([+\d-]+)", contact)
    return {"img": img, "areas": block("Research Area", "Research Experience"), "career": block("Research Experience", "Education"),
            "education": [x.replace("Ehwa", "Ewha") for x in block("Education", "Awards")],
            "awards": [split_awd(x) for x in block("Awards", None)], "fax": fax.group(1) if fax else ""}

def about():
    out, cur = [], None
    for k, v in tokens("about-lab"):
        if k == "txt" and v.startswith("- ") and v.endswith(" -"):
            cur = {"title": v.strip("- ").strip(), "img": None, "body": ""}
            out.append(cur)
        elif cur and k == "img" and not cur["img"]:
            cur["img"] = v
        elif cur and k == "txt":
            cur["body"] += v.replace("비트모인", "비트코인")
    return out

def norm_date(line):
    m = DATE.match(line)
    if not m:
        return ""
    y, mo, d = m.group(1), int(m.group(2)), m.group(3)
    return f"{y}.{mo:02d}" + (f".{int(d):02d}" if d else "")

def where(line):
    rest = re.sub(r"^(20\d\d년\s*\d{1,2}월(\s*\d{1,2}일)?(\s*\([^)]*\))?\s*(-\s*)?)+", "", line).strip()
    return rest.lstrip(", ").strip()

def news(year):
    toks = tokens(f"news_{year}")
    items, cur, cat, pending = [], None, "News", {}
    txts = [i for i, t in enumerate(toks) if t[0] == "txt"]
    def nxt(i, k=1):
        later = [j for j in txts if j > i]
        return toks[later[k - 1]][1] if len(later) >= k else ""
    for i, (kind, v) in enumerate(toks):
        if kind == "img":
            if cur:
                cur["imgs"].append(v)
            continue
        if v in CATS:
            cat, cur = v, None
            continue
        if re.match(r"^※ 20\d\d 신규 과제 선정$", v):
            pending = {"cat": "New research project"}
            cur = None
            continue
        if DATE.match(v) and pending and not cur:
            pending["date_line"] = v
            continue
        bodyish = DATE.match(v) or v.startswith(("연구목표", "연구내용", "-"))
        if not bodyish and cur and not cur["date"] and not cur["body"] and DATE.match(nxt(i)):
            # a subtitle such as "서승현 교수님" between an award title and its date
            cur["sub"] = v
            continue
        starts = not bodyish and (v.startswith("※") or DATE.match(nxt(i)) or nxt(i).startswith(("연구목표", "행사명"))
                                  or (cat in ("Award", "News") and DATE.match(nxt(i, 2)) and len(v) < 60))
        if starts:
            c = pending.get("cat", cat)
            dl = pending.get("date_line", "")
            cur = {"cat": c, "type": CATS.get(c, "event"), "title": v.lstrip("※ ").strip(), "sub": "", "date": norm_date(dl) or "",
                   "where": where(dl), "body": [], "imgs": [], "year": year}
            pending = {}
            items.append(cur)
            continue
        if not cur:
            continue
        if DATE.match(v) and not cur["date"]:
            cur["date"] = norm_date(v)
            cur["where"] = where(v)
        else:
            cur["body"].append(v)
    full = re.compile(r"(20\d\d)년\s*(\d{1,2})월\s*(\d{1,2})일")
    for it in items:
        if it["date"] == "":
            m = next((full.search(b) for b in it["body"] if full.search(b)), None)
            it["date"] = f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}" if m else str(year)
        it["sub"] = it["sub"].replace("빅가을", "박가을")
    return items

PAGES = ["members/ph-d-course", "members/m-s-course", "members/undergraduate-researchers", "members/alumni",
         "professor", "about-lab", "news/2026", "news/2025", "news/2024", "news/2023", "invited-talk"]

def refresh(only=None):
    """Image URLs on Google Sites are short-lived, so fetch the pages right before downloading."""
    import subprocess
    GS.mkdir(parents=True, exist_ok=True)
    for pg in (only or PAGES):
        subprocess.run(["curl", "-sfL", "--retry", "3", f"https://sites.google.com/view/esplab/{pg}", "-o", str(GS / (pg.replace("/", "_") + ".html"))], check=True)

def parse():
    mem = {"phd": members("members_ph-d-course", "phd"), "ms": members("members_m-s-course", "ms"),
           "ug": members("members_undergraduate-researchers", "ug"), "alumni": members("members_alumni", "alumni")}
    nws = [n for y in (2026, 2025, 2024, 2023) for n in news(y)]
    return {"members": mem, "professor": professor(), "about": about(), "news": nws}

if __name__ == "__main__":
    import json
    d = parse()
    for k, v in d["members"].items():
        for p in v:
            print(k, p["ko"], p["en"], "IMG" if p["img"] else "-", p["section"], p["role"], p["org"], p["period"], p["email"], p["interests"])
    print(json.dumps({k: v for k, v in d["professor"].items() if k != "img"}, ensure_ascii=False)[:600])
    for a in d["about"]:
        print("ABOUT", a["title"], bool(a["img"]), a["body"][:60])
    for n in d["news"]:
        print("NEWS", n["date"], n["cat"], "|", n["title"][:50], "|", n["sub"][:30], "|", n["where"][:30], "| imgs", len(n["imgs"]), "| body", len(n["body"]), (n["body"][0][:60] if n["body"] else ""))
