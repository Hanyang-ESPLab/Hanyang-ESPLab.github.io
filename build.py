"""Build index.html from the frozen JSON data in data/ (plus the Excel achievement lists when present).

  python build.py             read the Excel files if present -> data/publications.json; render data/*.json -> index.html, files.json;
                              then make_dist.py turns index.html into the deployable site in dist/ (skip with --no-dist)
  python build.py --sync-gs   also fetch the Google Sites pages (cached in _cache/gs/), download their photos into images/
                              and rewrite data/{members,news,professor,about,talks}.json
  python build.py --no-excel  skip the Excel files even if they are present (data/publications.json stays as it is)

Data flow:
  Excel (LAB/...)               -> data/publications.json  (SCI sheet: IF, JCR 상위%, JCR 연도 columns are shown as badges)
  Google Sites (--sync-gs)      -> data/members|news|professor|about|talks.json + images/
  data/*.json + index.tpl.html  -> index.html + files.json (manifest of the images index.html uses)
  index.html + images/          -> dist/ (full HTML document, images, .nojekyll); pushing main publishes dist/
                                   to GitHub Pages through .github/workflows/pages.yml
data/projects.json and data/theme_overrides.json (manual research-theme assignment per paper) are maintained by hand. Run with a Python that has openpyxl and beautifulsoup4 (requirements.txt).
"""
import argparse, base64, hashlib, json, re, subprocess, sys
from pathlib import Path
import openpyxl
from bs4 import BeautifulSoup
sys.path.insert(0, str(Path(__file__).parent))
import gs_parse
import make_dist

HERE = Path(__file__).parent
LAB = Path("/Users/yn4103/ESP Lab Dropbox/ESPLAB/연구실 실적")
EXCEL = [LAB / "논문실적/ESP 논문실적 (최신).xlsx", LAB / "특허_소프트웨어_실적/ESP 특허실적 (최신).xlsx"]
DATA = HERE / "data"

THEMES = {
    "quantum": r"(?<!post-)(?<!post )quantum(?![- ](resistant|safe))|grover|양자(?!\s?내성)|t-depth|qubit",
    "pqc": r"post-quantum|post quantum|quantum[- ](resistant|safe)|\bpqc|양자\s?내성|lattice|\blwe|rainbow|\buov\b|minrank|multivariate|다변수|\bmq\b|hybrid|하이브리드|\btls\b|certificate|인증서|\bpki\b|\bkem\b|number theoretic|\bntt\b|himq|paloma",
    "embedded": r"\biot\b|드론|drone|\buavs?\b|unmanned|무인|robot|로봇|\bros|embedded|임베디드|fpga|sensor|센서|8비트|mcu|microcontroller|smart home|스마트 ?홈|scada|android|mobile|모바일|device|디바이스|wearable|metering|acoustic|keypad|키패드|마우스|manufactur|web-of-things|wireless sensor",
    "mobility": r"blockchain|블록체인|vehicle|차량|v2v|전기차|자동차|mobility|모빌리티|\bdid\b|분산\s?(신원|id)|smart contract|스마트 ?컨트랙트|킥보드|톨 티켓|crowdsens|vpp",
}

# Only these conferences carry the "Top-tier" badge.
TOP_TIER = [r"\bCHES\b|\bTCHES\b", r"\bWWW\b|Web Conference"]
# Papers the spreadsheet files under the SCI sheet that are really conference papers, with the label to show.
AS_CONFERENCE = [(r"\bTCHES\b", "CHES {y} (IACR TCHES)"), (r"\bWWW\b|Web Conference", "WWW {y} (The ACM Web Conference)")]

def tag(title):
    t = title.lower()
    return [k for k, rx in THEMES.items() if re.search(rx, t)]

def year_of(v):
    s = str(v).strip()
    if re.match(r"^\d{4}", s):
        return int(s[:4]), s[:7].replace("-", ".")
    m = re.match(r"^(\d{2})(?:\.(\d{2}))?", s)
    if not m:
        return None, ""
    y = 2000 + int(m.group(1))
    mo = m.group(2) if m.group(2) and m.group(2) != "00" else ""
    return y, f"{y}.{mo}" if mo else str(y)

def header(ws):
    """Column index of each header cell (same offset as rows()), e.g. header(ws)["JCR 상위%"]."""
    first = next(r for r in ws.iter_rows(values_only=True) if any(c is not None for c in r))
    off = next(i for i, c in enumerate(first) if c is not None)
    return {clean(c): i - off for i, c in enumerate(first) if c is not None}

def rows(ws):
    out = [r for r in ws.iter_rows(values_only=True) if any(c is not None for c in r)]
    # sheets start in a column after A; align every row to the header's first cell
    off = next(i for i, c in enumerate(out[0]) if c is not None)
    return [r[off:] for r in out[1:]]

def clean(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""

def load_pubs():
    wb = openpyxl.load_workbook(EXCEL[0], read_only=True, data_only=True)
    kinds = {"SCI": "sci", "국내학술지": "kj", "국제학술대회": "ic", "국내학술대회": "kc"}
    pubs = []
    for ws in wb.worksheets:
        kind = kinds.get(ws.title)
        if not kind:
            continue
        hdr = header(ws)
        for r in rows(ws):
            r = list(r) + [None] * 10
            title, authors, venue = clean(r[2]), clean(r[3]), clean(r[4])
            if not title:
                continue
            y, d = year_of(r[1])
            # the spreadsheet has a few rows with Seong-Min Cho typed as "Seo"
            authors = authors.replace("Seong-Min Seo", "Seong-Min Cho")
            impact = clean(r[hdr["IF"]]) if kind == "sci" else ""
            # JCR JIF percentile of the publication-year JCR, filled in by hand in the SCI sheet; "" until it is
            jcr = clean(r[hdr["JCR 상위%"]]) if kind == "sci" and "JCR 상위%" in hdr else ""
            jcr_y = clean(r[hdr["JCR 연도"]]) if kind == "sci" and "JCR 연도" in hdr else ""
            pubs.append({"kind": kind, "y": y, "d": d, "t": title, "a": authors, "v": venue, "k": tag(title),
                         "if": impact if re.fullmatch(r"\d+(\.\d+)?", impact) else "",
                         "jcr": float(jcr) if re.fullmatch(r"\d+(\.\d+)?", jcr) else "",
                         "jcr_y": int(float(jcr_y)) if re.fullmatch(r"\d{4}(\.0)?", jcr_y) else (y or "")})
    # the same journal sometimes has its IF filled in on only one row
    known_if = {p["v"].lower(): p["if"] for p in pubs if p["if"]}
    for p in pubs:
        for rx, label in AS_CONFERENCE:
            if p["kind"] == "sci" and re.search(rx, p["v"], re.I):
                p["kind"], p["v"], p["if"], p["jcr"] = "ic", label.format(y=p["y"]), "", ""
        if p["kind"] == "sci" and not p["if"]:
            p["if"] = known_if.get(p["v"].lower(), "")
        p["top"] = p["kind"] == "ic" and any(re.search(rx, p["v"], re.I) for rx in TOP_TIER)
    pubs.sort(key=lambda p: (p["y"] or 0, p["d"]), reverse=True)
    return pubs

def load_ip():
    wb = openpyxl.load_workbook(EXCEL[1], read_only=True, data_only=True)
    items = []
    for ws in wb.worksheets:
        for r in rows(ws):
            r = [clean(c) for c in list(r) + [None] * 10]
            if not r[1]:
                continue
            if ws.title == "특허":
                reg = r[7] if re.match(r"^10-\d", r[7]) else ""
                y, _ = year_of(r[6] if reg else r[4])
                no = f"등록 {reg}" if reg else f"출원 {r[5]}"
                items.append({"kind": "pat", "y": y, "d": str(y), "t": r[1], "a": r[2], "v": no, "k": tag(r[1])})
            else:
                y, _ = year_of(r[5])
                items.append({"kind": "sw", "y": y, "d": str(y), "t": r[1], "a": r[2], "v": "SW 등록 " + r[5][:7].replace("-", "."), "k": tag(r[1])})
    items.append({"kind": "pat", "y": 2026, "d": "2026.04", "t": "Efficient Quantum Modular Multiplier and Quantum Modular Multiplication Method", "a": "Seong-Min Cho, Seung-Hyun Seo, Aeyoung Kim", "v": "US Patent US12608635B2", "k": ["quantum"]})
    items.sort(key=lambda p: p["y"] or 0, reverse=True)
    return items

def load_talks():
    s = BeautifulSoup((gs_parse.GS / "invited-talk.html").read_text(encoding="utf-8"), "html.parser")
    for t in s(["script", "style"]):
        t.decompose()
    lines = [l.strip() for l in s.get_text("\n").split("\n") if l.strip()]
    start = lines.index("Invited Talks & Seminars") + 1
    talks, cur = [], None
    for l in lines[start:]:
        if l.startswith("©"):
            break
        m = re.match(r"^(20\d\d\.[\d–\-.]+)\s+—\s+(.+)$", l)
        if m:
            cur = {"d": m.group(1), "en_ev": m.group(2), "rest": []}
            talks.append(cur)
        elif cur and not re.fullmatch(r"20\d\d", l):
            cur["rest"].append(l)
    out = []
    for t in talks:
        rest = [x for x in t["rest"] if x != "(예정)"]
        out.append({
            "d": t["d"], "y": int(t["d"][:4]), "planned": "(예정)" in t["rest"],
            "ev_en": t["en_ev"], "ev_ko": rest[0] if rest else t["en_ev"],
            "t_en": rest[1] if len(rest) > 1 else "", "t_ko": rest[2] if len(rest) > 2 else (rest[1] if len(rest) > 1 else ""),
        })
    return out

IMG = HERE / "images"
RAW = HERE / "_cache"

def local_img(url):
    """Download a Google Sites image, shrink it, and return its published path (named by content)."""
    if not url:
        return None
    IMG.mkdir(exist_ok=True); RAW.mkdir(exist_ok=True)
    raw = RAW / "dl.bin"
    r = subprocess.run(["curl", "-sfL", "--http1.1", "--retry", "4", "--retry-delay", "3", "--max-time", "180", url, "-o", str(raw)])
    if r.returncode != 0:
        print("  image skipped (curl", r.returncode, ")", url[-24:], file=sys.stderr)
        return None
    name = hashlib.sha1(raw.read_bytes()).hexdigest()[:12] + ".jpg"
    out = IMG / name
    if not out.exists():
        subprocess.run(["sips", "-s", "format", "jpeg", "-s", "formatOptions", "78", "-Z", "1000", str(raw), "--out", str(out)],
                       check=True, capture_output=True)
    return "images/" + name

def used_images(data):
    """Names of every images/<hash>.jpg the data refers to (what index.html needs next to it)."""
    out = set()
    def walk(x):
        if isinstance(x, dict):
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
        elif isinstance(x, str) and re.fullmatch(r"images/[0-9a-f]{12}\.jpg", x):
            out.add(x[len("images/"):])
    walk(data)
    return out

def save(name, obj):
    DATA.mkdir(exist_ok=True)
    (DATA / f"{name}.json").write_text(json.dumps(obj, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

def load(name, how):
    f = DATA / f"{name}.json"
    if not f.exists():
        sys.exit(f"build.py: data/{name}.json is missing; {how}")
    return json.loads(f.read_text(encoding="utf-8"))

# Awards that are not on the Google Sites professor page; merged in during --sync-gs.
EXTRA_AWARDS = [
    {"y": "2023.12", "ko": "정보처리학회 창립 30주년 ICT 분야 학문 및 산업발전 기여", "en": "Contribution to ICT Research and Industry, KIPS 30th Anniversary"},
    {"y": "2024.12", "ko": "한국인터넷진흥원 K-shield 올해의 우수 강사", "en": "K-shield Instructor of the Year, KISA"},
    {"y": "2025.07", "ko": "2025학년도 1학기 IC-PBL 및 IC-PBL+ 콘테스트 최우수상 (하이브리드 양자 암호)", "en": "Grand Prize, IC-PBL Contest (Hybrid Quantum Cryptography), Hanyang University ERICA"},
    {"y": "2025.12", "ko": "2025 최정예 정보보호 전문인력 양성 교육 우수강사상, 한국인터넷진흥원", "en": "Excellent Instructor Award, Elite Security Talent Training, KISA"},
]

def sync_gs():
    """Fetch the Google Sites pages and their photos, then freeze the parsed data under data/."""
    # fetch each page right before downloading its images; the image links expire within minutes
    def fresh(page, fn, *args):
        gs_parse.refresh([page])
        return fn(*args)
    members = {}
    for key, page, slug in [("phd", "members/ph-d-course", "members_ph-d-course"), ("ms", "members/m-s-course", "members_m-s-course"),
                            ("ug", "members/undergraduate-researchers", "members_undergraduate-researchers"), ("alumni", "members/alumni", "members_alumni")]:
        members[key] = fresh(page, gs_parse.members, slug, key)
        for m in members[key]:
            m["img"] = local_img(m["img"])
    # members without a photo get the same default icon Google Sites uses for undergraduates
    icon = next((m["img"] for g in members.values() for m in g if m["ph"] and m["img"]), None)
    for g in members.values():
        for m in g:
            if not m["img"]:
                m["img"], m["ph"] = icon, True
    save("members", members)
    news = []
    for y in (2026, 2025, 2024, 2023):
        items = fresh(f"news/{y}", gs_parse.news, y)
        for n in items:
            n["imgs"] = [x for x in (local_img(u) for u in n["imgs"]) if x]
        news += items
    save("news", sorted(news, key=lambda n: n["date"], reverse=True))
    prof = fresh("professor", gs_parse.professor)
    prof["img"] = local_img(prof["img"])
    awards = []
    for a in prof["awards"]:
        y = re.search(r"20\d\d", a["en"] + a["ko"])
        awards.append({"y": y.group(0) if y else "", "ko": a["ko"], "en": a["en"]})
    prof["awards"] = sorted(awards + EXTRA_AWARDS, key=lambda a: a["y"], reverse=True)
    save("professor", prof)
    about = fresh("about-lab", gs_parse.about)
    for a in about:
        a["img"] = local_img(a["img"])
    save("about", about)
    save("talks", fresh("invited-talk", load_talks))

def main():
    ap = argparse.ArgumentParser(description="Render data/*.json + index.tpl.html into index.html; see the module docstring.")
    ap.add_argument("--sync-gs", action="store_true", help="re-fetch the Google Sites pages and photos, rewrite data/{members,news,professor,about,talks}.json")
    ap.add_argument("--no-excel", action="store_true", help="do not read the Excel files even if they are present")
    ap.add_argument("--no-dist", action="store_true", help="do not rebuild dist/ after index.html")
    args = ap.parse_args()
    if not args.no_excel and all(f.exists() for f in EXCEL):
        pubs = load_pubs() + load_ip()
        save("publications", pubs)
        print("publications.json rewritten from Excel:", len(pubs), "entries", file=sys.stderr)
    if args.sync_gs:
        sync_gs()
    data = {
        "pubs": load("publications", "run build.py with the Excel files present (without --no-excel)"),
        "talks": load("talks", "run build.py --sync-gs"),
        "projects": load("projects", "it is maintained by hand ({\"ongoing\": [...], \"past\": [...]} of {ko, en, ag, p})"),
        "members": load("members", "run build.py --sync-gs"),
        "news": load("news", "run build.py --sync-gs"),
        "prof": load("professor", "run build.py --sync-gs"),
        "about": load("about", "run build.py --sync-gs"),
        "site": load("site", "it is maintained by hand ({\"hero\": {img, cap_ko, cap_en}})"),
    }
    overrides = {k.lower(): v for k, v in load("theme_overrides", "it is maintained by hand").items() if not k.startswith("_")}
    for p in data["pubs"]:
        hit = next((v for k, v in overrides.items() if p["t"].lower().startswith(k)), None)
        if hit is not None:
            p["k"] = hit
    imgs = sorted(used_images(data))
    (HERE / "files.json").write_text(json.dumps({f"images/{n}": str(IMG / n) for n in imgs}, indent=1))
    print("images", len(imgs), sum((IMG / n).stat().st_size for n in imgs) // 1024, "KB", file=sys.stderr)
    counts = {k: sum(1 for p in data["pubs"] if p["kind"] == k) for k in ["sci", "ic", "kj", "kc", "pat", "sw"]}
    themed = {k: sum(1 for p in data["pubs"] if k in p["k"]) for k in THEMES}
    untagged = sum(1 for p in data["pubs"] if not p["k"])
    print("counts", counts, "themes", themed, "untagged", untagged, "talks", len(data["talks"]), "news", len(data["news"]), file=sys.stderr)
    logo = "data:image/png;base64," + base64.b64encode((HERE / "logo.png").read_bytes()).decode()
    html = (HERE / "index.tpl.html").read_text(encoding="utf-8")
    html = html.replace("__LOGO__", logo).replace("__DATA__", json.dumps(data, ensure_ascii=False).replace("</", "<\\/"))
    (HERE / "index.html").write_text(html, encoding="utf-8")
    print(len(html), file=sys.stderr)
    if not args.no_dist:
        make_dist.build_dist()

if __name__ == "__main__":
    main()
