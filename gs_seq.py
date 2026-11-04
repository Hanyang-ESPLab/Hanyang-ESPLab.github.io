import re, sys
from bs4 import BeautifulSoup
URL = re.compile(r"https://(?:sites\.google\.com/sitesv-images-rt|lh\d\.googleusercontent\.com)/[^\"')\s&]+")
def seq(path):
    s = BeautifulSoup(open(path, encoding="utf-8").read(), "html.parser")
    for t in s(["script", "style", "nav", "footer", "meta", "link"]): t.decompose()
    for br in s.find_all("br"): br.replace_with("\n")
    out, seen_banner = [], False
    for el in s.find_all(True):
        srcs = []
        if el.name == "img":
            srcs = [el.get("src", "")]
        st = el.get("style", "") or ""
        if "background-image" in st:
            srcs += URL.findall(st)
        for u in srcs:
            if URL.match(u):
                if not seen_banner and el.find_parent("header"):
                    continue
                out.append(("img", u))
        if el.name in ("p", "h1", "h2", "h3", "li") and not el.find(["p", "h1", "h2", "h3", "li"]) and not el.find_parent("header"):
            t = "\n".join(x for x in (re.sub(r"[ \t\xa0]+", " ", l).strip() for l in el.get_text("").split("\n")) if x)
            if t: out.append(("txt", t))
    # drop consecutive duplicate images
    res = []
    for x in out:
        if not res or res[-1] != x: res.append(x)
    return res
if __name__ == "__main__":
    for k, v in seq(sys.argv[1]):
        print(k, v[-30:] if k == "img" else v[:110])
