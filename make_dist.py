"""Turn index.html (a body fragment) into a deployable static site in dist/.

index.html is written for a viewer that supplies the <html>/<head>/<body> skeleton. This wraps it into a
complete document (head metadata + the base reset that skeleton provided) and copies the images.
Standard library only:  python3 make_dist.py
"""
import json, re, shutil
from pathlib import Path

HERE = Path(__file__).resolve().parent

TITLE = "ESP Lab | 임베디드 보안 및 프라이버시 연구실"
DESC = "한양대학교 ERICA 임베디드 보안 및 프라이버시 연구실(ESP Lab). 암호공학, 양자보안, 지능형 모빌리티 보안을 연구합니다."
# What the viewer's skeleton used to provide; the page CSS assumes it ([hidden] drives the hash routing).
RESET = ("<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}"
         "body{margin:0;font-size:14px;font-family:system-ui,sans-serif}img{max-width:100%}"
         "[hidden]{display:none!important}</style>")

# The fragment's fixed preamble, in order. Each must match exactly once, at the top of what is left.
PREAMBLE = [
    ("<title> line", re.compile(r"\A<title>[^<]*</title>\n")),
    ("claude:viewport-fit-sides <meta> line", re.compile(r'\A<meta name="claude:viewport-fit-sides" content="cover">\n')),
    ("<link> block", re.compile(r"\A(?:<link [^>]*>\n)+")),
    ("page <style> block", re.compile(r"\A<style>.*?</style>\n", re.S)),
]


def split_fragment(frag):
    """Return (link tags, page <style> block, body markup); raise if the fragment deviates from the known layout."""
    parts, rest = {}, frag
    for name, rx in PREAMBLE:
        m = rx.match(rest)
        if not m:
            raise ValueError(f"fragment: {name} not found where expected (near: {rest[:80]!r})")
        parts[name], rest = m.group(0), rest[m.end():]
    body = rest.strip("\n")
    # Head-only tags must not survive into the body; the script must be the single trailing block.
    for tag in ("<title", "<meta", "<link", "<style"):
        if tag in body:
            raise ValueError(f"fragment: unexpected {tag} in the body markup")
    if body.count("<script") != 1 or not body.endswith("</script>"):
        raise ValueError("fragment: expected exactly one <script> block at the end of the body")
    return parts["<link> block"], parts["page <style> block"], body


def wrap(links, style, body):
    head = "\n".join([
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">',
        f"<title>{TITLE}</title>",
        f'<meta name="description" content="{DESC}">',
        f'<meta property="og:title" content="{TITLE}">',
        f'<meta property="og:description" content="{DESC}">',
        '<meta name="robots" content="index, follow">',
        links.rstrip("\n"),
        RESET,
        style.rstrip("\n"),
    ])
    return f'<!doctype html>\n<html lang="ko">\n<head>\n{head}\n</head>\n<body>\n{body}\n</body>\n</html>\n'


def build_dist(fragment_path=HERE / "index.html", out_dir=HERE / "dist"):
    fragment_path, out_dir = Path(fragment_path), Path(out_dir).resolve()
    if HERE not in out_dir.parents:  # out_dir is wiped, so never point it outside the project
        raise ValueError(f"out_dir must be inside {HERE}: {out_dir}")
    src = fragment_path.parent
    links, style, body = split_fragment(fragment_path.read_text(encoding="utf-8"))
    manifest = src / "files.json"
    if manifest.exists():
        images = sorted(json.loads(manifest.read_text(encoding="utf-8")))
        bad = [k for k in images if not k.startswith("images/") or "/" in k[len("images/"):]]
        if bad:
            raise ValueError(f"files.json: keys must be images/<name>: {bad[:3]}")
    else:
        images = sorted(f"images/{p.name}" for p in (src / "images").iterdir() if p.is_file())
    if out_dir.exists():
        shutil.rmtree(out_dir)
    (out_dir / "images").mkdir(parents=True)
    (out_dir / "index.html").write_text(wrap(links, style, body), encoding="utf-8")
    (out_dir / ".nojekyll").write_text("")
    for k in images:
        shutil.copyfile(src / k, out_dir / k)
    files = [p for p in out_dir.rglob("*") if p.is_file()]
    print(f"dist: {len(files)} files, {sum(p.stat().st_size for p in files) / 1024:.0f} KB -> {out_dir}")


if __name__ == "__main__":
    build_dist()
