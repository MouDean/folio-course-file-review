"""Build design-system/tokens.css from tokens.json and check that src/top.html uses the same colours.

Usage:  python3 design-system/build_tokens.py          (writes tokens.css, then checks top.html)
        python3 design-system/build_tokens.py --check  (check only; exit 1 if anything differs)
No dependencies.
"""
import json, re, sys
from pathlib import Path

HERE = Path(__file__).parent
TOKENS = json.loads((HERE / "tokens.json").read_text(encoding="utf-8"))
TOP = HERE.parent / "src" / "top.html"


def colour_values(theme):
    first = TOKENS["color"]["themes"][0]["id"]
    out = {}
    for t in TOKENS["color"]["tokens"]:
        v = t["value"]
        out[t["name"]] = v if isinstance(v, str) else v.get(theme, v[first])
    return out


def build_css():
    themes = [t["id"] for t in TOKENS["color"]["themes"]]
    lines = ["/* Folio — generated from tokens.json by build_tokens.py. Do not edit by hand. */"]
    for i, th in enumerate(themes):
        sel = ':root, [data-theme="%s"]' % th if i == 0 else '[data-theme="%s"]' % th
        lines.append(sel + " {")
        for k, v in colour_values(th).items():
            lines.append("  --%s: %s;" % (k, v))
        for s in TOKENS.get("shadow", {}).get("tokens", []):
            v = s["value"] if isinstance(s["value"], str) else s["value"].get(th, s["value"][themes[0]])
            lines.append("  --%s: %s;" % (s["name"], v))
        lines.append("}")
    lines.append(":root {")
    for fam in ("spacing", "radius"):
        for t in TOKENS.get(fam, {}).get("tokens", []):
            lines.append("  --%s: %s;" % (t["name"], t["value"]))
    for k, v in TOKENS["type"]["families"].items():
        lines.append("  --font-%s: %s;" % (k, v))
    lines.append("}")
    for g in TOKENS["type"]["groups"]:
        for s in g["styles"]:
            fam = s.get("family", g["family"])
            decl = ["font-family: var(--font-%s)" % fam, "font-size: %s" % s["fontSize"], "line-height: %s" % s["lineHeight"], "font-weight: %s" % s["fontWeight"]]
            if s.get("letterSpacing"):
                decl.append("letter-spacing: %s" % s["letterSpacing"])
            lines.append(".%s { %s; }" % (s["name"], "; ".join(decl)))
    (HERE / "tokens.css").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("Wrote design-system/tokens.css")


def check_top():
    """Compare the colour variables in src/top.html (light :root block and dark block) with tokens.json."""
    html = TOP.read_text(encoding="utf-8")
    light_block = re.findall(r"\n:root\{([^}]*)\}", html)[-1]  # the last :root block is the active identity
    dark_block = re.findall(r':root\[data-theme="dark"\]\{([^}]*)\}', html)[-1]
    problems = 0
    for theme, block in (("light", light_block), ("dark", dark_block)):
        page = {k: v.strip().upper() for k, v in re.findall(r"--([a-z0-9-]+):\s*(#[0-9A-Fa-f]{3,8})", block)}
        for k, v in colour_values(theme).items():
            if k not in page:
                print("  %s: --%s is in tokens.json but not in top.html" % (theme, k)); problems += 1
            elif page[k] != v.upper():
                print("  %s: --%s is %s in top.html but %s in tokens.json" % (theme, k, page[k], v)); problems += 1
    print("top.html colours match tokens.json" if not problems else "%d difference(s) found" % problems)
    return problems


if __name__ == "__main__":
    if "--check" not in sys.argv:
        build_css()
    sys.exit(1 if check_top() else 0)
