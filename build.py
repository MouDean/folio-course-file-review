"""Build dist/folio.html from the source files in src/ (no dependencies)."""
from pathlib import Path
src = Path(__file__).parent / "src"
top = (src / "top.html").read_text(encoding="utf-8")
app = (src / "app.js").read_text(encoding="utf-8").replace("/*@@EXPORT@@*/", (src / "export.js").read_text(encoding="utf-8"))
cls = (src / "classify.js").read_text(encoding="utf-8")
out = top + "\n<script>\n" + cls + "\n</script>\n<script>\n" + app + "\n</script>\n"
dist = Path(__file__).parent / "dist"; dist.mkdir(exist_ok=True)
(dist / "folio.html").write_text(out, encoding="utf-8")
print("Built dist/folio.html", len(out), "bytes")
