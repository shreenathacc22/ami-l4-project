"""Render SVG files to PNG with headless Chrome at 2x. Usage: python svg2png.py out_dir file.svg [file.svg ...]"""
import pathlib, subprocess, sys
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
out = pathlib.Path(sys.argv[1]).resolve(); out.mkdir(parents=True, exist_ok=True)
for f in sys.argv[2:]:
    svg = pathlib.Path(f).resolve()
    html = svg.with_suffix(".html")
    html.write_text(f"<!doctype html><html><head><meta charset='utf-8'><style>html,body{{margin:0;padding:0;background:#fff}}svg{{display:block}}</style></head><body>{svg.read_text()}</body></html>")
    png = out / (svg.stem + ".png")
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
                    "--window-size=1600,680", f"--screenshot={png}", html.as_uri()], check=True, capture_output=True)
    html.unlink()
    print("wrote", png)
