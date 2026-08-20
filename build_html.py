#!/usr/bin/env python3
"""
Genera pliegos_judiciales.html embebiendo pliegos_data.json en template.html.

Uso:
    python3 data_builder.py     # regenera pliegos_data.json a partir del Excel
    python3 build_html.py       # regenera el HTML final con esos datos

Ver README.md para instrucciones completas.
"""
import argparse
import base64
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", default=HERE / "pliegos_data.json")
    parser.add_argument("--template", default=HERE / "template.html")
    parser.add_argument("--fonts-dir", default=HERE / "fonts")
    parser.add_argument("--out", default=HERE / "pliegos_judiciales.html")
    args = parser.parse_args()

    data_path, template_path, fonts_dir, out_path = (
        Path(args.data), Path(args.template), Path(args.fonts_dir), Path(args.out)
    )

    data = json.loads(data_path.read_text(encoding="utf-8"))
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    data_json_safe = data_json.replace("</", "<\\/")  # evita cerrar el <script> prematuramente

    def font_b64(filename):
        return base64.b64encode((fonts_dir / filename).read_bytes()).decode("ascii")

    html = template_path.read_text(encoding="utf-8")
    html = html.replace("__FONT400__", font_b64("Montserrat-Regular.woff2"))
    html = html.replace("__FONT600__", font_b64("Montserrat-SemiBold.woff2"))
    html = html.replace("__FONT700__", font_b64("Montserrat-Bold.woff2"))
    html = html.replace("__DATA_JSON__", data_json_safe)

    out_path.write_text(html, encoding="utf-8")
    print(f"OK: {out_path} generado ({len(html):,} bytes, {data['total']} pliegos)")


if __name__ == "__main__":
    main()
