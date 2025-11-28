from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

import os

from flask import Flask, flash, redirect, render_template, request, send_file

from wordtotex import ConverterConfig, DocxToLatexConverter


app = Flask(__name__)
app.secret_key = "wordtotex-secret"  # for flash messages; replace in production


def convert_upload(temp_path: Path, include_preamble: bool, table_border: bool) -> Path:
    config = ConverterConfig(include_preamble=include_preamble, table_border=table_border)
    converter = DocxToLatexConverter(config=config)
    latex_content = converter.convert(temp_path)

    output = temp_path.with_suffix(".tex")
    output.write_text(latex_content, encoding="utf-8")
    return output


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    file = request.files.get("docx_file")
    if not file or file.filename == "":
        flash("请选择一个 .docx 文件后再试。")
        return redirect(request.url)

    include_preamble = request.form.get("include_preamble") == "on"
    table_border = request.form.get("table_border") == "on"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp:
        file.save(temp.name)
        temp_path = Path(temp.name)

    try:
        tex_path = convert_upload(temp_path, include_preamble, table_border)
        return send_file(
            tex_path,
            as_attachment=True,
            download_name=f"{Path(file.filename).stem}.tex",
            mimetype="text/x-tex",
        )
    except Exception as exc:  # noqa: BLE001
        flash(f"转换失败: {exc}")
        return redirect(request.url)
    finally:
        # Clean up the temporary files if they exist
        for path in (temp_path, temp_path.with_suffix(".tex")):
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass


def main() -> None:
    host = os.getenv("WORDTOTEX_HOST", "127.0.0.1")
    port_str = os.getenv("WORDTOTEX_PORT", "8000")
    try:
        port = int(port_str)
    except ValueError:
        port = 8000
    app.run(debug=True, host=host, port=port)


if __name__ == "__main__":
    main()
