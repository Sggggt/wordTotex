from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path
from typing import Optional
from zipfile import ZipFile

from flask import Flask, flash, redirect, render_template, request, send_file, after_this_request

from wordtotex import ConversionResult, ConverterConfig, DocxToLatexConverter


app = Flask(__name__)
app.secret_key = "wordtotex-secret"  # for flash messages; replace in production


def convert_upload(
    docx_path: Path, include_preamble: bool, table_border: bool, output_dir: Path
) -> tuple[Path, ConversionResult]:
    config = ConverterConfig(include_preamble=include_preamble, table_border=table_border)
    converter = DocxToLatexConverter(config=config)
    result = converter.convert(docx_path, output_dir=output_dir)

    output = output_dir / f"{docx_path.stem}.tex"
    output.write_text(result.latex, encoding="utf-8")
    return output, result


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    file = request.files.get("docx_file")
    if not file or file.filename == "":
        flash("Please choose a .docx file before converting.")
        return redirect(request.url)

    include_preamble = request.form.get("include_preamble") == "on"
    table_border = request.form.get("table_border") == "on"
    output_stem = Path(file.filename).stem or "output"

    temp_dir = Path(tempfile.mkdtemp())

    @after_this_request
    def _cleanup(response):
        shutil.rmtree(temp_dir, ignore_errors=True)
        return response

    upload_path = temp_dir / f"{output_stem}.docx"
    file.save(upload_path)

    try:
        tex_path, result = convert_upload(
            upload_path, include_preamble, table_border, output_dir=temp_dir
        )
        if result.image_paths:
            archive_path = temp_dir / f"{output_stem}_latex.zip"
            with ZipFile(archive_path, "w") as archive:
                archive.write(tex_path, arcname=tex_path.name)
                for image_path in result.image_paths:
                    archive.write(image_path, arcname=image_path.relative_to(temp_dir))

            return send_file(
                archive_path,
                as_attachment=True,
                download_name=archive_path.name,
                mimetype="application/zip",
            )

        return send_file(
            tex_path,
            as_attachment=True,
            download_name=f"{output_stem}.tex",
            mimetype="text/x-tex",
        )
    except Exception as exc:  # noqa: BLE001
        flash(f"Conversion failed: {exc}")
        return redirect(request.url)


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
