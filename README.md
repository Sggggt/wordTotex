# Word to LaTeX Converter

Python project that converts Word `.docx` files into LaTeX `.tex` documents. It focuses on common structures (headings, paragraphs, bold/italic/underline, bullet and numbered lists, and basic tables) and emits a minimal LaTeX document you can compile.

## Setup
1. Install dependencies in the current Python environment:
   ```powershell
   python setup_env.py
   ```

## Usage
Convert a file and write the result to `output.tex`:
```powershell
python -m wordtotex .\input.docx -o .\output.tex
```

Output to stdout instead:
```powershell
python -m wordtotex .\input.docx --no-preamble
```

Or use the helper script on Windows to set up the environment (if needed) and run conversion:
```powershell
.\run_wordtotex.bat .\input.docx -o .\output.tex
```

## Web frontend (upload and download)
A lightweight Flask frontend lets you upload a `.docx` and download the converted output.

```powershell
python setup_env.py
python app.py
```

Open `http://localhost:8000`, upload a file, and choose whether to include the preamble and table borders. If your local firewall blocks the default port, adjust it before running:
```powershell
set WORDTOTEX_HOST=127.0.0.1
set WORDTOTEX_PORT=8000
python app.py
```

## Features
- Maps Word headings to LaTeX section commands.
- Preserves inline formatting (bold/italic/underline), basic colors, and monospace fonts.
- Keeps paragraph alignment and maps Word lists to `itemize`/`enumerate`.
- Converts tables to `tabular` while keeping cell alignment and inline text formatting.
- Extracts embedded images into an `<tex_name>_images/` folder next to the `.tex` and inserts relative `\includegraphics` paths (the web flow returns a zip when images exist).
- Optionally wraps output with a minimal preamble (`--no-preamble` to disable).

## Notes
- The conversion aims for readable LaTeX, but complex Word formatting (footnotes, equations, custom styles) may need manual edits in the generated `.tex`.
- Always review the `.tex` output, especially for documents with advanced formatting.
