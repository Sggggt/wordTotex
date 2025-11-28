# Word to LaTeX Converter

Python project that converts Word `.docx` files into LaTeX `.tex` documents. It focuses on common structures (headings, paragraphs, bold/italic/underline, bullet and numbered lists, and basic tables) and emits a minimal LaTeX document you can compile.

## Setup
1. 一步安装依赖（使用当前系统 Python 环境）：
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

## Web 前端 (简洁上传转换)
一个轻量的 Flask 前端，上传 `.docx` 并下载生成的 `.tex`。

启动服务：
```powershell
python setup_env.py
python app.py
```

打开浏览器访问 `http://localhost:8000`，上传文件并选择是否包含 preamble、表格外框即可下载转换结果。

如遇本地防火墙限制，可以在运行前调整端口：
```powershell
set WORDTOTEX_HOST=127.0.0.1
set WORDTOTEX_PORT=8000
python app.py
```

## Features
- Maps Word headings to LaTeX section commands.
- Preserves inline formatting (bold/italic/underline), basic colors, and monospace fonts.
- Keeps paragraph alignment (居中/右对齐) and maps Word lists to `itemize`/`enumerate`.
- Converts tables to `tabular` while keeping cell alignment and inline text formatting.
- Optionally wraps output with a minimal preamble (`--no-preamble` to disable).

## Notes
- The conversion aims for readable LaTeX, but complex Word formatting (footnotes, images, equations, custom styles) is not handled. These will need manual edits in the generated `.tex`.
- Always review the `.tex` output, especially for documents with advanced formatting.
