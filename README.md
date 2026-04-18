# All-in-One PDF Management Tool

A simple web-based PDF toolkit built with Python (Flask) and plain HTML/CSS/JS.
Runs entirely on your own computer — your files never leave your machine.

## What it does

| Feature | What it gives you |
|---|---|
| **Merge** | Combine multiple PDFs into one |
| **Split** | Explode a PDF into one-page-per-file (zipped) |
| **Extract Pages** | Pull out specific pages, e.g. `1-3,5,7-9` |
| **Extract Text** | Save all text to a `.txt` file |
| **Extract Tables** | Detect tables and save them to `.xlsx` (one table per sheet) |
| **Rotate** | Rotate every page 90°, 180°, or 270° |
| **Password** | Add or remove a password |
| **Compress** | Shrink file size with `qpdf` |
| **Watermark** | Add a diagonal text watermark to every page |
| **To Images** | Convert each page to a PNG (zipped) |
| **Extract Images** | Pull out images embedded in the PDF |
| **Images → PDF** | Combine JPG/PNG images into one PDF |
| **PDF Info** | Show page count, metadata, encryption status |

---

## Setup (first time only)

### 1. Install Python

If you don't already have Python 3.10+, download it from
[python.org](https://www.python.org/downloads/) and make sure to tick
**"Add Python to PATH"** during installation.

Check it's installed by opening Command Prompt / PowerShell / Terminal and running:

```
python --version
```

### 2. Install qpdf (for the Compress feature)

- **Windows:** Download from https://github.com/qpdf/qpdf/releases, extract, and add the `bin` folder to your PATH.
- **macOS:** `brew install qpdf`
- **Ubuntu/Debian:** `sudo apt install qpdf poppler-utils`

`poppler-utils` (the `pdfimages` command) is used by the "Extract Images" feature.
On Windows, download Poppler from https://github.com/oschwartz10612/poppler-windows/releases
and add its `bin` folder to PATH too.

### 3. Install Python libraries

Open a terminal in this folder and run:

```
pip install -r requirements.txt
```

(Coming from VBA: this is similar to "adding references" in the VBE, except
it downloads and installs the libraries for you.)

---

## Running the tool

From this folder, run:

```
python app.py
```

You'll see something like:

```
 * Running on http://127.0.0.1:5000
```

Open your browser and go to **http://127.0.0.1:5000**. That's it — pick a tab, upload a file, click the button.

To stop the server, press `Ctrl + C` in the terminal.

---

## How it's structured (for learning)

```
pdf_tool/
├── app.py               <- the "module" with all functions/subroutines
├── requirements.txt     <- list of Python libraries needed
├── templates/
│   └── index.html       <- the UI (like a VBA UserForm)
├── static/
│   └── style.css        <- colors, spacing, layout
├── uploads/             <- incoming files (auto-cleaned)
└── outputs/             <- generated files (auto-cleaned)
```

### Mapping VBA ideas to Flask

| VBA | Python / Flask |
|---|---|
| `Sub MyMacro()` | `def my_function():` |
| Module | A `.py` file |
| `MsgBox "Error"` | `return jsonify({"error": ...})` |
| UserForm button | HTML `<button>` connected to a Flask route |
| `Application.GetOpenFilename` | `<input type="file">` in the browser |
| Workbook | `Workbook` object from `openpyxl` |

### How a click travels through the code

1. You click a button in the browser.
2. JavaScript (inside `index.html`) grabs the form and sends it to a URL like `/merge`.
3. Flask sees `/merge` and runs the `merge_pdfs()` function in `app.py`.
4. That function processes the PDFs and returns the finished file.
5. JavaScript receives the file and triggers a download.

---

## Adding your own feature

Say you want to add a "Reverse page order" tool. You would:

1. **Add a function in `app.py`:**

```python
@app.route("/reverse", methods=["POST"])
def reverse_pdf():
    file = request.files.get("file")
    if not file or not allowed_file(file.filename):
        return jsonify({"error": "Please upload a PDF."}), 400
    src_path = save_upload(file)
    reader = PdfReader(str(src_path))
    writer = PdfWriter()
    for page in reversed(reader.pages):
        writer.add_page(page)
    out_path = OUTPUT_DIR / f"reversed_{uuid.uuid4().hex}.pdf"
    with open(out_path, "wb") as fp:
        writer.write(fp)
    return send_and_cleanup(out_path, "reversed.pdf", "application/pdf",
                            extra_to_clean=[src_path])
```

2. **Add a tab and form in `index.html`:**

```html
<button class="tab" data-target="tab-reverse">Reverse</button>
...
<section id="tab-reverse" class="panel">
    <h2>Reverse page order</h2>
    <form data-endpoint="/reverse" data-result="download">
        <input type="file" name="file" accept="application/pdf" required>
        <button type="submit">Reverse</button>
    </form>
</section>
```

That's it — no other wiring needed. The JavaScript at the bottom of the page
automatically handles any form with a `data-endpoint` attribute.

---

## Notes

- Max upload size is **200 MB** per request (set in `app.py`).
- The server runs in **debug mode** by default so you can see errors in the browser.
  Change `debug=True` to `debug=False` in the last line of `app.py` before sharing.
- Files are deleted from the server right after each operation.
- The tool is designed for local use. If you ever want to run it on a public server,
  you'll need to add a production WSGI server (e.g. Gunicorn) and proper auth.
