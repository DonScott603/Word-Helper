# Word Helper

A small dark-themed desktop app for batch operations on Word (`.docx`) files.
The first tool is **Find & Replace** across many documents at once, with **no
255-character limit** (unlike Word's built-in dialog).

## Download

**[⬇ Download WordHelper.exe (latest release)](https://github.com/DonScott603/Word-Helper/releases/latest)**

A standalone 64-bit Windows executable — no Python or install required. Download,
double-click, done. (Unsigned, so on first run Windows may show a SmartScreen
prompt: **More info → Run anyway**.)

## Running

Double-click **`Word Helper.bat`**, or from a terminal:

```
py word_helper.py
```

## Standalone .exe (no Python needed)

A single-file `WordHelper.exe` can be built and copied to any 64-bit Windows
machine — the target machine does **not** need Python or any packages installed.

To build it:

```
py -m pip install pyinstaller
build_exe.bat
```

The result is `dist\WordHelper.exe` (~16 MB). Copy that one file anywhere and
double-click it. Note: the first launch may take a few seconds as it unpacks,
and antivirus / SmartScreen may warn about an unsigned executable the first time
(click "More info" → "Run anyway").

## Find & Replace tab

1. **+ Add files** — pick one or more `.docx` files (repeat to add more).
2. Type into **Find** and **Replace with** — both boxes are multi-line and
   accept text of any length (well over 255 characters).
3. Options:
   - **Match case** — off by default (case-insensitive).
   - **Scope** — `everywhere` (body, tables, headers, and footers) or `body`.
4. **Run replace**.

### Safety

Before a file is overwritten, the original is copied to `<name>.docx.bak` in the
same folder. If a `.bak` already exists it is left untouched so your first
original is always preserved. Files with no matches are not modified.

Close a document in Word before running, or saving will fail with a
"could not save (is it open in Word?)" message.

## Replace Image tab

Swap an embedded image across many documents at once — pick the **old** image
(the one currently in the documents) and the **new** image to put in its place.

1. **Old image → Choose…** — the picture as it appears in the documents now.
2. **New image → Choose…** — the replacement. A thumbnail and its size/format
   are shown for both.
3. **+ Add files** — the `.docx` files to update.
4. Optionally enable **"Also match by size if not an exact match"** — Word
   sometimes re-compresses images on insert, so the bytes may not be identical;
   this falls back to matching any image with the same pixel dimensions.
5. **Replace image**.

The new image is displayed at the **original's size and position** (only the
image data is swapped). Identical images used in several places are all updated
in one pass, including images in headers and footers. Different image formats
are handled (e.g. replacing a PNG with a JPG). Same `.bak` backup as above.

## Formatting tab

Find text and apply character formatting to **only the matched text** — and only
the attributes you choose. Anything left alone is untouched.

1. **+ Add files** and enter the **text to find**.
2. For **Bold / Italic / Underline / Strikethrough**, each control is
   three-state: **Leave** (don't change), **On**, or **Off**.
3. Tick **Font**, **Size (pt)**, or **Color (hex)** and fill in a value to set
   those; leave them unticked to keep them as-is.
4. **Apply formatting**.

Only the found text is changed — surrounding text and any formatting you didn't
select are preserved. Works across runs, tables, headers, and footers, with the
same `.bak` backup.

## Notes / limitations

- Matching happens within each paragraph. A Find string that spans a paragraph
  break (a hard Enter in Word) won't match; long single-paragraph text is fine.
- Text split across formatting "runs" is handled; the replacement inherits the
  formatting of the run where the match begins.

## Adding more tabs

`word_helper.py` is structured for growth. To add a feature:

```python
self.tabview.add("My Feature")
self.build_my_feature_tab(self.tabview.tab("My Feature"))
```

and write a `build_my_feature_tab(self, tab)` method. Reusable document logic
lives in `docx_replace.py`.

## Files

- `word_helper.py` — the GUI application.
- `docx_replace.py` — the find/replace engine (no GUI; reusable/testable).
- `Word Helper.bat` — double-click launcher.
- `requirements.txt` — dependencies.
