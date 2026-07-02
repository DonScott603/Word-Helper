# Word Helper

A small dark-themed desktop app for batch operations on Word (`.docx`) files.
The first tool is **Find & Replace** across many documents at once, with **no
255-character limit** (unlike Word's built-in dialog).

## Running

Double-click **`Word Helper.bat`**, or from a terminal:

```
py word_helper.py
```

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
