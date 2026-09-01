EDMS brand assets, provided by EDMS and generated with Pillow.

- `logo.png` — the "ED" circle mark, resized to 96x96, shown at the top
  of the app window. Loaded automatically by `load_logo_image()` in
  `edms_databridge.py`; the app falls back to a plain text title if this
  file is ever missing.
- `logo.ico` — the same mark, saved as a multi-resolution icon
  (16/24/32/48/64/128/256px). Used as the app's title bar/taskbar icon
  at runtime (`App.__init__` in `edms_databridge.py`), and can also be
  passed to PyInstaller via `--icon assets\logo.ico` to brand the built
  `.exe` itself.
- `banner.png` — the wide "Emergency Doctors Medical Service" lockup,
  resized to 800x450, used at the top of the repo's README.
- `source/` — the original, unedited files EDMS provided
  (`edms-icon-source.jpg` -> logo.png/logo.ico, `edms-banner-source.jpg`
  -> banner.png), kept so any of the above can be regenerated at a
  different size later without needing to track someone down for the
  originals again.

Regenerating any of these (e.g. at a different size) just needs Pillow:

```python
from PIL import Image
icon = Image.open("path/to/source.jpg").convert("RGBA")
icon.resize((96, 96), Image.LANCZOS).save("assets/logo.png")
icon.save("assets/logo.ico", sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
```
