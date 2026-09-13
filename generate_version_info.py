"""
Generates version_info.txt (Windows file-properties metadata for the built
exe) from the VERSION file, so the two can never drift out of sync. Run
before building - build.bat and CI both do this automatically:

    python generate_version_info.py
"""

from pathlib import Path

REPO_ROOT = Path(__file__).parent

version = (REPO_ROOT / "VERSION").read_text(encoding="utf-8").strip()
parts = (version.split(".") + ["0", "0", "0"])[:3]
major, minor, patch = (int(p) for p in parts)
version_tuple = f"({major}, {minor}, {patch}, 0)"

CONTENT = f"""# Windows file-properties metadata for the built exe
# (Explorer > Properties > Details). GENERATED from VERSION by
# generate_version_info.py - don't edit by hand, edit VERSION and
# re-run the generator instead.
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={version_tuple},
    prodvers={version_tuple},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'EDMS'),
        StringStruct(u'FileDescription', u'EDMS DataBridge'),
        StringStruct(u'FileVersion', u'{version}'),
        StringStruct(u'InternalName', u'EDMSDataBridge'),
        StringStruct(u'LegalCopyright', u'Copyright (c) 2026 Ash Kapow'),
        StringStruct(u'OriginalFilename', u'EDMSDataBridge.exe'),
        StringStruct(u'ProductName', u'EDMS DataBridge'),
        StringStruct(u'ProductVersion', u'{version}')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
"""

if __name__ == "__main__":
    (REPO_ROOT / "version_info.txt").write_text(CONTENT, encoding="utf-8")
    print(f"Wrote version_info.txt for version {version}")
