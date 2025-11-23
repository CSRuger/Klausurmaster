# Klausurmaster Release Guide

Follow these steps whenever you want to ship a new multi-platform release.

## 1. Prepare the workspace

1. Sync with main: `git pull origin main`.
2. Remove old `dist/` and `build/` folders to avoid stale binaries.
3. Ensure Python 3.11+ and PyInstaller 6.11.1 are installed on each target OS.

## 2. Build per platform

### Windows
```powershell
py -m venv .venv
.\.venv\Scripts\activate
py -m pip install --upgrade pip pyinstaller==6.11.1
pyinstaller packaging/main.spec
```
Artifact: `dist/Klausurmaster/Klausurmaster.exe`.

### macOS
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip pyinstaller==6.11.1
pyinstaller packaging/main.spec
```
Optional DMG:
```bash
brew install create-dmg
create-dmg --overwrite release/macos/Klausurmaster.dmg dist/Klausurmaster/Klausurmaster
```

### Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip pyinstaller==6.11.1
pyinstaller packaging/main.spec
```
Optional AppImage (after build):
```bash
ARCH=x86_64 ./appimagetool-x86_64.AppImage release/linux/Klausurmaster.AppDir release/linux/Klausurmaster-x.y.z.AppImage
```

## 3. Sanity checks

- Launch each binary to confirm it starts, loads assets, and prompts for a save folder on first run.
- Verify icons display correctly on each OS.
- Run a quick smoke test: create a table, add cards, save, reopen.

## 4. Tag and push

```powershell
git status
# ensure clean
git tag -a vX.Y.Z -m "Release vX.Y.Z"
git push origin main --tags
```

## 5. Publish GitHub release

1. Open the repo → **Releases** → **Draft a new release**.
2. Choose tag `vX.Y.Z`, set the title, and summarize changes.
3. Upload the Windows `.exe`, macOS `.dmg` (or `.zip`), and Linux `.AppImage`.
4. Optionally attach SHA256 checksum files.
5. Click **Publish release**.

## 6. Post-release

- Update README or docs if instructions changed.
- Announce the release (socials, README badge, etc.).
- Monitor Issues for bug reports tied to the new version.
