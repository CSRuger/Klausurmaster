# Klausurmaster 2.0

Klausurmaster is a cross-platform Tkinter application for organizing study cards, grade tables, and ratios. The app now ships with platform-specific installers and keeps all runtime assets in a predictable structure.

## Repository Layout

- `assets/` – bundled resources such as the default config, sample data, and application icon.
- `cards/`, `formula/`, `table/`, `main/` – Python packages that make up the UI and domain logic.
- `docs/` – supplemental documentation and historical notes.
- `packaging/` – PyInstaller specification files for generating distributable builds.
- `build/`, `dist/` – generated folders created by PyInstaller (safe to delete/recreate).

## Download & Install via GitHub Releases

Every tagged release publishes three artifacts under the repository's **Releases** tab. Choose the one for your operating system and follow the instructions below.

### Windows (.exe installer)
1. Navigate to **Releases** → download `Klausurmaster-Setup-x.y.z.exe`.
2. Double-click the installer, approve the SmartScreen dialog if prompted.
3. Choose the installation directory and complete the wizard.
4. Launch Klausurmaster from the Start Menu shortcut. On first start the app asks where to store your tables—pick any folder or accept the per-user default.

### macOS (.dmg bundle)
1. Download `Klausurmaster-x.y.z.dmg` from the release assets.
2. Open the DMG, then drag **Klausurmaster.app** into the **Applications** shortcut.
3. On first launch, macOS Gatekeeper may warn about an unidentified developer. Open **System Settings → Privacy & Security**, allow Klausurmaster, then launch again via Spotlight.
4. The first-run dialog lets you pick a custom storage directory; confirm to finish setup.

### Linux (.AppImage)
1. Grab `Klausurmaster-x.y.z.AppImage` from the release assets.
2. Make it executable: `chmod +x Klausurmaster-x.y.z.AppImage`.
3. Run the AppImage from your file manager or terminal. Optional: move it into `~/Applications` for easier access.
4. Select a save folder when prompted; the app will reuse it on every start.

## Build & Package Yourself

The single PyInstaller spec in `packaging/main.spec` works on all three platforms. You must run the build natively on each OS to get a compatible binary.

### Windows
1. Install Python 3.11+ from python.org and run `py -m pip install --upgrade pip pyinstaller==6.11.1`.
2. From the repo root run `pyinstaller packaging/main.spec`.
3. The signed/installer-ready binary lives at `dist/Klausurmaster/Klausurmaster.exe`. Wrap it with Inno Setup/NSIS if you need an installation wizard.

### macOS
1. Use macOS 13+ with Xcode command-line tools installed: `xcode-select --install`.
2. Create a virtual environment (`python3 -m venv .venv && source .venv/bin/activate`) and install PyInstaller (`pip install pyinstaller==6.11.1`).
3. Run `pyinstaller packaging/main.spec`. The bundle appears at `dist/Klausurmaster/Klausurmaster`.
4. (Optional) Convert to a DMG:
	```bash
	brew install create-dmg
	create-dmg --overwrite --volname "Klausurmaster" \
	  release/macos/Klausurmaster.dmg dist/Klausurmaster/Klausurmaster
	```
5. Codesign or notarize the app if you have an Apple Developer ID. Otherwise document the Gatekeeper bypass (already described above).

### Linux
1. Install Python 3.11+, Tk headers, and PyInstaller: `sudo apt install python3-tk && python3 -m pip install pyinstaller==6.11.1`.
2. Run `pyinstaller packaging/main.spec` to produce `dist/Klausurmaster/Klausurmaster`.
3. (Optional) Build an AppImage:
	```bash
	wget https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage
	chmod +x appimagetool-x86_64.AppImage
	# Prepare AppDir
	mkdir -p release/linux/Klausurmaster.AppDir/usr/bin
	cp dist/Klausurmaster/Klausurmaster release/linux/Klausurmaster.AppDir/usr/bin/
	cp assets/favicon.ico release/linux/Klausurmaster.AppDir/Klausurmaster.png
	cat > release/linux/Klausurmaster.AppDir/Klausurmaster.desktop <<'EOF'
	[Desktop Entry]
	Name=Klausurmaster
	Exec=Klausurmaster
	Icon=Klausurmaster
	Type=Application
	Categories=Education;
	EOF
	ARCH=x86_64 ./appimagetool-x86_64.AppImage \
	  release/linux/Klausurmaster.AppDir release/linux/Klausurmaster-x.y.z.AppImage
	chmod +x release/linux/Klausurmaster-x.y.z.AppImage
	```

## Release Checklist

1. Build fresh installers on Windows (`.exe`), macOS (`.dmg` or zipped `.app`), and Linux (`.AppImage` or raw binary) using the steps above.
2. Rename artifacts consistently, e.g. `Klausurmaster-Windows-x.y.z.exe`, `Klausurmaster-macOS-x.y.z.dmg`, `Klausurmaster-Linux-x.y.z.AppImage`.
3. Tag the commit: `git tag -a vX.Y.Z -m "Release vX.Y.Z" && git push origin main --tags`.
4. Draft a GitHub release for that tag, attach all platform binaries, and include highlights/changelog plus SHA256 checksums if available.
5. Update this README (and optionally `docs/stand-19.12.txt`) if installation behavior or requirements change.

## Support

Report issues or feature requests via the GitHub Issues page. Include your OS, app version, and (if relevant) the save-directory you selected.
