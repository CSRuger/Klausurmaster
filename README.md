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

## Building from Source

1. Install Python 3.11+ and ensure `pip` is available.
2. Install dependencies (if a `requirements.txt` file is added) or run `python -m pip install -r requirements.txt` once available.
3. Start the UI for local development:
	```powershell
	cd path\to\Karteisystem 2.0
	python -m main
	```
4. Build a distributable executable via PyInstaller:
	```powershell
	pyinstaller packaging/main.spec
	```
	The resulting binaries live under `dist/`.

## Support

Report issues or feature requests via the GitHub Issues page. Include your OS, app version, and (if relevant) the save-directory you selected.
