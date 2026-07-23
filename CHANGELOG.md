# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [v2.1.1] - 2026-07-23

### Added
- **Custom Branding Assets**:
  - `assets/logo.ico`: Windows file icon for `Outlook_Contact_Updater_Wizard.exe`.
  - `assets/unnamed.jpg`: UI header logo stamp integrated into the Minecraft grassland banner.
- **Python Dependencies File**:
  - Added `requirements.txt` listing `pywin32`, `Pillow`, `mcp`, and `pyinstaller`.

### Changed
- Updated `wizard_gui.py` to display the custom window icon and UI logo image.
- Updated `build_exe.py` with `--icon=assets/logo.ico` and `--add-data="assets;assets"`.

---

## [v2.1.0] - 2026-07-23

### Added
- **Minecraft 8-Bit Retro Sound Effects**:
  - `sounds/click.wav`: Retro frequency sweep for button clicks.
  - `sounds/xp_gain.wav`: Chime sound on EXP progress updates.
  - `sounds/victory.wav`: 4-note victory fanfare upon task completion.
- **Sound Manager (`sound_manager.py`)**: Asynchronous non-blocking audio playback using Windows `winsound`.
- **Git Version Control & Repository Standards**:
  - Semantic Versioning (`v2.1.0`).
  - Remote repository integration: `https://github.com/Undefinedgod2718/Outllok-update.git`.
  - Structured `.gitignore` and `CHANGELOG.md`.

---

## [v2.0.0] - 2026-07-23

### Added
- **Minecraft Rural Grassland GUI Wizard (`wizard_gui.py`)**:
  - Dirt block header, pixel font, green EXP progress bar.
  - Automatic Outlook detection and one-click process termination.
  - 3 Import Modes: `OVERWRITE` (覆蓋匯入), `COEXIST` (共存匯入), `SMART_MERGE` (比對匯入).
- **Embedded UnRAR Extractor (`tools/UnRAR.exe`)**: Enables direct decompression of `.rar`, `.zip`, and `.pst` archives.
- **COM MAPI Engine (`mcp_outlook_core.py`)**: Robust Outlook COM handling with `safe_copy_and_move` retry logic.
- **Real Machine Test Suite (`test_three_functions.py`)**: 100% test coverage with `260702mailaddress.rar`.
- **Single Executable Build (`build_exe.py`)**: Standalone PyInstaller compilation.

---

## [v1.0.0] - 2026-07-02

### Added
- Initial PowerShell import script (`import_contacts.ps1`).
- FastMCP Python Server (`MCP/server.py`) and test client (`MCP/test_client.py`).
- Initial documentation and wiki (`wiki.md`).
