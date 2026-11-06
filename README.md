# OpenInDocs

**Double-click a Word, Excel or PowerPoint file on Windows and edit it in Google Docs, Sheets or Slides.**

[Français](README.fr.md)

Google Drive for desktop mounts your Drive as a disk, but it gives no way to open a local `.docx` in the Google editor:
Windows only offers Word. OpenInDocs fills that gap. It needs **no Google API, no API key and no sign-in**.

> Not affiliated with or endorsed by Google. "Google Docs", "Sheets", "Slides" and "Drive" are trademarks of Google LLC.

## What it does

1. You double-click a file (for example a freshly downloaded `contract.docx`).
2. A small window shows the upload progress. The file is copied into a folder of your Drive (`My Drive\Opened from PC` by default); the original stays where it was.
3. As soon as Google Drive for desktop has uploaded it, the file opens in **Google Docs** (`.docx`), **Sheets** (`.xlsx`) or **Slides** (`.pptx`) in your default browser.

Files already inside your Drive open immediately, with no copy. Opening the same file twice never creates a duplicate.
If anything goes wrong (offline, Drive paused...), you are offered to open the file in Word / Excel / PowerPoint instead.

## Requirements

- Windows 10 or 11
- [Google Drive for desktop](https://www.google.com/drive/download/) installed and signed in
- Nothing else: the installer uses the C# compiler and SQLite library that ship with Windows

## Install

1. Download `OpenInDocs.zip` from the [latest release](../../releases/latest) and unzip it (or clone this repository).
2. Right-click `installer\Install.ps1` > **Run with PowerShell**. No administrator rights needed.
3. Windows' *Default apps* page opens: choose **Google Docs**, **Google Sheets** and **Google Slides** as default apps for their file types.
   (Alternative: double-click a `.docx`, pick *Google Docs*, then *Always*.)

To go back to Word for a single file: right-click > *Open with* > *Word*.
To uninstall: `powershell -ExecutionPolicy Bypass -File installer\Install.ps1 -Uninstall`.

Run `%LOCALAPPDATA%\OpenInDocs\OpenInDocs.exe` without a file to get a diagnostic of your setup.

## Settings

`%LOCALAPPDATA%\OpenInDocs\OpenInDocs.ini` (see [`config/OpenInDocs.ini`](config/OpenInDocs.ini)):

| Setting | Meaning |
|---|---|
| `TargetFolder` | Folder inside My Drive where files are copied |
| `GoogleAccount` | Account to use if your browser is signed in to several Google accounts |
| `MaxWaitSeconds` | How long to wait for the upload |
| `DriveRoot` | Location of My Drive, if automatic detection fails |

## Icons

Files get the Docs / Sheets / Slides icons. They are **not** shipped in this repository: the installer extracts them
from your own Google Drive for desktop installation. If they cannot be found or verified, default icons are used and
everything else works the same.

## How it works

See [docs/HOW_IT_WORKS.md](docs/HOW_IT_WORKS.md). In short: Drive for desktop keeps a local SQLite database of your files;
OpenInDocs reads it (read-only) to find the Drive ID of the copied file, then opens `docs.google.com/document/d/<ID>/edit`.

## Limitations

- **Relies on an undocumented database.** A Drive for desktop update may change its layout and break the lookup. The tool then shows an error message saying what failed instead of failing silently. Please open an issue if this happens.
- Developed and tested on Windows 11 with a French display language, one Google account and Drive in *Stream files* mode. *Mirror files* mode and other languages should work but are untested. Reports are welcome.
- `.docx` files open in Google's Office-compatibility mode and stay `.docx` on Drive (use *File > Save as Google Docs* to convert).
- Files are uploaded to **your** Drive. Don't use this for files you must not store there.
- Windows 10/11 do not allow programs to set default apps by themselves, so one manual choice per file type is required.

## Privacy

Nothing is sent anywhere by OpenInDocs itself: no telemetry, no network access. The file is uploaded by Google Drive for desktop,
and your browser opens a Google URL. The local database is only opened read-only.

## Contributing

Issues and pull requests are welcome, in particular reports for other Windows languages, Drive versions and account types
(Workspace, shared drives). Continuous integration compiles the three launchers on Windows for every change.

## License

[MIT](LICENSE)
