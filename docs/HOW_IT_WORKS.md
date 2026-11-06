# How it works

## The problem

Google opens a Drive file in its editor with a URL that contains the file's **Drive ID**:

```
https://docs.google.com/document/d/<ID>/edit          .docx / Docs
https://docs.google.com/spreadsheets/d/<ID>/edit      .xlsx / Sheets
https://docs.google.com/presentation/d/<ID>/edit      .pptx / Slides
```

A `.docx` stored on Drive can be edited there directly (Office-compatibility mode, no conversion).
But when Windows hands a path such as `G:\My Drive\report.docx` to a program, that path says nothing about the ID.
The official way to get it is the Drive API, which requires a Google Cloud project, OAuth consent and tokens.

## The trick: Drive for desktop already knows the ID

Drive for desktop keeps a SQLite database of everything it syncs:

```
%LOCALAPPDATA%\Google\DriveFS\<account>\metadata_sqlite_db
```

The relevant parts (observed on Drive for desktop 133.x; **not a documented format**):

| Table | Columns used |
|---|---|
| `items` | `stable_id`, `id` (the Drive ID), `local_title` (file name), `is_folder`, `trashed` |
| `stable_parents` | `item_stable_id`, `parent_stable_id` |

- The root of My Drive is the only folder without a row in `stable_parents`. Its name depends on the Drive language
  ("My Drive", "Mon Drive", "Meine Ablage"...), so it is read from the database instead of being hard-coded.
- To find the ID of `A\B\file.docx`: take the items titled `file.docx`, then walk up the parents and check that the folder names
  are `B`, `A` and that the chain ends at the root. This avoids mixing up two files with the same name in different folders.
- A file that was just copied appears first with a temporary ID `local-<n>`. Once uploaded, the row gets its real ID
  (about 8 seconds for a 40 MB file in our tests). OpenInDocs polls until the ID no longer starts with `local-`.
- The database is held open by Drive, so it is opened read-only through `winsqlite3.dll` (shipped with Windows).
  If that fails because of a lock, the database and its `-wal` / `-shm` files are copied to a temporary folder and read from there.
- Drive's local root folder name contains a non-breaking space in French ("Mon&nbsp;Drive"), hence the normalisation before comparing names.

## The flow

```
double-click  ->  OpenInDocs.exe "C:\Users\me\Downloads\contract.docx"
                    |
                    |-- already in My Drive?  yes -> look up ID -> open URL
                    |                         no  -> copy to My Drive\<TargetFolder>
                    |                                (reuse the copy if identical: SHA-256)
                    |-- poll the database until the ID is real (default: up to 90 s)
                    '-- open https://docs.google.com/.../d/<ID>/edit in the default browser
                        (on failure: offer Word / Excel / PowerPoint)
```

## Three executables

`OpenInDocs.exe`, `OpenInSheets.exe` and `OpenInSlides.exe` are built from the same source
(`/define:DOCS|SHEETS|SLIDES`). Only the name and the icon differ, which lets Windows show
"Google Docs", "Google Sheets" and "Google Slides" with their own icons in *Open with* and *Default apps*.
The editor chosen is based on the file extension, not on the executable.

They are built on the user's machine by the installer, with `csc.exe` from the .NET Framework 4 that ships with Windows,
so there is nothing to download or trust besides this source code.

## Icons

Icons of Docs / Sheets / Slides belong to Google and are not redistributed. The installer reads them from the user's own
`GoogleDriveFS.exe` (PE resources `RT_GROUP_ICON` / `RT_ICON`), rebuilds `.ico` files, and **verifies each one by its dominant colour**
(blue / green / yellow) before using it, so a Drive update that reshuffles resource IDs results in default icons rather than wrong ones.

## Known weak points

- The database layout can change with any Drive for desktop update.
- Shared drives, "Computers" backups and shortcuts are not handled specially.
- Several Google accounts in Drive for desktop: all databases are searched, the first match wins.
