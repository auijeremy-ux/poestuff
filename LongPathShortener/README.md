# Long Path Shortener

A guide for staff. There is a section for IT at the end.

## What it does

Clients often send zip files or folders with folders inside folders inside folders, and very long file names. Once those files sit in one of our synced SharePoint folders, the full path (the folder names plus the file name) gets too long:

- Excel will not open a file whose full path is over 218 characters.
- Word will not open a file whose full path is over 259 characters.
- OneDrive will not sync very long paths.

This tool works out how long every path will be once the files are in the SharePoint folder where they will finally live. It shortens folder and file names only where it has to. It then copies the files (or unzips them) into a separate folder with the new names, ready for you to upload.

## What it never does

- It never changes, renames or deletes your original zip file or folder.
- It never changes what is inside a file. Only names change. Every file gets a digital fingerprint (a SHA-256 hash) as it is read and again after it is written, and the two must match.
- It never connects to the internet or sends anything anywhere.
- It never uploads anything. You upload the result yourself once you are happy with it.
- It never overwrites a file that is already there.

## One-time setup

1. Copy the whole `LongPathShortener` folder to your computer, for example to `C:\CL\Tools\LongPathShortener`. Keep all the files together.
2. Optional: right-click `Shorten Long Paths.bat`, choose **Send to > Desktop (create shortcut)**. You can drag zips onto the shortcut.
3. If Windows shows a security warning when you run it, ask IT to unblock the folder.

## How to use it

1. Save the client's zip file somewhere on your computer, for example Downloads. **Do not unzip it first.** If the client sent a folder, leave it where it is.
2. Drag the zip file (or folder) onto `Shorten Long Paths.bat` or its shortcut. A black window opens.
3. The window asks where the files will finally live:
   - In File Explorer, open the matter's synced SharePoint folder.
   - Click in the address bar at the top and copy the path. It looks like `C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd`.
   - Right-click in the black window to paste it, then press Enter.
   - Or type `B` and press Enter to pick the folder from a list.
   - Next time, pressing Enter on its own reuses the last folder.
4. **Step 1 is only a check.** Nothing is copied or renamed. When it finishes, the report opens in Excel. Have a look through it (see "Reading the report" below).
5. Close Excel, go back to the black window and type `Y` to go ahead, or `N` to stop without changing anything.
6. **Step 2** copies the files with their new names into `C:\CL\Out` and opens that folder.
7. Upload the **contents** of the renamed folder (for example everything inside `C:\CL\Out\Smith docs`) into the SharePoint folder. Drag the contents, not the folder itself. The lengths were worked out for files sitting directly in the destination folder, so an extra folder level could push some paths over the limit again.
8. Deal with anything in the "Needs attention" folder (see below). Keep the log file with the matter.

## What you get

Everything goes into `C:\CL\Out`. For a zip called `Smith docs.zip`:

| Item | What it is |
| --- | --- |
| `Smith docs` | The renamed files. Upload the contents of this folder. |
| `Smith docs - Needs attention` | Files that could not be shortened enough. Only there if there are any. |
| `Smith docs - Dry run report <date time>.csv` | The report from the check in step 1. |
| `Smith docs - Report <date time>.csv` | The report from step 2, showing what actually happened. |
| `Smith docs - Log <date time>.csv` | The evidence log. Keep it with the matter. |

If you run the same zip again, nothing is overwritten. The tool makes `Smith docs (2)` instead.

## Reading the report

The report has one row for every file, and for any empty folders. Problems are listed first.

| Column | What it means |
| --- | --- |
| OriginalPath | Where the file would have ended up in SharePoint with its original name. |
| NewPath | Where it will end up with its new name. Blank for files that were skipped or rejected. |
| OriginalLength | How many characters are in OriginalPath. |
| NewLength | How many characters are in NewPath. The limit is 218 less a safety margin of 10, so this should be 208 or less. |
| RulesApplied | Which renaming rules changed the path (see "How names are shortened"). Blank if nothing changed. A rule listed here may have changed a folder above the file rather than the file name itself. |
| Status | What happened to the file (see below). |

### Status values

| Status | What it means |
| --- | --- |
| OK | Already fits. Name unchanged. |
| Renamed | Now fits. |
| Needs manual attention | Still too long after all the shortening the tool is allowed to do. See "Files that need manual attention". |
| Expanded (nested zip) | A zip inside the zip (or folder). It was unpacked into a folder with the same name, without `.zip`. |
| Skipped - system file | Mac or Windows housekeeping, not a document: `__MACOSX` folders, `.DS_Store`, `Thumbs.db`, `desktop.ini`. |
| Skipped - encrypted | A password-protected file inside the zip. See "Password-protected files". |
| Skipped - online-only file | The file is in OneDrive but not downloaded to this computer, and the tool will not download it. Right-click the folder, choose **Always keep on this device**, wait for it to download, then run the tool again. |
| Skipped - shortcut or link | Links to other folders are not followed. |
| Rejected - unsafe path (zip slip) | An item in the zip tried to write outside the folder it should be unzipped into. This is a known trick used by harmful zip files. Nothing was written. Tell IT if you see this in a client's zip. |
| Error - ... | Something went wrong copying that one file. The message explains what. Other files are not affected. |

**Apostrophes:** if a file name starts with `=`, `+`, `-` or `@`, the report and log show an apostrophe (`'`) in front of it. This stops Excel treating a client's file name as a formula. The apostrophe is not part of the real name.

## How names are shortened

The rules run in this order and stop as soon as a path fits:

1. **Clean-up.** This always applies to every name. It removes spaces at the start and end, removes full stops at the end, and replaces characters SharePoint does not accept (`" * : < > ? / \ |`) with `_`. It turns runs of spaces into one space. It adds `_` to names Windows reserves, such as CON, PRN, AUX, NUL, COM1 and LPT1, so `CON.txt` becomes `CON_.txt`. Names starting with `~$` or containing `_vti_` are also adjusted, because SharePoint blocks them.
2. **Abbreviations** from `Abbreviations.csv`. Whole words only, and capitals do not matter. For example `Correspondence` becomes `Corro`.
3. **Repeated parent name removed.** `Smith Pty Ltd\Smith Pty Ltd - Invoices` becomes `Smith Pty Ltd\Invoices`. The last folder of the SharePoint destination counts as a parent too.
4. **Filler words removed** from folder names only: "the", "and", "of", "for".
5. **Folder names shortened.** The longest folder names on the path are cut first, at a word break, and never below 12 characters. Cutting one folder shortens every path inside it, so this comes before cutting file names.
6. **File names shortened**, at a word break. The extension (such as `.pdf` or `.docx`) is always kept, and so is any date or document number at the start (such as `2026-03-14`, `20260314`, `14.03.2026` or `001`).

Rules 2 to 6 only touch names on a path that is too long. Every other name stays exactly as the client sent it.

If two names end up the same, the later one gets ` (2)`, ` (3)` and so on. Windows ignores capitals, so `Letter.pdf` and `LETTER.pdf` count as the same name.

Zips inside the zip become folders, up to 3 levels deep. Any deeper than that are copied as zip files.

## Files that need manual attention

Some paths cannot be made short enough without cutting names down to meaningless stubs. This usually happens because the client's folders are nested very deeply. The tool will not force these. Instead it:

- lists them in the report with Status **Needs manual attention**, with NewPath showing where each one was meant to go, and
- copies them all into the `... - Needs attention` folder, in one place without subfolders, with names short enough to open.

What to do:

1. Open the report and filter the Status column to **Needs manual attention**. NewPath shows where each file belongs.
2. The usual fix is a flatter folder structure for that part of the matter. For example, create a shorter, shallower folder in SharePoint (such as `Hearing bundle`) and put those files there. Or shorten a folder name that the firm controls.
3. Move or rename the files as needed, then upload them. If you are not sure how the documents should be organised, ask the supervising lawyer. Please do not just keep cutting names until they are meaningless.
4. The evidence log still has each of these files' fingerprints, so you can show their contents did not change.

## The evidence log

The log is the firm's record that only names changed, never contents. Keep it with the matter.

| Column | What it means |
| --- | --- |
| OriginalRelativePath | The file's path inside the zip or folder as the client sent it. |
| NewRelativePath | The file's path inside the output folder. For files that need attention, this is the name in the Needs attention folder. |
| SizeBytes | File size. |
| LastModified | The file's own last-modified date, in the local time of the computer that ran the tool. The copy keeps this date. |
| SHA256 | The file's fingerprint. Identical fingerprints mean identical contents. |
| Result | Copied, Copied to Needs attention folder, Skipped, Rejected or Error, with the reason. |

The first row is the fingerprint of the original zip itself. Zips inside zips, skipped items and rejected items all have rows too, so the log accounts for everything the client sent.

To check a file's fingerprint later, for example after it has been uploaded and downloaded again, open PowerShell and run:

```powershell
Get-FileHash -Algorithm SHA256 -LiteralPath "C:\path\to\the\file.pdf"
```

The Hash shown should match the SHA256 column in the log.

## Password-protected files

The tool cannot open password-protected files inside a zip. It lists them as **Skipped - encrypted** and leaves them out. To get them, double-click the original zip in File Explorer, open the file and enter the password the client gave you. If Windows cannot open it, ask IT.

## Changing the settings

The settings for the drag-and-drop launcher are at the top of `engine\Start-Interactive.ps1`. Open it in Notepad to change:

- the output folder (default `C:\CL\Out`)
- the path limit (default 218) and safety margin (default 10)
- whether zips inside zips are unpacked (default yes)
- whether a new zip of the renamed files is also made (default no)

Keep the limit at 218 even for matters with no spreadsheets. Someone may add one later.

## Editing the abbreviations

`Abbreviations.csv` has two columns, **Find** and **Replace**. Open it in Excel, add or change rows, then save it as **CSV UTF-8 (Comma delimited)**.

- Capitals do not matter when matching.
- Only whole words are replaced, so `Construction` matches "Construction" but not "Reconstruction".
- Longer phrases are tried first, so `Statement of Claim` wins over any shorter entry.

## Common problems

| Problem | What to do |
| --- | --- |
| The window flashes and closes, or says scripts are disabled | Ask IT. A firm policy may be blocking PowerShell scripts. |
| Dragging a file onto the .bat does nothing | Names containing `&` or `%` can confuse Windows. Double-click the .bat instead and drag the file into the black window. |
| "The output folder is inside a OneDrive or SharePoint synced folder" | Anything written there starts uploading straight away, including files that still need attention. Type anything other than YES to stop, then use the default `C:\CL\Out`. |
| "The destination folder path is already N characters long" | The SharePoint folder itself is too deep. Pick a shallower destination folder. |
| "Stopped: the zip contents add up to ... GB" | The zip is unusually large once unpacked. Check with IT before raising the limit. |

## For IT

**Requirements.** Windows 10 or 11 with the built-in Windows PowerShell 5.1 and .NET Framework 4.6.2 or later (current builds ship 4.8). No modules, no installs, no admin rights, no network access. Scripts are plain ASCII.

**Files.**

| File | Purpose |
| --- | --- |
| `Shorten Long Paths.bat` | Drag-and-drop launcher. Runs `powershell.exe -NoProfile -ExecutionPolicy Bypass` for its own process only. |
| `engine\Start-Interactive.ps1` | The prompts the launcher shows. |
| `engine\Shorten-LongPaths.ps1` | Command-line entry point. `Get-Help .\engine\Shorten-LongPaths.ps1 -Full` lists every parameter. |
| `engine\LongPathShortener.psm1` | All the logic. |
| `Abbreviations.csv` | Sample abbreviations. |
| `tests\` | Synthetic fixture generator and Pester tests. |

**Command line example.**

```powershell
.\engine\Shorten-LongPaths.ps1 -Source 'C:\Users\jsmith\Downloads\Smith docs.zip' `
    -DestinationPrefix 'C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd' `
    -AbbreviationsCsv .\Abbreviations.csv -ExpandNestedZips -Apply
```

Without `-Apply` it is a dry run. Other switches: `-OutputFolder`, `-MaxPathLength`, `-SafetyMargin`, `-CreateZip`, `-MaxTotalSizeGB`, `-AllowSyncedOutput`, `-IncludeSystemFiles`.

**Long paths.** All file access uses .NET `System.IO` with the `\\?\` prefix (`\\?\UNC\` for network paths). This bypasses the 260 character limit without the LongPathsEnabled registry setting. PowerShell cmdlets such as Get-ChildItem and Copy-Item are not used for file work. At startup the tool creates, reads and deletes a path of more than 300 characters in the output folder. If that fails it stops before touching anything.

**Zips.** Each entry is streamed straight from the zip to its shortened output path, so long original names never touch the disk. Zips inside zips are copied to a short temporary file under the output folder (`_work-xxxxxxxx`), which is deleted at the end. Entry dates are copied onto the extracted files. Entry names not marked as UTF-8 are read as UTF-8 if they are valid UTF-8 (Mac zips), otherwise in the OEM code page (zips made by Windows Explorer).

**Security.**

- Zip slip: entries with absolute paths, drive letters or `..` parts are rejected. As a second check, every output path is confirmed to sit inside the output folder before it is written.
- Encrypted entries are detected from the zip's own index (the central directory), because .NET Framework cannot tell. They are reported and skipped.
- Zip bombs: the declared uncompressed total is checked against `-MaxTotalSizeGB` (default 20) before anything is extracted. Each entry is also cut off if it produces more data than it declares.
- Output files are opened with `CreateNew`, so nothing is ever overwritten. Source files are opened read-only.
- Junctions and symbolic links in a source folder are not followed.
- OneDrive online-only placeholders (Offline, RecallOnOpen or RecallOnDataAccess attributes) are skipped rather than read, so the tool never makes OneDrive download anything.
- An output folder that overlaps the source folder is refused.

**Synced output warning.** The tool checks `%OneDrive%`, `%OneDriveCommercial%`, `%OneDriveConsumer%`, the synced library list under `HKCU\Software\SyncEngines\Providers\OneDrive` and `HKCU\Software\Microsoft\OneDrive\Accounts` (read only), and the destination folder itself. If the output folder is inside any of them, it asks the user to type YES.

**Execution policy.** `-ExecutionPolicy Bypass` on the command line is overridden by a Group Policy that sets AllSigned. In that case, sign the `.ps1` and `.psm1` files with the firm's code signing certificate.

**Tests.** Run from the `LongPathShortener` folder:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1
```

This uses the Pester 3.4 that ships with Windows PowerShell 5.1 and only synthetic data. The data is generated under `%TEMP%` and deleted afterwards. Set `LPS_KEEP_TEST_FILES=1` to keep it. The fixtures include a 12-level tree with names over 100 characters, reserved names, illegal characters, names that collide once shortened, dated file names, zips nested four deep, zip slip entries, an entry marked as encrypted and accented names.

The tests were developed and passed on PowerShell 7.4 on Linux with Pester 3.4.0. That run cannot exercise the `\\?\` handling, reserved device names on NTFS, the registry checks or the .bat. **Run the tests once on a firm Windows 10 or 11 laptop before staff use the tool.**

**Known limits.**

- One zip or folder at a time. If several are dragged onto the .bat, only the first is used.
- Only file contents and last-modified dates are carried over. Created dates, NTFS permissions and alternate data streams (such as the "downloaded from the internet" mark) are not.
- Report and log dates are in the local time of the computer that ran the tool.
