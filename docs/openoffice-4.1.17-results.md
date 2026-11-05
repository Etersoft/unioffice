# Windows OpenOffice 4.1.17: integration results

Date: 2026-11-05. Host: builder64, Wine@Etersoft 11.20.1-eter1,
Windows Apache OpenOffice 4.1.17 x86, UniOffice MinGW PE32 built on Fedora.
This run does not establish PE64 automation compatibility.

## Environment and basic checks

Installer: `Apache_OpenOffice_4.1.17_Win_x86_install_en-US.exe`, downloaded from
the official SourceForge mirror and checked against Apache's SHA-256 file.
Extraction: `7z x -oinstaller`. Install the extracted `openoffice4117.msi` with
`/qn /norestart ADDLOCAL=ALL SETUP_USED=1` under Xvfb. Without SETUP_USED the
package's launch condition requires setup.exe. Installed into a dedicated prefix:
`/home/lav/tmp/claude/unioffice-aoo/prefix`.

Both basic probes passed:

- Direct COM: ServiceManager, Desktop, Calc creation, cell write/read, close.
- UniOffice: Excel.Application, Workbooks.Add, cell write/read, Close, Quit.

An immediate first COM attempt failed before office/profile initialization;
a subsequent run after starting headless OpenOffice passed. The backend probe
now retries ServiceManager activation for up to 30 attempts. Wine cscript can
return exit code zero after a runtime error, so the harness requires a PASS marker.

## Legacy test results

Run: `scripts/run-excel-suite.py`, under Xvfb with the above WINEPREFIX and
`WINE=/usr/lib/wine-etersoft/wine64`. Each original test runs in a separate
working directory through explicit 32-bit cscript. Checks and source encodings
are preserved; blocking MsgBox dialogs are replaced with completion files.

| Test | Reported success | Reported failure | Result |
| --- | ---: | ---: | --- |
| test_1.vbs | 28 | 4 | Failed |
| test_2.vbs | 12 | 1 | Failed |
| test_3.vbs | 28 | 2 | Failed |
| test_4.vbs | 9 | 24 | Failed |
| test_5.vbs | 32 | 0 | Passed |
| test_6.vbs | 12 | 0 | Passed |
| test_7.vbs | 13 | 0 | Passed |
| Total | 134 | 31 | Partial compatibility |

These are the legacy scripts' counters, not independently audited assertion counts.
Test 6 reports 12 successes but emits only 11 SUCCESS markers; other counters
match their markers. The runner returned 1 because some tests failed.

Failures:

- Test 1: PUT/GET IFont.SubScript; PUT/GET IFont.SuperScript.
- Test 2: IOutline.Showlevels.
- Test 3: PUT Border LineStyle; PUT Border ColorIndex.
- Test 4: PUT/GET each of PageSetup LeftMargin, RightMargin, TopMargin,
  BottomMargin, Orientation, Zoom, CenterHorizontally, CenterVertically,
  FooterMargin, HeaderMargin, FitToPagesTall, FitToPagesWide.

The Subscript/Superscript getters are E_NOTIMPL stubs in UniOffice. Other
failures require separate diagnosis before attributing them to Wine or OpenOffice.
No Wine issue was filed solely from these compatibility test failures.

The first asynchronous legacy-chain attempt mixed 32- and 64-bit cscript and
failed to finish. Its cascade of errors is superseded by the isolated runs above.

## Artifacts and next work

Full logs and individual otchet.txt files:
`/home/lav/tmp/claude/unioffice-excel-suite-qp5hmjge/test_<n>/`.
The earlier invalid chain is retained under `unioffice-excel-suite-5hr5n_7l`.
Office download, MSI installation logs, and initial COM traces are under
`/home/lav/tmp/claude/unioffice-aoo/`.

Next: diagnose PageSetup as a group, implement the font getters, then isolate
Outline and Border failures. Re-run affected cases after fixes. Keep winegcc
builds on builder64 and direct MinGW builds on Fedora. The tested prefix is
retained with both OpenOffice and UniOffice installed for continued work.
