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

## Follow-up: PageSetup fixed

The shared failure was a PROPERTYGET of StyleFamilies. All 24 call sites now
invoke the documented UNO getStyleFamilies method with DISPATCH_METHOD.
Both MinGW/Fedora and winegcc/builder64 builds passed. Targeted test 4 passed,
then all seven tests were re-run with the updated PE32 DLL.

Regression results (`~/tmp/claude/unioffice-excel-suite-_8c6q9m1`):

| Test | Reported success | Reported failure |
| --- | ---: | ---: |
| 1 | 28 | 4 |
| 2 | 12 | 1 |
| 3 | 28 | 2 |
| 4 | 34 | 0 |
| 5 | 32 | 0 |
| 6 | 12 | 0 |
| 7 | 13 | 0 |

Failed checks decreased from 31 to 7. Tests 4–7 pass; no new failures appeared.
Test 4's success counter is 34 versus 33 markers; test 6 remains 12 versus 11.
These legacy counter discrepancies do not affect the remaining failure list.

## Follow-up: font scripts fixed

Subscript/Superscript getters are implemented; setters now use VT_I2 on the
cell text cursor because the range property set has no CharEscapement. The
operation covers every cell, with Null returned for mixed range states.

New tests/font-script.vbs passed enabling/disabling, by-reference and numeric
inputs, mutual switching, a 2x2 range and mixed-state reads. The original font
test passed 32/0 (`unioffice-excel-suite-wo8_d87e`). Both compiler builds pass.
The remaining known failures are Outline.ShowLevels and two Border operations.

## Follow-up: Outline.ShowLevels fixed

Marked the IDL result as retval, handled both axes and optional/zero arguments,
and mapped Excel's 1-based levels to Calc's 0-based depth. The new outline-levels
probe verified actual row/column visibility and passed all checks. Original
test 2 passed 13/0 (unioffice-excel-suite-l2f44geo). Both drivers compile.
Only the two Border failures remain to be diagnosed.

## Final regression: Border fixed (2026-11-06)

Diagonal BorderLine structures are retrieved through explicit UNO
XPropertySet.getPropertyValue instead of the failing dynamic PROPERTYGET path.
Both diagonal directions now work. Root cause inside the Automation bridge
was not established, so this is not recorded as a proven Wine defect.

The new tests/border-diagonal.vbs checks LineStyle/ColorIndex readback for all
12 legacy indices, direct UNO diagonal color/width, double lines and clearing.
It passed. The original test 3 passed 30/0. Both compiler builds passed.

All seven tests were then re-run together on updated PE32 DLL/TLB with
Windows OpenOffice 4.1.17. Results at
`~/tmp/claude/unioffice-excel-suite-dh6epjsc`:

| Test | Reported success | Reported failure | SUCCESS markers |
| --- | ---: | ---: | ---: |
| 1 | 32 | 0 | 32 |
| 2 | 13 | 0 | 13 |
| 3 | 30 | 0 | 30 |
| 4 | 34 | 0 | 33 |
| 5 | 32 | 0 | 32 |
| 6 | 12 | 0 | 11 |
| 7 | 13 | 0 | 13 |
| Total | 166 | 0 | 164 |

Runner exit status: 0. All seven completion markers were checked. The legacy
counter discrepancies in tests 4 and 6 remain documented rather than silently
counted as independently verified assertions. This validates the covered PE32
automation scenarios, not every Office interface or full PE64 automation.

The final MSI32 was rebuilt with the current DLL/TLB. Unicode metadata check
passed; administrative extraction in the clean test prefix succeeded and cmp
confirmed that both extracted payload files match the tested artifacts byte
for byte. An /a attempt in the prefix with the product installed failed at
InstallAdminPackage; the prefix-state difference has not been isolated.
