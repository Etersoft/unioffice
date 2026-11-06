# Installed-package acceptance, 2026-11-06

The existing Excel Automation subset passed an end-to-end acceptance run:
full rebuild, packaging, MSI installation, functional tests, repair and removal.
No new COM functionality was added in this stage. Usage and scope are described
in [supported usage](supported-usage.md).

## Artifacts and environment

- PE32: full `make -B libraries TOOLCHAIN=mingw MINGW=i686-w64-mingw32 -j4`
  on Fedora, shared checkout under /srv.
- PE64: full `make -B libraries TOOLCHAIN=winegcc -j4` on builder64;
  Clang 21.1.8 and WIDL 11.20.1-eter1, targeting Windows PE.
- Runtime: Wine@Etersoft 11.20.1-eter1 and Windows Apache OpenOffice 4.1.17 x86;
  Wine's built-in cscript of the selected client architecture, one Xvfb session.
- Both MSI packages installed into `~/tmp/claude/unioffice-msi64/prefix`.
  Installed DLL and TLB bytes matched the build outputs before testing.
  Functional probes used registered, installed DLLs; no manual payload replacement.
- Excel implementation source: 4146fac. Strict suite-runner fix: 540fff4.
  Word builds, but its class factory remains a stub and it is excluded from MSI.

The accepted packages are MSI/build/mingw-32/unioffice.msi and
MSI/build/winegcc-64/unioffice.msi. Their SHA256 values are recorded in
[installed-packages.sha256](installed-packages.sha256). From the repository root:

```sh
sha256sum -c docs/installed-packages.sha256
```

These identify this acceptance run's artifacts. Repackaging generates a fresh
PackageCode, so later builds will have different hashes.

## Functional results

| Check | PE32 | PE64 |
| --- | --- | --- |
| DLL exports, factory lifecycle, TLB loading | 0 failures | 0 failures |
| TypeLib registration/unregistration with peer preserved | 0 failures | 0 failures |
| Application dual ABI, common native COM, Range enumerator | 0 failures | 0 failures |
| 18 existing VBS integration probes | PASS | PASS |
| All seven original Excel tests, strict result checks | PASS | PASS |
| MSI architecture, identity and Unicode metadata | 0 failures | 0 failures |

The VBS probes cover direct backend and Excel cell round trips, font scripts,
outline levels, diagonal/border formatting, supported weights/mixed results,
unsupported-style errors, Range/Cells ownership and addressing, row/column
collections and enumeration. Existing assertion counts include Cells 49, Item 36,
bounded Rows/Columns 98, sheet Rows/Columns 177 and enumeration 166.

Each legacy suite reports counters 158/0 with 156 successful assertion markers
and no failed markers. Tests 4/6 have documented extra counter increments.
Suite result directories on builder64:

- PE32: `~/tmp/claude/unioffice-excel-suite-89eyeiej`
- PE64: `~/tmp/claude/unioffice-excel-suite-ey1b95w7`

## Fixed validation defect

The old runner accepted an empty report and completion counters `0 0` as success.
A fake loader reproduced this with exit 0. After 540fff4 the identical input exits 1.
The runner now requires the complete baseline of assertions, valid numeric
completion counters, no failed report markers and successful process statuses.
Seven Python unit tests cover valid legacy exceptions and false-pass cases.
The original CP1251 VBScript sources/assertions were preserved.

## Installer lifecycle

The coexistence runner passed both repairs: 32-bit and 64-bit. After each repair,
both clients remained usable and both payloads matched. Removing 64 first left 32
usable; after reinstalling 64, removing 32 left 64 usable. Final 64 removal deleted
the remaining payload and the shared Excel.Application ProgID. Both UniOffice
test products are now uninstalled from this disposable fixture; the OpenOffice
backend and other runtime prefixes remain available.

Reproduce after installing both packages into a disposable OpenOffice prefix:

```sh
export TMPDIR="$HOME/tmp/claude" WINEDEBUG=-all
export WINE=/usr/lib/wine-etersoft/wine64
export WINEPREFIX="$HOME/tmp/claude/unioffice-msi64/prefix"
xvfb-run -a python3 scripts/test-msi-coexistence.py
```

Logs on builder64: `~/tmp/claude/unioffice-completion-*`, plus
`~/tmp/claude/unioffice-completion/runner-false-pass-{before,after}.log`.
Installer action logs are under unioffice-msi64/coexistence-logs.

## Limits

Acceptance covers sequential clients in the stated Wine/OpenOffice environment.
Native Windows and broader concurrency behavior were not certified. The old
parallel PE32 test5 fault was not isolated; this stage's sequential suite passed.
Unsupported API/backend cases remain outside the supported scope; the new guide
lists them explicitly. No Wine defect was established in this stage.
