# Repository Guidelines

## Project Structure & Module Organization

UniOffice implements Windows COM compatibility components for office automation. `unioffice_excel/` contains the C implementation, with interfaces split into `i_*.c` files, shared helpers, and an IDL definition. `unioffice_word/` contains the smaller C++ Word component. `MSI/` holds installer tables (`idt/`), packaging utilities (`bin/`), and its Makefile. `tests/` contains VBScript and native COM regression probes, `tests/Excel/` the seven legacy suite scripts and `tests/compat/` Wine/Windows difference reproducers; `scripts/` contains build, test and publication tooling; `docs/` holds test reports and working notes.

## Build, Test, and Development Commands

Build from the repository root using GNU Make:

- `make`: build both DLLs and the MSI installer.
- `make -C unioffice_excel`: build only the Excel DLL.
- `make -C unioffice_word`: build only the Word DLL.
- `make -C MSI`: package the built Excel DLL and type library.
- `make clean`: remove build outputs, including generated headers and type libraries.

The default cross-compiler prefix is `i586-pc-mingw32`, configured in `Makefile.mingw`; override it with `make MINGW=<prefix>` when appropriate. IDL generation requires Wine's `widl` and headers under `/usr/include/wine/windows`. Packaging currently invokes `wine98` and `winepath`; verify these legacy dependencies before building the installer.

## Coding Style & Naming Conventions

Follow neighboring C/C++ code: four-space indentation, existing brace placement, and established COM interface signatures. Keep interface implementations in `i_<interface>.c` and preserve naming such as `MSO_TO_OO_I_Workbook_QueryInterface`. Use tabs for Makefile recipes. Edit `.idl` definitions rather than hand-editing generated declarations. Preserve license notices, existing text encodings, and a final newline. No formatter or linter configuration is provided; compiler flags include `-Wall`.

## Testing Guidelines

Tests use Windows Script Host or native clients that instantiate `Excel.Application`, backed by Windows Apache OpenOffice 4.1.17. Run them only in a disposable Wine prefix or a dedicated Windows test VM; they create documents and quit the office.

- `tests/*.vbs`: per-area probes (`windows.vbs`, `range-properties.vbs`, ...). Each ends with a `<Name>: PASS` marker; success is the marker, not the exit code.
- `tests/*.c`: native COM probes (`com-smoke.c`, `range-value-errors.c`, ...), built with MinGW (PE32) or `winegcc -b x86_64-windows` (PE64) and linked with `-lole32 -loleaut32 -luuid`.
- `tests/Excel/`: legacy `test_1..7.vbs` with CP1251 assertions; use `test_template.vbs` for new sequential cases. Results go to `otchet.txt`.
- `tests/compat/`: minimal reproducers for Wine/Windows differences, see its README.

Runners:

- Wine: `WINEPREFIX=<prefix> WINE=/usr/lib/wine-etersoft/wine64 CLIENT_BITS=32|64 xvfb-run -a sh tests/run-openoffice-backend.sh <probe.vbs|exe> '<PASS marker>'`.
- Legacy suite in Wine: `python3 scripts/run-excel-suite.py --bits 32|64`; it requires complete assertion totals and completion markers. Runner self-tests: `PYTHONDONTWRITEBYTECODE=1 python3 -m unittest discover -s tests -p test_excel_suite_runner.py`.
- Windows: `scripts\run-windows-tests.ps1 -Mode All|Suite|Probes [-Bits 32|64] [-Only probe.vbs]` from 64-bit PowerShell.
- Installer: `scripts/test-windows-msi.ps1` (install/repair/uninstall on Windows), `scripts/test-msi-coexistence.py` (MSI32 + MSI64 in Wine).

Record results in `docs/`: `docs/README.md` indexes the reports, `docs/native-windows-results.md` holds Windows and package acceptance, `docs/*-results.md` per-area results, `docs/supported-usage.md` the tested scope. Update `docs/*packages.sha256` when accepted packages change. Logs go under `~/tmp/claude`. No automated coverage threshold is configured.

## Commit & Pull Request Guidelines

History uses short English action summaries, such as “Fix makefiles for compile.” Write English commit titles and separate unrelated changes. Pull requests should explain affected interfaces, behavior changes, build/test results, and environment limitations; link relevant issues. Publish with `gpush`, not `git push`, to target the configured `pub.*` repositories.
