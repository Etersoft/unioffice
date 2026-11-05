# Repository Guidelines

## Project Structure & Module Organization

UniOffice implements Windows COM compatibility components for office automation. `unioffice_excel/` contains the C implementation, with interfaces split into `i_*.c` files, shared helpers, and an IDL definition. `unioffice_word/` contains the smaller C++ Word component. `MSI/` holds installer tables (`idt/`), packaging utilities (`bin/`), and its Makefile. `tests/Excel/` contains VBScript integration tests; `scripts/` contains repository publication tooling.

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

Tests use Windows Script Host and instantiate `Excel.Application`. Configure the target COM implementation before running them. From `tests/Excel/`, run `cscript.exe /E:vbscript all_test.vbs` in a compatible Windows/Wine environment. Inspect `otchet.txt` after the chained scripts finish; launcher completion alone does not establish success. Use `test_template.vbs` for new sequential `test_<number>.vbs` cases. No automated coverage threshold is configured.

## Commit & Pull Request Guidelines

History uses short English action summaries, such as “Fix makefiles for compile.” Write English commit titles and separate unrelated changes. Pull requests should explain affected interfaces, behavior changes, build/test results, and environment limitations; link relevant issues. Publish with `gpush`, not `git push`, to target the configured `pub.*` repositories.
