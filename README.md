# UniOffice

UniOffice provides legacy Excel and Word COM compatibility components.

## Building Windows DLLs

GNU Make, Wine development tools (`widl`, `winegcc`/`wineg++` for that driver),
are required; direct MinGW additionally needs its cross-compiler. Both compiler choices produce **PE
Windows DLLs**, not Winelib `.dll.so` libraries.

```sh
# Direct MinGW driver:
make libraries TOOLCHAIN=mingw MINGW=i686-w64-mingw32 -j4

# Wine compiler driver targeting MinGW PE:
make libraries TOOLCHAIN=winegcc -j4
```

Use Fedora for direct MinGW and builder64 for winegcc. Both see the shared checkout:

```sh
ssh builder64 'cd /srv/lav/Projects/obsoleted/UniOffice/unioffice && make libraries TOOLCHAIN=winegcc -j4'
```

`TOOLCHAIN` defaults to `mingw`, and the historical `MINGW` default is
`i586-pc-mingw32`. The winegcc driver receives `-b x86_64-windows` and uses Clang/lld with Wine
headers and import libraries. Override `WINE_TARGET` if needed. `BITS` defaults
to 64 for winegcc and 32 for MinGW (64 with an `x86_64` prefix). Outputs and generated headers are isolated under
`unioffice_{excel,word}/build/<toolchain>-<bits>/`, including `.dll` and `.tlb`.

`make -C unioffice_excel TOOLCHAIN=winegcc` builds one
component. Override `IDL`, `IDLFLAGS`, `CFLAGS`, `CXXFLAGS`, or `LDFLAGS` as needed.
Fedora's MinGW `stdole2.tlb` directory is detected automatically; elsewhere add
`IDLFLAGS="-L /path/to/typelibs"`. After changing flags or compiler versions,
clean the selected variant with `make clean TOOLCHAIN=winegcc`.

## Installer and validation

`make` builds libraries. MSI packaging is explicit and currently supports only
PE32 i386: the legacy installer has no 64-bit component/template support.
Build MinGW on Fedora, then package those artifacts on builder64:

```sh
WINEPREFIX=/path/to/test-prefix WINE=/usr/lib/wine-etersoft/wine64 \
  xvfb-run -a make -C MSI EXCEL_BUILD=../unioffice_excel/build/mingw-32 \
  OUTPUT=build/mingw-32/unioffice.msi
```

Packaging uses Python 3 and the bundled CABARC/MsiDb tools under Wine. MsiDb
requires a graphical session (Xvfb is sufficient). The script stages tables in
`~/tmp/claude`, updates file sizes, preserves CP1251 text, and publishes the MSI
only after successful imports. Existing `MSI/unioffice.msi` is preserved.
`WINE` may include loader arguments. PE64 input is rejected before packaging.
MSI32 has passed a Unicode metadata check, administrative extraction with matching
payload hashes, and silent installation with COM registration in an isolated
Wine prefix; silent uninstallation also passed. The MSI import workaround is tracked in
[Wine bug 19875](https://bugs.etersoft.ru/show_bug.cgi?id=19875).

## OpenOffice integration tests

Windows Apache OpenOffice 4.1.17 x86 has been installed and checked on builder64
in `~/tmp/claude/unioffice-aoo/prefix`. Its MSI requires `SETUP_USED=1` when
installed directly. UniOffice MSI32 is installed into the same prefix.
The backend test creates a Calc document through ServiceManager and verifies a
cell round trip; the UniOffice test does the same through Excel.Application.

```sh
export WINEPREFIX="$HOME/tmp/claude/unioffice-aoo/prefix"
export WINE=/usr/lib/wine-etersoft/wine64
xvfb-run -a sh tests/run-openoffice-backend.sh
xvfb-run -a sh tests/run-openoffice-backend.sh excel-backend.vbs 'UniOffice integration: PASS'
xvfb-run -a python3 scripts/run-excel-suite.py
```

Use a dedicated prefix: the tests create documents and terminate the office.
The suite runner copies the seven legacy tests into separate temporary directories,
preserves their checks/encoding, replaces blocking MsgBox dialogs with completion
markers, and explicitly uses 32-bit cscript. Inspect `otchet.txt` and `script.log`
in the printed result directory. Wine cscript may return zero after a script
error, so completion markers and assertion totals determine success.

The OpenOffice 4.1.17 run completed all seven legacy tests: after fixing PageSetup, tests 4–7 passed;
tests 1–3 report 7 remaining failed checks. See the
[detailed results](docs/openoffice-4.1.17-results.md) for limitations and failures.

`tests/com-smoke.c` checks only the factory and TLB, without an office backend.
Other legacy warnings remain.

Detailed findings, test outcomes, and pending work are recorded in
[development notes](docs/development-notes.md). Its historical Winelib experiments
are not the intended output format of this project.
