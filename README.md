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

`make` builds libraries. MSI packaging is explicit and detects the DLL's PE
architecture. It supports i386 and AMD64; other architectures are rejected.

```sh
# On builder64, build and package winegcc PE64:
make msi TOOLCHAIN=winegcc

# Build MinGW PE32 on Fedora, then package its shared artifacts on builder64:
WINEPREFIX=/path/to/test-prefix WINE=/usr/lib/wine-etersoft/wine64 \
  xvfb-run -a make -C MSI EXCEL_BUILD=../unioffice_excel/build/mingw-32 \
  OUTPUT=build/mingw-32/unioffice.msi
```

Outputs are `MSI/build/winegcc-64/unioffice.msi` and
`MSI/build/mingw-32/unioffice.msi`. Packaging uses Python 3 and bundled
CABARC/MsiDb under Wine. MsiDb needs a graphical session; Xvfb is sufficient.
The script stages tables in `~/tmp/claude`, updates file sizes, preserves CP1251
text, and publishes only a completed package. Existing `MSI/unioffice.msi`
is preserved. `WINE` may include loader arguments.

MSI64 has independent product/component/upgrade identities and installs into
system32; MSI32 installs into syswow64 on 64-bit Windows. Both variants passed
metadata/Unicode checks. Coexistence, repair and both removal orders were tested
under Wine; the surviving architecture remains usable. See
[MSI64 validation](docs/msi64-results.md) for identities and reproduction.
The MSI import workaround is tracked in
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

Both PE32/MinGW and PE64/winegcc passed all seven legacy tests with Windows
OpenOffice 4.1.17 x86 (zero failures),
plus dedicated font-script, outline visibility and diagonal-border probes.
See the [detailed results](docs/openoffice-4.1.17-results.md) for exact counters
and coverage limits, and [PE64 results](docs/pe64-results.md).

For PE64 testing, use a separate prefix with the PE64 DLL/TLB registered in
system32, set `CLIENT_BITS=64` for shell probes, and pass `--bits 64` to the Python
suite. Both selectors default to 32. Tests print the script host to identify the
actual client architecture. Both MSI32 and MSI64 packaging are supported.

`tests/com-smoke.c` checks only the factory and TLB, without an office backend.
Other legacy warnings remain.

Detailed findings, test outcomes, and pending work are recorded in
[development notes](docs/development-notes.md). Its historical Winelib experiments
are not the intended output format of this project.
