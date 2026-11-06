# UniOffice

UniOffice lets existing Windows applications automate an Excel-compatible COM
interface backed by OpenOffice. Clients create `Excel.Application`; UniOffice
translates the covered operations into office UNO calls through Automation.

The working backend is **Windows Apache OpenOffice 4.1.17 x86**, under native
Windows or Wine. **Linux LibreOffice/OpenOffice integration remains open**: the
current DLL expects a registered Windows `com.sun.star.ServiceManager`.
See the [Linux bridge findings and plan](docs/linux-office-bridge.md).

## Status and supported use

The tested subset includes workbooks and worksheets, cell values/formulas,
range indexing and enumeration, fonts, supported borders, page setup and
outlines. Word is an unfinished component and is excluded from the MSI.
See [supported functions, boundaries and installation](docs/supported-usage.md).

Both variants passed all seven original VBScript tests, targeted COM probes,
and MSI installation, repair, coexistence and both removal orders:

| Client/package | Build | Verified environments |
| --- | --- | --- |
| PE32 / MSI32 | MinGW on Fedora | Windows 10 Pro 22H2 x64; Wine@Etersoft 11.20.1-eter1 |
| PE64 / MSI64 | winegcc on builder64 | Windows 10 Pro 22H2 x64; Wine@Etersoft 11.20.1-eter1 |

Both use the same x86 OpenOffice backend. Native functional suites also passed
under a separate standard user with a fresh profile. These results cover the
stated configurations and scenarios, rather than full Microsoft Excel coverage.
Evidence: [native Windows](docs/native-windows-results.md) and
[Wine acceptance](docs/installed-acceptance-results.md).

## Install and run

Choose the package by the **client process bitness**:

- 32-bit client: `MSI/build/mingw-32/unioffice.msi`.
- 64-bit client: `MSI/build/winegcc-64/unioffice.msi`.

Install Windows OpenOffice first and complete its first-start setup in the
runtime user's profile. Install UniOffice with administrator privileges;
Automation itself runs without elevation. Both packages can coexist and MSI
handles COM registration. Older per-user packages must be removed under their
original installing identity before installing the current machine-wide MSI.

Follow the [Windows and Wine installation steps](docs/supported-usage.md).
Verify the accepted packages from the repository root:

```sh
sha256sum -c docs/native-windows-packages.sha256
```

That manifest identifies the tested artifacts. Repackaging generates a new
PackageCode and changes the MSI hashes.

## Build Windows PE DLLs

Requirements: GNU Make, Wine development headers and `widl`; the selected driver
also needs MinGW or Wine's `winegcc`/`wineg++` and its Windows-target compiler.
Use **direct MinGW on Fedora** and **winegcc on builder64**. Both hosts see the
same `/srv` checkout through NFS. Run from the repository root:

```sh
# Fedora: direct MinGW, PE32.
make libraries TOOLCHAIN=mingw MINGW=i686-w64-mingw32 -j4

# builder64: winegcc with Clang/lld, PE64.
make libraries TOOLCHAIN=winegcc -j4
```

Both produce PE `.dll` files. Outputs, generated headers and type libraries are
isolated under `unioffice_{excel,word}/build/<toolchain>-<bits>/`.
`make` builds libraries; MSI packaging is explicit. To build only Excel, use
`make -C unioffice_excel` with the same toolchain arguments.

`TOOLCHAIN` defaults to `mingw`; its historical `MINGW` default is
`i586-pc-mingw32`. `BITS` defaults to 32 for MinGW (64 for an `x86_64` prefix)
and 64 for winegcc. The winegcc target defaults to `x86_64-windows` for PE64;
`WINE_TARGET` overrides it. Compiler selection lives in [Makefile.mingw](Makefile.mingw).

Override `IDL`, `IDLFLAGS`, `CFLAGS`, `CXXFLAGS` or `LDFLAGS` as needed. Fedora's
MinGW `stdole2.tlb` directory is detected automatically; elsewhere supply
`IDLFLAGS="-I /usr/include/wine/windows -L /path/to/typelibs"`.
After changing compiler versions or flags, clean only the selected variant,
for example `make clean TOOLCHAIN=winegcc`.

## Build MSI packages

Package on builder64 after building the matching DLL/TLB. Packaging needs
Python 3, Wine, a configured disposable `WINEPREFIX`, and a graphical session
or Xvfb. Set `WINE` to the chosen loader, for example
`/usr/lib/wine-etersoft/wine64`, then run:

```sh
# Build and package winegcc PE64.
xvfb-run -a make msi TOOLCHAIN=winegcc

# Package the MinGW PE32 artifacts already built on Fedora.
xvfb-run -a make -C MSI EXCEL_BUILD=../unioffice_excel/build/mingw-32 \
  OUTPUT=build/mingw-32/unioffice.msi
```

The packager uses bundled CABARC/MsiDb, stages CP1251 tables in `~/tmp/claude`,
detects i386/AMD64 from the DLL and writes the architecture-specific output.
MSI64 has separate product/component/upgrade identities. Both packages install
machine-wide; on 64-bit Windows, MSI32 uses SysWOW64 and MSI64 uses System32.
See [MSI identities and lifecycle checks](docs/msi64-results.md) and the
[Wine IDT import issue](https://bugs.etersoft.ru/show_bug.cgi?id=19875).

## Run functional tests

With the backend and both packages installed, use ordinary **64-bit PowerShell**
in a dedicated Windows test account:

```powershell
.\scripts\run-windows-tests.ps1 -Mode Suite
# Select one installed architecture with -Bits 32 or -Bits 64.
```

For Wine, use an isolated prefix containing Windows OpenOffice and the matching
UniOffice package. The existing PE32 fixture on builder64 is shown below:

```sh
mkdir -p "$HOME/tmp/claude"
export TMPDIR="$HOME/tmp/claude"
export WINEPREFIX="$HOME/tmp/claude/unioffice-aoo/prefix"
export WINE=/usr/lib/wine-etersoft/wine64 CLIENT_BITS=32
xvfb-run -a sh tests/run-openoffice-backend.sh
xvfb-run -a sh tests/run-openoffice-backend.sh excel-backend.vbs 'UniOffice integration: PASS'
xvfb-run -a python3 scripts/run-excel-suite.py --bits 32
```

For PE64, select its own prefix, `CLIENT_BITS=64` and `--bits 64`. Direct backend
MSI installation needs `SETUP_USED=1`; see [backend setup](docs/openoffice-4.1.17-results.md).
Tests create documents and quit the office; close other office work first.

The runners preserve the legacy assertions/CP1251 encoding and require complete
assertion totals, completion markers and zero process status. A zero exit from
Wine cscript alone is insufficient. Logs and reports go under `~/tmp/claude`.
Runner result checks can be exercised without Wine:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest discover -s tests -p test_excel_suite_runner.py
```

Full native probes and destructive installer tests are described in the
[native report](docs/native-windows-results.md). Regression details are indexed
in the [documentation guide](docs/README.md).

## Related projects

| Project | Relevant work | Fit for the Linux backend |
| --- | --- | --- |
| [LibreOffice Automation Bridge](https://github.com/LibreOffice/core/tree/master/extensions/source/ole) | COM/UNO object wrapping and type conversion | Main implementation reference; the existing bridge is built for Windows |
| [Apache OpenOffice Automation Bridge](https://github.com/apache/openoffice/tree/trunk/main/extensions/source/ole) | The Automation contract used by the current backend | Existing Windows implementation and documentation, requiring adaptation for Linux |
| [Collabora COLEAT](https://github.com/CollaboraOnline/COLEAT) | Redirects Word/Excel COM calls to Collabora Office; tracing and generated proxies | Windows compatibility tooling; dormant according to its README, not a ready Linux bridge |
| [LibreOffice SDK DocumentLoader](https://github.com/LibreOffice/core/blob/master/odk/examples/cpp/DocumentLoader/DocumentLoader.cxx) | Native C++ client connecting to an office over UNO/URP | Source-based starting point for the host-side prototype |
| [unoserver](https://github.com/unoconv/unoserver) | Runs LibreOffice for document conversion through Python UNO | Useful office process-management reference; does not expose a COM frontend |

The search did not identify a ready-to-use Wine COM ↔ Linux UNO bridge.
See [research notes and source links](docs/related-projects.md) for the search
scope, reuse candidates and remaining gap. None of these is a new dependency.

## Repository and documentation

`unioffice_excel/` contains the C Excel adapter; `unioffice_word/` the unfinished
C++ Word component; `MSI/` installer tables/tools; `scripts/` build/test runners;
`tests/` VBScript and native COM regressions; `docs/` setup, evidence and design.

Start with the [documentation guide](docs/README.md). The
[development notes](docs/development-notes.md) retain historical experiments;
their Winelib builds are not the current output format.
License: [GNU LGPL 2.1 or later](COPYING), as stated in the source headers.
