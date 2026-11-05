# PE64 Automation validation

Date: 2026-11-06. Host: builder64. Compiler: winegcc/wineg++ targeting
x86_64-windows (Clang/lld). Runtime: Wine@Etersoft 11.20.1-eter1.
Backend: Windows Apache OpenOffice 4.1.17 x86.

## Isolation and architecture

A stopped copy of the working PE32/OpenOffice prefix was created at
`/home/lav/tmp/claude/unioffice-aoo64/prefix` using cp -a --reflink=auto.
The original `unioffice-aoo/prefix` was preserved. The copy contains both DLLs:

- PE64 winegcc DLL and TLB in C:\windows\system32.
- PE32 MinGW DLL and TLB in C:\windows\syswow64.

The PE64 DLL was registered by system32\regsvr32.exe. Tests explicitly use
system32\cscript.exe for 64-bit clients; the readiness log prints WScript.FullName.
The client bitness is independent of the 32-bit out-of-process office server.

```sh
export WINEPREFIX="$HOME/tmp/claude/unioffice-aoo64/prefix"
export WINE=/usr/lib/wine-etersoft/wine64
export CLIENT_BITS=64
xvfb-run -a sh tests/run-openoffice-backend.sh
xvfb-run -a sh tests/run-openoffice-backend.sh excel-backend.vbs 'UniOffice integration: PASS'
xvfb-run -a python3 scripts/run-excel-suite.py --bits 64
```

## Results

Both direct OpenOffice COM and the Excel.Application workbook/cell round trip
passed. The full seven-test suite returned zero:

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

Results: `~/tmp/claude/unioffice-excel-suite-8qgvlq1p`. Known legacy counter
inconsistencies in tests 4 and 6 remain; zero failed checks and completion markers
were verified. Dedicated font-script.vbs, outline-levels.vbs and border-diagonal.vbs
also passed via the 64-bit script host. This establishes the covered Excel PE64
scenarios with an x86 office server, not complete Office API compatibility.

## Type library unregister bug

The PE64 DLL used SYS_WIN32 in UnRegisterTypeLib. With both registrations present,
a PE64 /u reported success but deleted the win32 typelib entry while retaining
win64. The test tests/typelib-registration.c reproduced two failures on the old DLL.

Commit ce5d3fd selects SYS_WIN64 under _WIN64. The regression test now passes for
both PE64 and PE32: own registration disappears, the peer path is preserved,
and own registration is restored afterward. After this cycle, original test 2
passed 13/0 through both clients in the cloned prefix:
`unioffice-excel-suite-ldpjs0yw` (64-bit) and `unioffice-excel-suite-0j9voquh` (32-bit). The test requires a disposable prefix
with both architectures registered; do not run it in a normal user environment.

Compile the native test using winegcc -b x86_64-windows on builder64 or
MinGW i686-w64-mingw32-gcc on Fedora, linking advapi32. Pass the DLL path as the
sole argument. Run each executable through the selected Wine loader in the cloned
prefix. Runtime test binaries are in the corresponding component build directories.

This was a UniOffice defect, so no Wine bug was created. The independent Windows
OpenOffice install remains 32-bit. MSI packaging still supports only PE32; these
results do not establish MSI64 support or functioning Word Automation.
