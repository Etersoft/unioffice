# Borders enumeration regression

Validated on builder64 with Wine@Etersoft 11.20.1-eter1 and Windows Apache
OpenOffice 4.1.17 x86. PE32 was built with MinGW on Fedora; PE64 with winegcc
on builder64. Native Windows was not tested.

## Defect and behavior

Previously all GetEnumerator calls shared one cursor. A nested VBScript
For Each therefore consumed the outer loop's remaining elements; the new
tests/borders-enumeration.vbs failed in its first inner loop on the old DLL.
Skip also stopped at count minus one, repeated the final item, returned S_FALSE
when skipping exactly the remaining count, and could overflow on ULONG_MAX.
Clone returned E_NOTIMPL.

Each GetEnumerator now creates an independently referenced collection context
retaining the same range and UNO backend, with a fresh cursor. Clone retains
the current position in another independent context. Skip compares its unsigned
request to the remaining count before advancing, and uses count as the end
position. Next clears its fetched count and publishes a returned dispatch only
after successful item creation. Related null output arguments return E_POINTER.

This follows the COM contracts for
[Skip](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-ienumvariant-skip),
[Next](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-ienumvariant-next)
and [Clone](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-ienumvariant-clone).
At this stage Count=12 and index aliases 1 through 4 remained. The later
[membership change](borders-membership-results.md) enumerates eight distinct
keys while retaining alias lookup. Neither stage establishes full Excel
Borders collection semantics or exact native Excel ordering.

## Tests

tests/borders-enumerator.c obtains the actual range Borders COM interface.
It tests independent cursors, exact/end/overflow Skip, partial and exhausted
Next, Reset, Clone at a live position and at end, null outputs, and range
lifetime after releasing the original collection. The old DLL failed cursor,
Skip and Clone assertions. The initial runner also incorrectly required S_OK
from workbook Close; this was corrected to SUCCEEDED before final validation.

The VBScript test requires every inner loop and the outer loop to visit Count
items, checks their product (144 with the current count), and repeats the loop.
The native and script tests both pass for PE32 and PE64. The existing diagonal
probe also passes in both, and legacy test_3 reports 30 successes, zero failures
and 30 SUCCESS markers for each architecture. Other legacy tests were not
repeated in this stage.

The backend runner now accepts a repository-relative .exe path in addition to
its existing VBScript names. It starts OpenOffice and checks readiness before
each test, then requires the exit status and explicit success marker.
Excel.Quit terminates the backend: do not chain a new native test after a VBS
probe without starting OpenOffice again.

## Reproduction

After building the corresponding DLL, build the regression executable:

```sh
# Fedora, repository root:
i686-w64-mingw32-gcc -Wall -Iunioffice_excel/build/mingw-32 \
  tests/borders-enumerator.c unioffice_excel/build/mingw-32/unioffice_excel_i.o \
  -lole32 -loleaut32 -luuid -o unioffice_excel/build/mingw-32/borders-enumerator.exe

# builder64, repository root:
winegcc -b x86_64-windows -m64 -Wall -Iunioffice_excel/build/winegcc-64 \
  tests/borders-enumerator.c unioffice_excel/build/winegcc-64/unioffice_excel_i.o \
  -lole32 -loleaut32 -luuid -o unioffice_excel/build/winegcc-64/borders-enumerator.exe
```

Use a disposable OpenOffice/UniOffice prefix with the matching updated DLL
installed. Set WINEPREFIX before invoking any Wine tools. Existing fixtures are
~/tmp/claude/unioffice-aoo/prefix for PE32 and unioffice-aoo64/prefix for PE64.
Stop only that prefix's server before replacing an installed DLL; copy to a
temporary sibling and rename. Run within one Xvfb session:

```sh
export WINE=/usr/lib/wine-etersoft/wine64 WINEDEBUG=-all CLIENT_BITS=64
sh tests/run-openoffice-backend.sh \
  unioffice_excel/build/winegcc-64/borders-enumerator.exe 'Borders enumerator: 0 failure(s)'
sh tests/run-openoffice-backend.sh borders-enumeration.vbs 'Borders enumeration integration: PASS'
sh tests/run-openoffice-backend.sh border-diagonal.vbs 'Border integration: PASS'
python3 scripts/run-excel-suite.py --bits 64 --tests 3
```

For PE32 use the other prefix, CLIENT_BITS=32, the mingw-32 executable and
--bits 32. Avoid stale OpenOffice processes tied to an earlier Xvfb display.
An initial readiness attempt was stopped in its disposable prefix and repeated
with bounded waits; this was not classified as a Wine defect.

Logs: ~/tmp/claude/unioffice-enumerator{32,64}-green.log,
unioffice-enumeration-vbs{32,64}-green.log and
unioffice-diagonal{32,64}-enumerator.log. Legacy results:
unioffice-excel-suite-z3rb7nqi (PE32) and unioffice-excel-suite-6i6t0x6k (PE64).
Both MSI packages were rebuilt from the updated DLLs; architecture/Unicode
metadata checks returned zero failures for each. Installer lifecycle was not
repeated because the change affects the Borders implementation.
These are UniOffice defects; no new Wine issue was filed.
