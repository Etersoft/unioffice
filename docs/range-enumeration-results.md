# Range enumeration

`Range._NewEnum` now returns an independent `IEnumVARIANT` cursor instead of
`E_NOTIMPL`. Ordinary ranges and Cells enumerate single cells left-to-right,
then top-to-bottom. Rows enumerate row ranges; Columns enumerate column ranges.
Cells conversion restores single-cell enumeration without changing its source.
This follows the examples in Microsoft's
[Range.Cells documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cells)
and the collection ordering described for
[Range.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item).

## Ownership and COM behavior

Each cursor retains its source Range, which retains the backend and worksheet.
Returned dispatch objects have their own references and survive cursor release.
Separate `_NewEnum` calls start at the beginning; Clone copies the current
position into another independent cursor. Reset rewinds it. Skip clamps to the
end and distinguishes an exact skip (`S_OK`) from insufficient elements
(`S_FALSE`), without unsigned addition overflow.

Next returns owned `VT_DISPATCH` values, reports the fetched count and returns
`S_FALSE` for partial or exhausted requests. Its fetched-count pointer may be
NULL. Failed batches release produced objects and leave the cursor unchanged.
QueryInterface exposes IUnknown/IEnumVARIANT with a stable identity. Automation
returns `VT_UNKNOWN` for DISPID_NEWENUM, accepts no arguments and releases an
ignored result. Existing IDL, IIDs, method slots and registration are unchanged.
See the primary [Next contract](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-ienumvariant-next).

Enumeration is lazy: construction and Skip allocate no list of range elements.
Counts and positions use 64 bits; cell coordinates advance with row/column
counters. The first implementation divided a 64-bit position by the width,
which introduced a MinGW PE32 import of `libgcc_s_dw2-1.dll` and prevented COM
activation in the fixture. Using 32-bit coordinate arithmetic removed that new
dependency. Verify PE imports after arithmetic changes as well as after adding
libraries; a successful link alone does not establish a deployable DLL.

## Validation

`tests/range-enumeration.vbs` passes 166 checks on PE32 and PE64: ordered cells,
Rows/Columns shapes, writes through yielded cells, Cells conversion, nested
independent loops, early exit, singleton ranges and retained parent lifetime.
Whole-sheet Rows/Columns and Application.Cells stop after two elements, checking
lazy iteration without traversing the entire sheet.

`tests/range-enumerator.c` returns zero failures on both architectures: native
and Automation getters, QueryInterface identity/error outputs, Next batches and
partial/exhausted results, optional fetched counts, Clone/Reset/Skip, exact end,
ULONG_MAX skips, cursor/source lifetime and yielded objects after cursor release.
It also skips directly to the last cell of the actual backend sheet and checks
its coordinates. AOO4.1.17 has 1,048,576 rows and 1024 columns: 1,073,741,824 cells.

Build hosts: direct i686 MinGW on Fedora; PE64 winegcc on builder64. Runtime:
Wine@Etersoft11.20.1-eter1, Windows Apache OpenOffice4.1.17 x86, Wine's built-in
32/64-bit cscript and Xvfb, with separate disposable prefixes. Deployed DLLs
matched the build outputs. Both MSI packages were rebuilt and passed
architecture/Unicode metadata checks. Installer lifecycle was not repeated;
registration and IDL did not change.

Existing Application ABI and common native COM probes return zero failures on
both builds. Sheet Rows/Columns177, bounded Rows/Columns98, Cells49, Item36,
ownership, parent-lifetime and Range-address probes pass. All seven original
Excel tests pass sequentially on both: counters158/0 and156 SUCCESS markers,
with the known extra counters in test4/test6. Suite outputs:

- PE32: `~/tmp/claude/unioffice-excel-suite-k32i_04_`
- PE64: `~/tmp/claude/unioffice-excel-suite-bjqolzso`

Logs are in `~/tmp/claude/unioffice-range-enum-*`: the initial PE64 `-red.log`
stops at the first For Each before assertions; `-{32,64}-native-final.log` and
`-{32,64}-vbs-final.log` contain the new passing probes. The initial PE32 run
timed out during backend readiness; after restarting only its disposable prefix,
the missing libgcc dependency was exposed and removed. No Wine defect was
established.

Regression logs use `-{32,64}-*-regression.log` and `-{32,64}-suite-final.log`;
MSI logs use `-{mingw-32,winegcc-64}-msi-{build,check}.log`.

## Reproduction and limits

Compile the C test using the same generated header and IID object as the DLL:

```sh
# Fedora, from the shared checkout:
i686-w64-mingw32-gcc -Wall -Iunioffice_excel/build/mingw-32 \
  tests/range-enumerator.c unioffice_excel/build/mingw-32/unioffice_excel_i.o \
  -lole32 -loleaut32 -luuid -o unioffice_excel/build/mingw-32/range-enumerator.exe
# builder64: replace compiler with winegcc -b x86_64-windows -m64,
# and mingw-32 paths with winegcc-64.
```

Inside one Xvfb session, with the matching updated DLL already installed and
WINEPREFIX/WINE/CLIENT_BITS explicitly set for the disposable fixture:

```sh
sh tests/run-openoffice-backend.sh \
  unioffice_excel/build/mingw-32/range-enumerator.exe 'Range enumerator: 0 failure(s)'
sh tests/run-openoffice-backend.sh range-enumeration.vbs 'Range enumeration: PASS'
```

Coverage is limited to a single rectangular area. Native Windows Excel,
multi-area enumeration and structural sheet edits during an active enumeration
were not tested. Dimensions are captured when the cursor is created; cell
values remain live. General Range.Count retains its published LONG contract;
this work does not introduce CountLarge or fix its independent overflow limits.
