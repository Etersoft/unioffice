# Range.Item and indexed assignment

Range.Item now returns a cell Range and supports assignment to its value. The
same cell selection implements numeric default indexing and indexed Range.Cells.
The public IDL, IIDs and native vtable signatures remain unchanged. Item's LCID
was declared as an ordinary IDL argument; explicit IDispatch routing accepts the
one/two indices used by scripts and supplies the Invoke locale internally.

For cell ranges, supported positive indices include:

- Relative row/column numbers and numeric strings for Item.
- Case-insensitive column letters, including AA and later columns.
- One index in row-major order using the original range's width, continuing
  downwards beyond its last row.
- Cells outside the original range, within the backend sheet's bounds.
- Item/default/relative Cells assignment, with preserved caller variants.

For example, B2:D4.Item(2,"D") selects E3 and Item(10) selects B5. UNO rejects
out-of-bounds offsets on the original subrange, so UniOffice resolves absolute
sheet coordinates and creates the cell through an owned worksheet Cells range.
Arguments and assigned values use VariantCopyInd and common cleanup paths.

The existing BSTR default lookup used internally by Worksheet.Range is retained;
it is not claimed to implement every numeric-string default-indexing rule.
Rows/Columns collection semantics, multi-area ranges, nonpositive offsets and
native Excel equivalence remain outside this validation. Numeric coercion follows
the existing OLE conversion rules. See Microsoft's
[Range.Item contract](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item).

## TableBorder interaction

The expanded C test exposed nine later Weight failures after range operations.
On the same UNO target, PROPERTYGET TableBorder returned E_NOINTERFACE while
getPropertyValue("TableBorder") succeeded. A control without range operations
passed; removing only value assignments did not help. The direct VBScript UNO
control passed both routes, so this does not establish a Wine defect.

TableBorder reads/writes now use explicit UNO getPropertyValue/setPropertyValue,
including Color and Weight paths. The diagonal getter shares the same helper.
The regression checks Weight, Color and LineStyle after Item; the C test checks
exact HRESULTs, empty failed results, by-reference Double inputs, BSTR ownership,
overflow and argument counts. Diagnostic logs retain the original failure and
the route comparison under ~/tmp/claude/unioffice-range-item-32-native-*.log.

## Reproduction and results

Use Windows Apache OpenOffice 4.1.17 x86 and Wine@Etersoft 11.20.1-eter1 on
builder64. Build MinGW PE32 only on Fedora and winegcc PE64 only on builder64.
With explicit WINEPREFIX, WINE and CLIENT_BITS in a disposable fixture, inside
Xvfb:

```sh
sh tests/run-openoffice-backend.sh range-item.vbs 'Range Item: PASS'
sh tests/run-openoffice-backend.sh unioffice_excel/build/mingw-32/borders-enumerator.exe 'Borders enumerator: 0 failure(s)'
python3 scripts/run-excel-suite.py --bits 32
```

Use winegcc-64 and `--bits 64` for the PE64 counterparts.
Logs use ~/tmp/claude/unioffice-range-item-{32,64}-*-final.log.
The initial PE64 script failed with an argument-count error before implementation.

Both architectures passed 36 Item checks, the C COM contract test with zero
failures, and the existing worksheet-range, cell-index, parent-lifetime,
Application.Range, mixed Weight, Weight and style-error probes. All seven legacy
tests passed sequentially on each architecture: counters158/0,156 SUCCESS markers
(the existing test4/test6 counter discrepancies remain). Full logs are
~/tmp/claude/unioffice-excel-suite-clxbrb0m (PE32) and -f807h133 (PE64).
MSI32/MSI64 were rebuilt and passed architecture/Unicode metadata checks with
zero failures. Installer lifecycle was not repeated: registration and IDL did
not change. Native Windows was not tested.

An incremental Fedora build once retained an older i_range.o after a rapid NFS
edit. Source hashes matched, but the object predated the source. A targeted
`make -C unioffice_excel TOOLCHAIN=mingw MINGW=i686-w64-mingw32 -W i_range.c`
recompiled it. Check source/object times and deployed DLL bytes when a fix appears
absent; do not attribute a stale artifact to runtime behavior.
