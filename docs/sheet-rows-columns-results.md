# Worksheet and Application Rows/Columns

Worksheet.Rows/Columns and Application.Rows/Columns now return the corresponding
Range collection. With no index, Count reports the backend's row/column count.
Indexed calls share Range.Item selection and preserve the other dimension:
Rows(2) selects the second whole row; Columns("B") selects the second whole column.
Application resolves the active worksheet and propagates its getter's HRESULT.

The initial PE64 script stopped at Worksheet.Rows before executing assertions:
the published legacy Worksheet IDL requires an index. Dedicated Invoke handlers
now accept zero or one index without changing that IDL, IID or native signatures.
They initialize supplied results, allow ignored results, reject extra arguments
and return getter errors. Native getters retain their VARIANT index slot;
VT_EMPTY, VT_NULL and the omitted-argument marker request the collection.
Other VT_ERROR values are no longer silently treated as omitted indices.

The old independent parsers could use an uninitialized first boundary for a
single letter and ignored invalid characters. Variant normalization could also
modify caller-owned arguments. Native getters now copy/dereference indices,
obtain the Range collection and call its shared Item handler. Owned variants
and intermediate COM references are released on every exit.

## Indices and dimensions

The shared Range handler accepts positive numeric indices, case-insensitive
column letters and ascending slices such as "2:4" or "b:d". Slices also work
on bounded Range.Rows/Columns and remain relative to the source rectangle.
Empty endpoints, multiple colons, descending/nonpositive bounds, invalid text
and LONG overflow fail. Row/column Item still rejects a second argument;
Cells restores cell indexing. Ordinary Range named lookup is unchanged.

Dimensions come from the actual UNO sheet. In the tested AOO4.1.17 backend,
there are 1,048,576 rows and1024 columns. Old Columns constructors hardcoded
65,536 rows; Rows used version constants for width. EntireRow/EntireColumn
continue through the same native getters, now with the full backend dimensions.
Their address buffers were enlarged to hold two LONG-sized endpoints: the old
10-WCHAR row buffer was too small for "1048576:1048576".

This follows the collection relationship documented for
[Worksheet.Rows](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.rows),
[Worksheet.Columns](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.columns)
and [Application.Rows](https://learn.microsoft.com/en-us/office/vba/api/excel.application.rows).
The backend dimensions are not a claim of modern Excel sheet-size compatibility.

## Validation

tests/sheet-rows-columns.vbs passes177 checks on PE32 and PE64: collection Count,
native-backed direct/Item selection, preserved Double/BSTR indices, lowercase
letters, slices, invalid syntax, Cells conversion, parent lifetime, and the last
EntireRow/EntireColumn. The bounded Range test now passes98 checks, including
relative slices. The expanded C test returns zero failures on both: native
Application/Worksheet getter slots, NULL outputs, omitted indices,
coordinates/Count, BYREF Double/BSTR ownership, exact error HRESULTs, failed
NULL/VT_EMPTY results, ignored results, combined get/method flags and bad counts.

This native test also exposed the independent Application dual-vtable error,
fixed separately in b268269; see [its ABI report](application-interface-results.md).
The dedicated application-interface.exe returns zero failures on both builds.

PE32 was built with MinGW on Fedora; PE64 with winegcc on builder64. Runtime:
Wine@Etersoft11.20.1-eter1, Windows Apache OpenOffice4.1.17 x86, Xvfb and separate
disposable prefixes. Deployed DLLs matched build outputs byte-for-byte. Existing
Cells49/Item36/ownership/parent/Worksheet.Range/Application.Range probes passed.
All seven legacy tests passed sequentially on each architecture: counters158/0,
156 SUCCESS markers, with the known test4/test6 counter discrepancies. Outputs:

- PE32: ~/tmp/claude/unioffice-excel-suite-undy_3_x
- PE64: ~/tmp/claude/unioffice-excel-suite-yybcp7d7

Logs: ~/tmp/claude/unioffice-sheet-axes-{32,64}-*-final.log; initial script failure
is -64-red.log; the native pre-ABI failure is -32-native-before-qi.log. Both MSI
packages were rebuilt with final DLLs and passed architecture/Unicode metadata
checks (-{mingw-32,winegcc-64}-msi-final-check.log). Installer lifecycle was not
repeated because registration and IDL did not change. No Wine defect was established.

Single-area enumeration was added in a subsequent stage; see
[Range enumeration](range-enumeration-results.md).
Native Windows, multi-area ranges, exact Excel script error numbers,
reversed slices and exhaustive value/format/delete/insert behavior remain outside
this validation. Selected backend failures are propagated rather than translated
to Excel's exact error codes.
