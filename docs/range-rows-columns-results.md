# Range Rows and Columns collections

Range.Rows and Range.Columns now retain the original rectangle instead of
expanding it to entire worksheet rows or columns. Each returns a separate view
with a private collection kind. Count reports the number of rows or columns;
Item and indexed/default access select a subrange, retaining its other dimension.
For B2:D4, Rows(2) selects B3:D3, Columns(2) selects C2:C4, and Columns("AA")
selects AB2:AB4. Positive indices can extend beyond the original rectangle.
Rows accepts numeric indices, including numeric strings; Columns also accepts
case-insensitive letters. A second Item argument is rejected with E_INVALIDARG.

Cells creates a cell collection view, restoring linear/two-index cell selection
and cell Count without changing the original collection. The view owns its UNO
range and worksheet references, so selected ranges survive release of their
source objects. Group/Ungroup use the collection kind to choose the axis;
the old width heuristic remains for ordinary cell ranges.

The selection rules follow Microsoft's
[Range.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item) and
[Range.Columns](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columns)
documentation. LibreOffice's independent
[VBA Range implementation](https://github.com/LibreOffice/core/blob/master/sc/source/ui/vba/vbarange.cxx)
also retains row/column collection state through Item and changes Count by kind.
This comparison is supporting evidence, not a native Excel runtime oracle.
IDL, IIDs and public vtables are unchanged. The added private RangeImpl field
requires rebuilding all dependent objects; both builds forced the private header.

## Validation

The original PE64 DLL failed 49 of 71 executed VBS checks. The final
tests/range-rows-columns.vbs passes 86 checks on each architecture: bounds,
Count, Item/default/direct selection, outside offsets, string ownership,
rejected second indices, Cells conversion, unchanged source and parent lifetime.
The C COM test checks native getters, NULL outputs, exact HRESULTs and VT_EMPTY
failures, by-reference Double indices, numeric/letter conversion errors,
ignored results and native Count/coordinate slots. Both return zero failures.

PE32 was built with direct MinGW on Fedora; PE64 with winegcc on builder64.
Runtime was Wine@Etersoft 11.20.1-eter1 with Windows Apache OpenOffice 4.1.17 x86,
using separate disposable prefixes and matching client bitness. Deployed DLLs
matched build outputs. Existing Cells indexing49, Item36, cell ownership37,
parent lifetime49, Worksheet.Range40 and Application.Range probes passed both.
All seven legacy tests passed sequentially per architecture: counters158/0,
156 SUCCESS markers, including Group/Ungroup calls. The known test4/test6
counter discrepancies remain. Full outputs:

- PE32: ~/tmp/claude/unioffice-excel-suite-x7x89q44
- PE64: ~/tmp/claude/unioffice-excel-suite-bqcaa4i6

Probe logs are ~/tmp/claude/unioffice-rows-columns-{32,64}-*-final.log;
the original failure is -64-red.log. MSI32/MSI64 were rebuilt and passed metadata
and Unicode checks. Installer lifecycle was not repeated: registration and IDL
did not change. No Wine defect was established; native Windows was not tested.

## Scope

This change covers single-rectangle Range.Rows/Columns selection and Count.
The [subsequent sheet collection change](sheet-rows-columns-results.md) adds
Worksheet/Application.Rows/Columns and ascending slices such as "1:3"/"A:C";
the bounded Range probe now has98 checks. The validation above records the
initial collection stage. Multi-area selection, enumeration, nonpositive offsets
and exhaustive row/column value/format/delete/insert behavior remain outside
validation. Excel's script error number1004 is not reproduced by E_INVALIDARG.
Existing named lookup for ordinary Range default remains.
