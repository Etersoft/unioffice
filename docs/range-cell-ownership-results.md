# Worksheet.Range and cell-index argument ownership

## Reproduced defects

The earlier Application.Range fix left other shallow-normalization sites to
audit. Two new probes failed against the preceding DLLs:

- PE64 Worksheet.Range(address): the address string changed after querying
  Borders.Count; the first preservation assertion failed.
- PE32 Application.Cells(row, column), with column="3": the first string check
  passed, then subsequent allocation exposed corruption at the second check.

Worksheet.Range normalized borrowed variants in place. Range's default member
also coerced shallow copies of index variants in place; converting a borrowed
BSTR column could free the caller's string. These are application defects.

## Changes

Both entry points now use VariantCopyInd into initialized owned locals, convert
only those copies, and release them on every exit. Failed index conversions
return their HRESULT before constructing a cell. Output pointers are checked
and initialized to null.

Worksheet.Range resolves each endpoint independently as a string or a queried
Range interface, holds owned references while reading coordinates, and releases
temporaries afterwards. String/object pairs work in either order. Numeric
endpoints and objects without I_Range fail rather than being blindly cast.

Application.Cells get/put and Worksheet.Cells put now propagate the default
member's failure and avoid dereferencing a missing cell. Temporary collections
and returned cells are released correctly when calls fail or results are ignored.
The separate [Cells parent lifetime fix](cells-parent-lifetime-results.md)
keeps exported collections' worksheet references alive (commit2ad7996).

The public IDL, IIDs and vtable signatures are unchanged. This stage does not
implement the separate Range.Item stub or new single-index/column-letter rules.
References: [VariantCopyInd](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantcopyind),
[Worksheet.Range parameters](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.range)
and [Range.Item indexing](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item).

## Coverage and reproduction

Tests run on builder64, Wine@Etersoft 11.20.1-eter1, Windows Apache OpenOffice
4.1.17 x86. PE32 was compiled with MinGW only on Fedora; PE64 with winegcc only
on builder64. Both passed:

- worksheet-range-ownership.vbs: preserved strings, one/two endpoints, usable
  caller-owned Range objects, mixed string/object endpoints, repeated cleanup,
  and rejection of invalid endpoints.
- cell-index-ownership.vbs: preserved numeric-string columns and Double rows,
  independent checks of absolute and relative target cells, Cells puts, invalid
  get/put indices with intact arguments and unchanged cell data.
- cells-parent-lifetime.vbs: parent and child lifetime after releasing callers.

The probes completed40,37 and49 checks respectively on each architecture.

The native borders-enumerator.c test also exercises Range/Worksheet null outputs,
typed conversion HRESULTs, null failed results and the preserved caller BSTR.
Its compile commands remain those in [enumerator results](borders-enumerator-results.md).

Within Xvfb, with WINEPREFIX/WINE/CLIENT_BITS selected for the disposable fixture:

```sh
sh tests/run-openoffice-backend.sh worksheet-range-ownership.vbs 'Worksheet Range ownership: PASS'
sh tests/run-openoffice-backend.sh cell-index-ownership.vbs 'Cell index ownership: PASS'
sh tests/run-openoffice-backend.sh cells-parent-lifetime.vbs 'Cells parent lifetime: PASS'
```

Logs are ~/tmp/claude/unioffice-range-ownership-{32,64}-*-final.log. Original
failures are unioffice-worksheet-range-ownership-64-red.log and
unioffice-cell-index-ownership-32-red.log. The intermediate PE64 Cells crash,
recorded in unioffice-range-ownership-64-cells-first.log, led to the parent fix.
All failures and passing reruns are retained. Native Windows was not tested.

The final DLLs passed all seven legacy tests sequentially on each architecture:
counters158/0 and156 SUCCESS markers, with the known test4/test6 counter
discrepancies. Full outputs: ~/tmp/claude/unioffice-excel-suite-kndhu5eq (PE32)
and unioffice-excel-suite-e3lqpp45 (PE64). The older Application.Range and mixed
Weight probes also passed. The expanded native contract test returned0 failures
on both architectures (unioffice-range-ownership-{32,64}-native-contract.log).

MSI32 and MSI64 were rebuilt from the final DLLs and passed architecture/Unicode
metadata checks with zero failures. Installer lifecycle was not repeated because
packaging/registration/IDL were unchanged. No Wine defect was established.
