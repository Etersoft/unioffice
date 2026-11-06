# Application and Worksheet Cells indexing

Application.Cells and Worksheet.Cells now forward their Automation calls to the
returned Range's Cells handler. This shares the previously tested Item selection,
value assignment, argument ownership and error propagation. The native getters,
IDL, IIDs and vtable signatures remain unchanged.

The original PE32 and PE64 probes each reported nine failures in 39 executed
checks. Application.Cells(4) failed, both owners rejected Cells(index) assignment,
and Cells("2","AA") assignment reported success while changing the wrong target.
Invalid string indices could also report success. The separate handlers used
VT_EMPTY for an omitted column, required two indices for assignment, and routed
string row indices through the legacy named-range default lookup.

The shared handler accepts zero-argument collection retrieval and one/two-index
cell retrieval or value assignment. Numeric strings follow Item coercion; column
letters are supported. Nonnumeric row indices fail conversion. Address/name
lookup remains available through Range; it is not the Cells selection rule.
Calls without a result pointer use a temporary VARIANT that is always cleared.
The wrappers initialize supplied results and return the child's HRESULT.

This follows the documented relationship between
[Application.Cells](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cells),
[Worksheet.Cells](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.cells)
and [Range.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.range.item).
Full native Excel equivalence and Excel sheet dimensions are not established;
the existing backend dimensions and positive-index limits still apply.

## Tests and reproduction

The new cells-indexing.vbs completes 49 checks on each architecture: direct
one/two-index get/put, independent A1-address value checks, preserved Double/BSTR
arguments and assigned strings, conversion/argument-count errors, collection
retrieval, and parent/child lifetime after collection release. Valid getters must
return a non-Nothing object as well as a zero error number.

The expanded C COM test verifies exact HRESULTs, VT_EMPTY failed results,
by-reference Double indices, caller BSTRs, ignored results for collection/cell
gets, combined METHOD|PROPERTYGET flags, a NULL DISPPARAMS pointer, and invalid
get/put argument counts. Both architectures returned zero failures.

Build PE32 with MinGW on Fedora and PE64 with winegcc on builder64. Tests use
builder64's Wine@Etersoft 11.20.1-eter1 and Windows Apache OpenOffice 4.1.17 x86.
Inside Xvfb, with explicit WINEPREFIX/WINE/CLIENT_BITS for a disposable fixture:

```sh
sh tests/run-openoffice-backend.sh cells-indexing.vbs 'Cells indexing: PASS'
sh tests/run-openoffice-backend.sh unioffice_excel/build/mingw-32/borders-enumerator.exe 'Borders enumerator: 0 failure(s)'
python3 scripts/run-excel-suite.py --bits 32
```

Use winegcc-64 and `--bits 64` for PE64. New probe logs are
~/tmp/claude/unioffice-cells-indexing-{32,64}-{indexing,native}-first.log;
original failures are retained as -red.log. No Wine defect was established.

On both final DLLs, the existing Item, cell-index ownership, Cells parent lifetime,
Worksheet.Range and Application.Range probes passed. All seven legacy tests
passed sequentially on each architecture: 158/0 counters and156 SUCCESS markers,
with the existing test4/test6 counter discrepancies. Full outputs are
~/tmp/claude/unioffice-excel-suite-0lyrmghp (PE32) and -_ucysqo5 (PE64).
Deployed DLLs matched build outputs byte-for-byte. MSI32/MSI64 were rebuilt and
passed architecture/Unicode metadata checks with zero failures. Installer lifecycle
was not repeated because registration and IDL did not change. Native Windows was
not tested.
