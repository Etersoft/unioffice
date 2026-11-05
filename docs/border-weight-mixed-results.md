# Mixed Borders.Weight without changing the legacy vtable

## Automation and binary contracts

Automation Weight now returns VT_I4 for a uniform thickness and VT_NULL for
mixed thicknesses. Diagonals remain outside the aggregate, matching its setter.
The previous top-edge-only implementation failed tests/border-weight-mixed.vbs
when the left edge differed from the other edges.

The published I_Borders interface still has its original IID, method ordering
and XlBorderWeight-pointer getter. Native callers retain that getter's top-edge
behavior. Replacing its output with a VARIANT under the same IID would overwrite
old callers' enum-sized buffers; COM interfaces are immutable.

The type library instead adds a dispatch-only D_Borders contract, UUID
38CD5803-18E6-4F67-9525-110B4F31896B, with the same member names/DISPIDs and a
VARIANT Weight property. Borders supports that DIID and advertises its type
information through IDispatch. Weight property gets are handled explicitly;
other calls still use the original type information to invoke the unchanged
native vtable. TypeInfo outputs are AddRef'd; invalid indices/null outputs fail.

References: [COM interface immutability](https://learn.microsoft.com/en-us/windows/win32/com/interface-pointers-and-interfaces),
[dispatch type descriptions](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/type-descriptions)
and [Excel Borders.Weight](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.borders.weight?view=excel-pia).
The uniform/mixed policy is validated against the adapter's UNO backend; no
native Excel oracle was available.

## Geometry and mixed values

The aggregate compares the six non-diagonal line categories using the same
physical-width decoder as the individual getter. Missing interior lines are
excluded for single-cell, single-row and single-column ranges.

UNO TableBorder validity flags describe all line attributes. A false flag can
mean differing colors/styles rather than weights. Uniform lines use the fast
aggregate read; ambiguous lines are inspected cell by cell along the relevant
edge/interior segments. A difference in decoded weight yields Null. Equal
weights with different colors or single/double styles remain numeric.

The fallback uses range-relative cell positions, including ranges away from A1.
Its work scales with the number of ambiguous segments. Backend failures are
returned as errors rather than silently reported as mixed values.

## Tests and reproduction

Wine@Etersoft 11.20.1-eter1, Windows Apache OpenOffice 4.1.17 x86 on builder64.
MinGW PE32 was built only on Fedora; winegcc PE64 only on builder64.

tests/border-weight-mixed.vbs checks all four uniform constants, every differing
aggregate member, diagonal exclusion, mixtures within a side and internal lines,
equal weights with different colors/styles, restoration to uniform values, and
non-origin cell/row/column/rectangle ranges. It also checks preservation of the
Range address argument. That exposed a separate fixed input-ownership defect;
see [Range argument results](range-argument-results.md), commit9922be8.

The native borders-enumerator.c test verifies the dispatch-only TypeInfo and
VARIANT getter descriptor, uniform/mixed result tags, DIID support, and old
enum-buffer canaries. Both architectures pass. Existing Weight, enumeration,
diagonal and Range-argument probes also pass.

Build DLL and TLB together, stop only the disposable prefix's server, replace
both installed files, then register using that architecture's regsvr32. The
existing build commands remain valid. Within Xvfb with WINEPREFIX already set:

```sh
export WINE=/usr/lib/wine-etersoft/wine64 WINEDEBUG=-all CLIENT_BITS=64
sh tests/run-openoffice-backend.sh border-weight-mixed.vbs 'Mixed border weight integration: PASS'
sh tests/run-openoffice-backend.sh \
  unioffice_excel/build/winegcc-64/borders-enumerator.exe 'Borders enumerator: 0 failure(s)'
```

For PE32 use its fixture, CLIENT_BITS=32 and the mingw-32 executable. A new DLL
with an old TLB cannot provide the new dispatch metadata; ship the pair together.
Logs are ~/tmp/claude/unioffice-weight-mixed-{32,64}-*-final.log; the old failure
is unioffice-weight-mixed-red.log. These were application defects; no new Wine
issue was filed. Exact native Excel rendering/ordering remains outside coverage.

## Full regression and MSI lifecycle

All seven legacy Excel tests completed on PE32 and PE64: counters158/0 and
156 SUCCESS markers per architecture, including the known test4/test6 counter
discrepancies. Outputs: ~/tmp/claude/unioffice-excel-suite-w_qyjlm2 (32) and
unioffice-excel-suite-hjmdhjh4 (64). The mixed probe completed55 checks and the
physical Weight probe374 on each architecture; the native test returned0 failures.

The initial parallel PE32 run (unioffice-excel-suite-e85hajwi) stopped at test5:
cscript reported an execute-access page fault after workbook creation, before
Weight calls, and the runner reached its timeout. A separate tests5/6/7 run
(unioffice-excel-suite-wwimq370) passed, followed by the complete standalone
PE32 run above. The original failure remains recorded; its root cause was not
isolated and it is not classified as a Wine defect.

DLL and TLB were rebuilt together, packaged into MSI32/MSI64 and installed into
the separate unioffice-msi64/prefix fixture. Metadata/Unicode checks returned0
for both packages; the mixed probe passed against each installed architecture.
The coexistence runner then passed repair64, both removal orders with a usable
surviving client, and cleanup after the final removal. Both products are now
uninstalled from that fixture. Existing Automation fixtures retain the new pair.
Installer logs remain under unioffice-msi64/coexistence-logs and
unioffice-weight-mixed-install-*.log. Native Windows installation was not tested.
