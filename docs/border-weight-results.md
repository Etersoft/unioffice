# Border weight implementation and validation

## Behavior

Border.Weight previously returned E_NOTIMPL on reads and S_OK without changing
anything on writes. Borders.Weight also discarded child HRESULTs. Both are
implemented now using actual UNO line geometry, without per-object cached weights.
All twelve accepted indices, including diagonal lines and legacy aliases, work.

The four XlBorderWeight values map to nominal total widths in 1/100 mm:

| Weight | Enum value | Nominal width |
| --- | --- | --- |
| xlHairline | 1 | 2 |
| xlThin | 2 | 26 |
| xlMedium | -4138 | 88 |
| xlThick | 4 | 141 |

The numerical scale follows the widths used in
[LibreOffice's VBA border implementation](https://github.com/LibreOffice/core/blob/master/sc/source/ui/vba/vbaborders.cxx).
Here total width includes inner stroke, outer stroke and gap for double lines.
Double-line writes split it approximately into thirds. Reads use the nearest
category, tolerating UNO's rounding through twips. A line with zero dimensions
reports xlThin; it does not retain a hidden previous weight after clearing.

Weight preserves color and the existing single/double style. Setting Weight
on a cleared line makes a continuous line visible. LineStyle now shares the
geometry helpers: continuous/double transitions preserve thickness category,
and None clears dimensions. Dashed/dotted styles remain E_NOTIMPL.

UNO single lines use OuterLineWidth with zero InnerLineWidth/LineDistance.
The old code wrote the single stroke into the inner field and overwrote its
width with a fixed value whenever LineStyle changed. The shared implementation
fixes both issues and releases all temporary variants on failures.

For TableBorder writes, only the selected Is...LineValid flag is enabled;
other line and distance flags are disabled. Changing one edge therefore keeps
other edges' differing per-cell formatting. Diagonal structs continue to use
explicit getPropertyValue, followed by the corresponding property write.

Aggregate Weight updates the six non-diagonal members once each and propagates
errors. Its getter retains the existing top-edge behavior. The enum return type
in the legacy IDL does not support returning Null for mixed aggregate weights;
that contract was not changed here. Invalid weights return E_INVALIDARG before
modifying line geometry; null getter outputs return E_POINTER.

References: [Excel Weight](https://learn.microsoft.com/en-us/office/vba/api/excel.border.weight),
[weight constants](https://learn.microsoft.com/office/vba/api/Excel.XlBorderWeight),
[UNO BorderLine](https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1BorderLine.html)
and [TableBorder flags](https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1TableBorder.html).
The chosen geometry and style interactions were validated against OpenOffice,
not a native Excel runtime; exact Excel rendering equivalence is not established.

## Integration coverage

Wine@Etersoft 11.20.1-eter1 on builder64, Windows Apache OpenOffice 4.1.17 x86.
PE32 was built with MinGW on Fedora; PE64 with winegcc on builder64.
The former failing tests/border-weight.vbs is now a passing, expanded probe.
It completed 374 explicit checks on each architecture:

- Four weights on all twelve accepted indices, with ColorIndex and LineStyle.
- Direct UNO reads confirming visible single strokes and increasing widths.
- Double-line dimensions and retained categories through LineStyle changes.
- Clear/visible transitions, aggregate writes and untouched diagonals.
- Invalid individual/aggregate weights without changes to the old value.
- A mixed left edge on A1/A2 surviving a top-edge update on A1:D4.

tests/borders-enumerator.c also checks Weight and LineStyle null outputs and
invalid enum values through the native interfaces, alongside cursor/lifetime
regressions. Use the build commands in
[enumerator results](borders-enumerator-results.md).

Run within a disposable fixture and Xvfb, with WINEPREFIX already set:

```sh
export WINE=/usr/lib/wine-etersoft/wine64 WINEDEBUG=-all CLIENT_BITS=64
sh tests/run-openoffice-backend.sh border-weight.vbs 'Border weight integration: PASS'
sh tests/run-openoffice-backend.sh \
  unioffice_excel/build/winegcc-64/borders-enumerator.exe 'Borders enumerator: 0 failure(s)'
sh tests/run-openoffice-backend.sh borders-enumeration.vbs 'Borders enumeration integration: PASS'
sh tests/run-openoffice-backend.sh border-diagonal.vbs 'Border integration: PASS'
python3 scripts/run-excel-suite.py --bits 64
```

For PE32 select its fixture, CLIENT_BITS=32, mingw-32 executable and --bits 32.
Logs: ~/tmp/claude/unioffice-weight-{32,64}-first.log and
unioffice-weight-{32,64}-{native,collection,diagonal}.log. The original failing
logs remain unioffice-border-weight-{32,64}-known-failure.log.
This was an application defect; no Wine issue was filed.

## Regression and packaging results

All seven legacy tests completed with exit0 on both architectures:

| Test | Legacy successes/failures | SUCCESS/FAILED markers |
| --- | --- | --- |
| 1 | 32/0 | 32/0 |
| 2 | 13/0 | 13/0 |
| 3 | 22/0 | 22/0 |
| 4 | 34/0 | 33/0 |
| 5 | 32/0 | 32/0 |
| 6 | 12/0 | 11/0 |
| 7 | 13/0 | 13/0 |

Totals are 158/0 by the old counters and 156 SUCCESS markers; the known counter
discrepancies in tests4/6 remain. The earlier reduction from 166 to 158 came
from removing four duplicate Borders members, not from deleting assertions.
Full outputs: ~/tmp/claude/unioffice-excel-suite-mw8b7s1v (PE32) and
unioffice-excel-suite-75jv9aff (PE64).

Both MSI packages were rebuilt from these DLLs; architecture/Unicode metadata
checks returned zero failures for each. Installer lifecycle was not repeated:
installer and registration logic were unchanged. Native Windows and native
Excel were not exercised in this stage.
