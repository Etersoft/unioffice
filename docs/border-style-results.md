# Unsupported border styles and aggregate errors

## Installed backend capabilities

Windows Apache OpenOffice 4.1.17 x86 was inspected through both PE32 and PE64
script hosts on builder64, Wine@Etersoft 11.20.1-eter1. The standalone UNO probe
tests/openoffice-border-capabilities.vbs reports:

| API | Observed result |
| --- | --- |
| TableBorder, DiagonalTLBR | Available |
| TableBorder2, LeftBorder2, DiagonalTLBR2 | Unavailable |
| com.sun.star.table.BorderLine | Reflection type available |
| com.sun.star.table.BorderLine2, BorderLineStyle | Reflection types unavailable |
| LineStyle on the classic BorderLine struct | Error438 (1B6) |

The [Apache BorderLine API](https://www.openoffice.org/api/docs/common/ref/com/sun/star/table/BorderLine.html)
exposes color and inner/outer/gap dimensions. Its
[cell properties](https://www.openoffice.org/api/docs/common/ref/com/sun/star/table/CellProperties.html)
use that struct. The
[LibreOffice extended struct](https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1BorderLine2.html)
adds an explicit line style and is documented since LibreOffice3.4. Capabilities
must be checked against the installed backend rather than inferred from similar
version numbers or another product's SDK.

No dotted/dashed rendering support was added for the installed Apache backend.
The adapter continues to support continuous, double and None. The five other
[Excel line styles](https://learn.microsoft.com/en-us/office/vba/api/excel.xllinestyle)
(Dash, DashDot, DashDotDot, Dot, SlantDashDot) return E_NOTIMPL individually.
Supporting an extended backend would require a separate capability/mapping and
rendering validation stage; that backend has not been installed or tested here.

## Fixed application behavior

Borders.LineStyle previously ignored Item and child setter HRESULTs and returned
S_OK after every failed unsupported-style request. Borders.Value forwarded that
false success. The new regression failed on the old DLL at its first aggregate
Dash request after all individual requests correctly reported errors.

The aggregate setter now releases each child and returns its first failure.
Value inherits the fix through the existing alias. The getter checks its output
pointer and returns the actual child HRESULT. Native null getter outputs return
E_POINTER; invalid style12345 returns E_INVALIDARG; the five unsupported styles
return E_NOTIMPL. No IDL, IID, vtable or type library contract changed.

The tested unsupported/invalid requests fail before geometry changes, preserving
line style, Weight and ColorIndex. This does not add rollback for a later backend
failure after some members have already been updated.

## Validation and reproduction

PE32 was built with MinGW on Fedora; PE64 with winegcc on builder64. Both passed
the expanded native borders-enumerator.c contract checks, border-style-errors.vbs,
physical Weight (374 checks), mixed Weight (55 checks), and the previous diagonal
probe. The unchanged legacy test3 passed22/0, with22 SUCCESS markers on each.
Other six legacy tests were not repeated in this stage.

The new style probe completed213 explicit checks per architecture.

The style probe checks five unsupported constants and one invalid value across
all12 accepted individual indices, aggregate LineStyle and Value, preservation
of all8 distinct members, and supported double/continuous aliases. Both MSI
packages were rebuilt; architecture/Unicode checks returned zero failures each.
Installer lifecycle was not repeated because registration/packaging was unchanged.

Within Xvfb, with the usual disposable WINEPREFIX and CLIENT_BITS selected:

```sh
sh tests/run-openoffice-backend.sh openoffice-border-capabilities.vbs 'OpenOffice border capabilities: PASS'
sh tests/run-openoffice-backend.sh border-style-errors.vbs 'Border style errors: PASS'
python3 scripts/run-excel-suite.py --bits 64 --tests 3
```

Use --bits32 and CLIENT_BITS=32 for PE32. Compile the native regression test using
the commands in [enumerator results](borders-enumerator-results.md).
Logs: ~/tmp/claude/unioffice-style-errors-{32,64}-*.log; old failure:
unioffice-style-errors-red.log. Legacy outputs:
unioffice-excel-suite-6wh01dyo (32), unioffice-excel-suite-e0zfiu8l (64).
This was an adapter error-propagation defect and an observed backend API limit;
no Wine defect was established and no Wine issue was created. Native Windows
and exact Excel rendering were not tested.
