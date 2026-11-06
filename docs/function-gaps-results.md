# Reopened function gaps: implementation and results

The [2026-11-06 audit](issue-audit-2026-11-06.md) reopened several stub-only or
incorrect Excel members. This stage implements the ones that have a clear
Excel contract and a verifiable OpenOffice/LibreOffice UNO counterpart. Every
check reads the resulting backend document, not just a successful HRESULT.
Members without a backend counterpart or with an unclear scope are listed
under [Not implemented](#not-implemented).

## Range

| Member | Issue | Behaviour |
| --- | --- | --- |
| `Hidden` | [#1851](https://bugs.etersoft.ru/show_bug.cgi?id=1851) | Row/column visibility (`IsVisible`), no longer `CellProtection.IsHidden`. Column views and full-height ranges address columns; other ranges address their rows. Reading a partly hidden set returns Null. |
| `CurrentRegion` | [#4257](https://bugs.etersoft.ru/show_bug.cgi?id=4257) | `createCursorByRange` + `collapseToCurrentRegion`. |
| `NumberFormat` | [#3148](https://bugs.etersoft.ru/show_bug.cgi?id=3148), [#4041](https://bugs.etersoft.ru/show_bug.cgi?id=4041) | English (en-US) format codes are looked up or added in the document's `NumberFormats`. The locale's standard format reads as `General`. |
| `FormulaR1C1` | [#2077](https://bugs.etersoft.ru/show_bug.cgi?id=2077) | `FormulaParser` with `AddressConvention.XL_R1C1`; the top-left cell's tokens are copied to the rest of the range, so relative references shift as in Excel. Reading prints the tokens back in R1C1. |
| `FormulaLocal` | [#4042](https://bugs.etersoft.ru/show_bug.cgi?id=4042) | The backend's localized cell formula (`FormulaLocal`), filled across the range in the same way. |

Excel R1C1 formulas use `,` between arguments and `,`/`;` inside array
constants, while the UNO API grammar uses `;` and `;`/`|`. The separators are
converted outside string literals and quoted sheet names. The Excel union
operator `,` therefore cannot be expressed in R1C1 text. Arrays assigned to
`FormulaR1C1`/`FormulaLocal` are rejected with `DISP_E_TYPEMISMATCH`.

## Worksheet, page setup and font

| Member | Issue | Behaviour |
| --- | --- | --- |
| `Worksheet.UsedRange` | [#4342](https://bugs.etersoft.ru/show_bug.cgi?id=4342), [#5812](https://bugs.etersoft.ru/show_bug.cgi?id=5812) | From the first to the last used cell (`gotoStartOfUsedArea`/`gotoEndOfUsedArea`); `Rows.Count`/`Columns.Count` follow. |
| `PageSetup.PrintArea` | [#3149](https://bugs.etersoft.ru/show_bug.cgi?id=3149) | The getter returns the sheet print areas as comma-separated `$A$1:$B$2` references. An empty string clears them. |
| `PageSetup.PrintTitleRows` | [#1874](https://bugs.etersoft.ru/show_bug.cgi?id=1874) | `setTitleRows`/`setPrintTitleRows`. Accepts `$1:$2`, `1:2`, `$3`, a cell reference such as `A3` and an optional `Sheet!` prefix. Reads back as `$1:$2`. An empty string disables title rows. |
| `Font.OutlineFont` | [#2512](https://bugs.etersoft.ru/show_bug.cgi?id=2512) | UNO `CharContoured`. |
| Full-width `Group`/`Ungroup`, `Worksheet.Copy`, `Outline.AutomaticStyles` | [#4471](https://bugs.etersoft.ru/show_bug.cgi?id=4471) | The sheet width comes from the backend (`getColumns().getCount()`), not from the `ooSetupVersion` first digit. AOO 4.x and LibreOffice versions matched neither `2` nor `3`, and LibreOffice sheets can have 16384 columns. |

## Workbook and Application

| Member | Issue | Behaviour |
| --- | --- | --- |
| `Workbook.Saved` | [#4470](https://bugs.etersoft.ru/show_bug.cgi?id=4470), [#4467](https://bugs.etersoft.ru/show_bug.cgi?id=4467) | The negation of the document's `isModified`; assignment calls `setModified`. |
| `Application.CentimetersToPoints` | [#3673](https://bugs.etersoft.ru/show_bug.cgi?id=3673) | `cm * 72 / 2.54`. |
| `Application.ScreenUpdating` | [#4028](https://bugs.etersoft.ru/show_bug.cgi?id=4028) | Works without an active workbook and applies to every open workbook and to workbooks added or opened later. Each document holds at most one `lockControllers`/`addActionLock` pair, so repeated `False` needs a single `True`. |
| `Application.Interactive` | [#4027](https://bugs.etersoft.ru/show_bug.cgi?id=4027) | Enables or disables the container windows of all current and later workbooks (`XWindow.setEnable`). |
| `Application` property | [#3578](https://bugs.etersoft.ru/show_bug.cgi?id=3578) | A regression checks it on Application, Workbooks, Workbook, Sheets, Worksheets, Worksheet, Range, Cells, Rows, Font, Interior, Borders, Border, PageSetup, Outline, Names, Name, Shapes and Shape. Shapes and Shape lacked it; `Application`/`Parent` are appended to their interfaces. |

`UsedRange`, `Saved`, `CentimetersToPoints`, `Interactive`, `Name.RefersTo` and
`Name.RefersToR1C1` declared their locale argument as `[in] LCID`, so late
binding demanded an extra argument. They now use `[lcid]`: the native vtable is
unchanged, and `ITypeInfo::Invoke` supplies the locale. The other 400 such
declarations were left as they were.

## Names and Name

[#2446](https://bugs.etersoft.ru/show_bug.cgi?id=2446): `Names.Add` takes `Name` or `NameLocal`, and
`RefersTo`, `RefersToLocal`, `RefersToR1C1` or `RefersToR1C1Local`. It converts
Excel syntax (`=Sheet1!$A$1:$B$2`) to UNO content (`$Sheet1.$A$1:$B$2`) and
replaces an existing name, as Excel does. `Names(n)` now treats `n` as a
one-based index; it used to pass it directly to the zero-based UNO
`getByIndex`.

[#2447](https://bugs.etersoft.ru/show_bug.cgi?id=2447): `Name` implements `Delete`, `Index`,
`RefersTo`, `RefersToR1C1`, `Value` and the default member (all of these equal
`RefersTo`), `NameLocal`, `RefersToLocal`/`RefersToR1C1Local` (the same English
syntax) and `Visible` (always True). `Name.Visible` previously shared the
`Value` DISPID. A missing `Name.Name` result used to release an uninitialized
pointer.

UNO named ranges have no hidden state, category, XLM macro type or shortcut
key. `Names.Add` and `Name` setters return `E_NOTIMPL` for these instead of
ignoring them. `Visible=False` is rejected rather than creating a visible name.

## Shapes

[#2094](https://bugs.etersoft.ru/show_bug.cgi?id=2094), [#4497](https://bugs.etersoft.ru/show_bug.cgi?id=4497):
`Shapes.AddLine` produced no visible line for three reasons:

- `i_shapes.c` called the line initializer without a prototype, so the C
  default promotions passed the `float` coordinates as `double`; the callee
  received zeros.
- Accessing `Shapes` called `DrawPages.insertNewByIndex`, which in Calc inserts
  a whole new sheet. Lines were then drawn on the first sheet's page.
- Coordinates were multiplied by 10 instead of being converted from points to
  1/100 mm (×2540/72).

Lines are now inserted into the worksheet's own draw page before their
geometry is set. A negative width or height preserves the line direction.

## Reproduction and results

Use the OpenOffice 4.1.17 fixtures described in
[PE64 integration](pe64-results.md) (builder64, `winegcc -b x86_64-windows`)
and the MinGW PE32 build on Fedora. Copy the DLL and TLB into the fixture's
`system32` (PE64) or `syswow64` (PE32) directory, then run inside Xvfb:

```sh
sh tests/run-openoffice-backend.sh range-properties.vbs 'Range properties: PASS'
sh tests/run-openoffice-backend.sh sheet-properties.vbs 'Sheet properties: PASS'
sh tests/run-openoffice-backend.sh application-ui-state.vbs 'Application UI state: PASS'
sh tests/run-openoffice-backend.sh names.vbs 'Names: PASS'
sh tests/run-openoffice-backend.sh shapes-line.vbs 'Shapes line: PASS'
sh tests/run-openoffice-backend.sh application-property.vbs 'Application property: PASS'
```

Both architectures (PE32 MinGW from Fedora, PE64 winegcc on builder64) passed
the six new probes, the 18 earlier VBScript regressions and the four rebuilt C
COM tests (application interface, Borders/Range enumerators, TypeInfo
lifetime): 28 of 28 each. The seven legacy tests passed sequentially on each
architecture with counters 158/0 and 156 SUCCESS markers, unchanged from the
previous baseline. Logs: `~/tmp/claude/uo-all-{32,64}.txt`,
`~/tmp/claude/unioffice-excel-suite-vcb8agjg` (PE32) and `-lfjghvaw` (PE64).

The new probes were not run against the original DLL; the earlier behaviour
above comes from the source code. The AddLine causes were confirmed with
temporary debug output. MSI packages were not rebuilt, and native Windows was not rerun at this stage.

## Not implemented

Closed as WONTFIX, as agreed for large or unclear tasks:

- `Outline.SummaryRow`/`SummaryColumn` ([#2345](https://bugs.etersoft.ru/show_bug.cgi?id=2345)): UNO
  sheet outlines have no summary-position setting to store or apply.
- `Application.WindowState` ([#3609](https://bugs.etersoft.ru/show_bug.cgi?id=3609)): the container
  window's `IsMaximized`/`IsMinimized` do not change in the headless fixture
  without a window manager, so the behaviour cannot be verified.
- `Workbook.PrintPreview` ([#3610](https://bugs.etersoft.ru/show_bug.cgi?id=3610)): Excel's preview is
  modal user interaction; nothing about it can be checked headlessly.
- `Workbook.Styles` ([#3674](https://bugs.etersoft.ru/show_bug.cgi?id=3674)): the requested scope
  ("getStyle") is unclear and implies a complete Styles/Style object model.
- `Range.PasteSpecial`/`Application.CutCopyMode` ([#2570](https://bugs.etersoft.ru/show_bug.cgi?id=2570)):
  these require a system clipboard and dispatcher model.
- Complete `Window`/`Windows` and `ActiveWindow` ([#2439](https://bugs.etersoft.ru/show_bug.cgi?id=2439)):
  `Application.Windows` still returns `E_NOTIMPL`; it is a whole object model.

`Range.Address` still declares `[in] LCID` and therefore fails under late binding.
`Range.Value` returns 0 for a formula with a text result. Both predate this
stage and are not part of the reopened issues.
