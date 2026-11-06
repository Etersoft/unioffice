# UniOffice issue audit: 2026-11-06

## Scope and evidence

Reviewed the histories of all **112 UniOffice product issues**, including closed
issues, and **9 related issues in other products**. Referenced duplicate targets
were also inspected. Public issue rows are listed below; restricted issue IDs,
titles and comments are omitted.

The code reference is [9ac6620](https://gitlab.eterfund.ru/etersoft/unioffice/-/commit/9ac6620).
This was an issue-history and targeted source review, not a new execution of
every historical customer application. Existing runtime evidence is the
[accepted native Windows baseline](native-windows-results.md), the
[Wine acceptance](installed-acceptance-results.md) and the individual regression
reports in the [documentation guide](README.md).

The accepted baseline uses Windows Apache OpenOffice 4.1.17 with PE32/PE64 clients.
Native Linux LibreOffice is the development priority, with its bridge still
unfinished. Neither Word nor the full Excel API is covered by the passing legacy
suite. An error-free call to a setter can conceal a no-op implementation.

## Disposition rules

- Close completed work when its original scope has implementation or recorded
  acceptance evidence; distinguish historical tests from current results.
- Reopen work closed merely for age or postponed implementation. Narrow broad
  tickets to the identifiable unfinished portion.
- For old client failures without a current reproduction, require a targeted
  recheck; do not claim that the old failure still reproduces.
- Preserve duplicate closures, completed historical release/test records,
  invalid requests for another API and cancelled administrative work.
- Do not infer Vista/Windows 7 support from Windows 10 tests, or Linux UNO support
  from a successful Windows PE build.

Reopening a request records unfinished work; it does not authorize implementing
every stub immediately. Keep the accepted workflow stable and follow the
[roadmap](roadmap.md). Require a concrete client scenario before broadening old
compatibility requests.

## Verified tracker changes

Applied and reread **24 changes**: **18 reopened**, **3 closed FIXED**,
**1 RESOLVED FIXED** pending final acceptance, and **2 existing closed issues**
corrected from LATER to FIXED. These include the related Linux-integration issue.

Within the 112-issue UniOffice product, the snapshot changed as follows:

| State | Before | After |
| --- | ---: | ---: |
| CLOSED/FIXED | 57 | 56 |
| CLOSED/INVALID | 6 | 3 |
| CLOSED/DUPLICATE | 3 | 3 |
| CLOSED/LATER | 10 | 0 |
| CLOSED/MOVED | 1 | 1 |
| RESOLVED/FIXED | 0 | 1 |
| DEFERRED | 34 | 31 |
| NEW | 1 | 0 |
| REOPENED | 0 | 17 |

The product still has **48 unfinished issues** (31 DEFERRED and 17 REOPENED).
The separate Linux-integration issue is also REOPENED. Existing open requests
are not marked completed merely because the baseline suite passes.

| Issue | Final state | Evidence or remaining work |
| --- | --- | --- |
| [#731](https://bugs.etersoft.ru/show_bug.cgi?id=731) | REOPENED | Wine-to-native-Linux UNO bridge is absent; current winegcc PE output does not complete it. |
| [#1726](https://bugs.etersoft.ru/show_bug.cgi?id=1726) | REOPENED | Grand-Smeta failure was confirmed after initial closure; current client result is unknown. |
| [#1851](https://bugs.etersoft.ru/show_bug.cgi?id=1851) | REOPENED | Hidden changes CellProtection.IsHidden instead of row/column visibility. |
| [#1874](https://bugs.etersoft.ru/show_bug.cgi?id=1874) | REOPENED | PrintTitleRows getter returns an empty string and setter silently succeeds without work. |
| [#2077](https://bugs.etersoft.ru/show_bug.cgi?id=2077) | REOPENED | FormulaR1C1 delegates to Value; R1C1 formula semantics are absent. |
| [#2135](https://bugs.etersoft.ru/show_bug.cgi?id=2135) | REOPENED | Closure applied to another C++ version; current C lifetime/cleanup paths need targeted revalidation, including error paths. |
| [#2255](https://bugs.etersoft.ru/show_bug.cgi?id=2255) | REOPENED | No targeted save/reopen formatting confirmation; current reproduction is unknown. |
| [#2345](https://bugs.etersoft.ru/show_bug.cgi?id=2345) | REOPENED | SummaryRow/SummaryColumn setters do nothing; getters return constants. ShowLevels remains implemented. |
| [#2512](https://bugs.etersoft.ru/show_bug.cgi?id=2512) | REOPENED | OutlineFont always reads False and ignores writes; even the historical limited contract required retained values. |
| [#2558](https://bugs.etersoft.ru/show_bug.cgi?id=2558) | REOPENED | Requested persistence interfaces have no implementation/QI support; a concrete client case is still needed. |
| [#2679](https://bugs.etersoft.ru/show_bug.cgi?id=2679) | CLOSED FIXED | Subscript was actually implemented in 1624f69 and tested in the accepted PE32/PE64 baseline. |
| [#2680](https://bugs.etersoft.ru/show_bug.cgi?id=2680) | CLOSED FIXED | Superscript was actually implemented in 1624f69 and tested in the accepted PE32/PE64 baseline. |
| [#2684](https://bugs.etersoft.ru/show_bug.cgi?id=2684) | CLOSED FIXED | The specifically listed redundant pApplication fields are absent from current structs. |
| [#2716](https://bugs.etersoft.ru/show_bug.cgi?id=2716) | RESOLVED FIXED | Historical English NSIS acceptance and current English MSI language/dialog tables establish the requested installer. |
| [#2816](https://bugs.etersoft.ru/show_bug.cgi?id=2816) | REOPENED | Word activation remains a stub; the original WinDealer export was not rerun. |
| [#3149](https://bugs.etersoft.ru/show_bug.cgi?id=3149) | REOPENED | PrintArea setter calls setPrintAreas, but getter returns E_NOTIMPL. |
| [#3212](https://bugs.etersoft.ru/show_bug.cgi?id=3212) | REOPENED | SPDS GraphiCS export was closed administratively; current client result is unknown. |
| [#3245](https://bugs.etersoft.ru/show_bug.cgi?id=3245) | REOPENED | Long-term release/backlog work was closed for age; Linux and Word work remains. |
| [#3673](https://bugs.etersoft.ru/show_bug.cgi?id=3673) | REOPENED | Application.CentimetersToPoints still returns E_NOTIMPL. |
| [#3674](https://bugs.etersoft.ru/show_bug.cgi?id=3674) | REOPENED | Workbook.Styles getter returns E_NOTIMPL; original getStyle wording needs client-scope clarification. |
| [#4467](https://bugs.etersoft.ru/show_bug.cgi?id=4467) | REOPENED | Workbook.Saved remains unimplemented (#4470); full gldsw client was not rerun. |
| [#6600](https://bugs.etersoft.ru/show_bug.cgi?id=6600) | REOPENED | Usable export test material was never received; obtain an available reproducible scenario first. |
| [#6671](https://bugs.etersoft.ru/show_bug.cgi?id=6671) | CLOSED FIXED | Legacy VBScript tests have strict Wine and native Windows runners and recorded acceptance. |
| [#9806](https://bugs.etersoft.ru/show_bug.cgi?id=9806) | CLOSED FIXED | Current architecture-specific MSI deployment has native/Wine lifecycle acceptance; not a new Windows 7 claim. |

The font fixes refer to [1624f69](https://gitlab.eterfund.ru/etersoft/unioffice/-/commit/1624f69)
and [font-script.vbs](../tests/font-script.vbs). The three completed engineering
tasks refer to the [suite runner](../scripts/run-excel-suite.py),
[structure definitions](../unioffice_excel/mso_to_oo_private.h) and
[native MSI acceptance](native-windows-results.md). English installer evidence
includes the historical acceptance in #4043 and current
[language settings](../MSI/idt/Property.idt).

## Existing public issues kept open

These 30 product issues were already open. The reasons below explain why the
current baseline does not justify closing them. Their existing tracker fields
were preserved.

| Issue | Scope | Remaining evidence or work |
| --- | --- | --- |
| [#1905](https://bugs.etersoft.ru/show_bug.cgi?id=1905) | Работа в Vista и Windows 7 | Vista/Windows 7 acceptance remains separate; Windows 10 results do not prove it. |
| [#2094](https://bugs.etersoft.ru/show_bug.cgi?id=2094) | Не работает функция AddLine | AddLine exists, but the reported visual result has not been rechecked. |
| [#2439](https://bugs.etersoft.ru/show_bug.cgi?id=2439) | Реализовать IApplication->ActiveWindow | Complete Window/Windows behavior remains unfinished. |
| [#2446](https://bugs.etersoft.ru/show_bug.cgi?id=2446) | Реализация интерфейса Names | Names.Add still returns E_NOTIMPL; a working getter/enumerator is insufficient. |
| [#2447](https://bugs.etersoft.ru/show_bug.cgi?id=2447) | Реализация интерфейса Name | The Name interface still contains unimplemented operations. |
| [#2570](https://bugs.etersoft.ru/show_bug.cgi?id=2570) | Реализовать Copy (IRange), PasteSpecial (IRange), CutCopyMode (IApplication) интерфейса | Copy support does not complete PasteSpecial/CutCopyMode. |
| [#3131](https://bugs.etersoft.ru/show_bug.cgi?id=3131) | UniOffice не вызывает объекты Word | Word class factory returns E_NOTIMPL; no accepted Word workflow. |
| [#3148](https://bugs.etersoft.ru/show_bug.cgi?id=3148) | Реализация IRange->NumberFormat | NumberFormat remains unimplemented. |
| [#3207](https://bugs.etersoft.ru/show_bug.cgi?id=3207) | Добавить результаты тестов на wiki | Current published results do not complete the requested historical wiki/comparison matrix. |
| [#3208](https://bugs.etersoft.ru/show_bug.cgi?id=3208) | Добавить информацию на SourceForge и FireForge | Not all requested external project registrations/publications are established. |
| [#3561](https://bugs.etersoft.ru/show_bug.cgi?id=3561) | Выпуск версии UniOffice 0.7 | Publishing code is not evidence of the planned formal release. |
| [#3578](https://bugs.etersoft.ru/show_bug.cgi?id=3578) | Свойства "Application" | Application identity has not been checked across every requested interface. |
| [#3609](https://bugs.etersoft.ru/show_bug.cgi?id=3609) | Реализация Application->WindowState | WindowState remains unfinished. |
| [#3610](https://bugs.etersoft.ru/show_bug.cgi?id=3610) | реализовать Workbook->PrintPreview | Workbook.PrintPreview returns E_NOTIMPL. |
| [#3666](https://bugs.etersoft.ru/show_bug.cgi?id=3666) | Добавить IDL файл с описанием интерфейсов Word | Word IDL is a small Application-only skeleton, not a complete interface definition. |
| [#3713](https://bugs.etersoft.ru/show_bug.cgi?id=3713) | По умолчанию нет открытия файлов (vbs,xls) | Explicit cscript execution works; default VBS/XLS file associations are not established. |
| [#3874](https://bugs.etersoft.ru/show_bug.cgi?id=3874) | Изучение возможности использования OpenOffice в Host системе | Historical UNO research is useful; the host-office bridge remains absent. |
| [#4027](https://bugs.etersoft.ru/show_bug.cgi?id=4027) | реализовать I_ApplicationExcel_put_Interactive | Application.Interactive remains unimplemented. |
| [#4028](https://bugs.etersoft.ru/show_bug.cgi?id=4028) | исправить функцию screenupdating | ScreenUpdating still depends on an active workbook. |
| [#4041](https://bugs.etersoft.ru/show_bug.cgi?id=4041) | Реализовать NumberFormat | NumberFormat remains unimplemented; related to #3148. |
| [#4042](https://bugs.etersoft.ru/show_bug.cgi?id=4042) | Реализовать FormulaLocal | FormulaLocal remains unimplemented. |
| [#4257](https://bugs.etersoft.ru/show_bug.cgi?id=4257) | Свойство CurrentRegion интерфейса IRange | CurrentRegion remains unimplemented. |
| [#4342](https://bugs.etersoft.ru/show_bug.cgi?id=4342) | Реализовать UsedRange | UsedRange remains unimplemented. |
| [#4470](https://bugs.etersoft.ru/show_bug.cgi?id=4470) | Реализовать метод Saved ( Workbook ) | Workbook.Saved getters/setters return E_NOTIMPL. |
| [#4471](https://bugs.etersoft.ru/show_bug.cgi?id=4471) | Ошибка определения версии OOo | Historical office-version detection issue needs a targeted current-backend check. |
| [#4497](https://bugs.etersoft.ru/show_bug.cgi?id=4497) | Исправить работу интерфейса IShape | The requested visible-line behavior has no new acceptance result. |
| [#5812](https://bugs.etersoft.ru/show_bug.cgi?id=5812) | Не реализовано ActiveSheet.UsedRange.Rows.Count | UsedRange.Rows/Columns.Count is blocked by UsedRange. |
| [#5902](https://bugs.etersoft.ru/show_bug.cgi?id=5902) | 1с + unioffice + oocalc вылетает при работе с большими xls-файлами | The attached large-document crash case has not been rerun. |
| [#6357](https://bugs.etersoft.ru/show_bug.cgi?id=6357) | Выпуск версии UniOffice 0.6 | No evidence that the specifically planned formal release was completed. |
| [#6426](https://bugs.etersoft.ru/show_bug.cgi?id=6426) | Не работает перехват вызовов Excel в Само-Тур | The SamoTour export client has not been rerun. |

## Existing public closures retained

These 57 product issues retain their previous closed status. Historical
implementation/test records remain historical; a retained closure does not
claim that its original application or operating system was retested today.
For partially related work, the remaining operation has its own open issue.

| Issue | Resolution | Basis and limits |
| --- | --- | --- |
| [#675](https://bugs.etersoft.ru/show_bug.cgi?id=675) | FIXED | Historical client export fix/test recorded; that client was not rerun. |
| [#1251](https://bugs.etersoft.ru/show_bug.cgi?id=1251) | FIXED | MSI packaging exists; current lifecycle acceptance adds independent evidence. |
| [#1377](https://bugs.etersoft.ru/show_bug.cgi?id=1377) | FIXED | Historical installer fix; current MSI acceptance is documented separately. |
| [#1639](https://bugs.etersoft.ru/show_bug.cgi?id=1639) | FIXED | Historical release record; do not reinterpret it as a new release. |
| [#1847](https://bugs.etersoft.ru/show_bug.cgi?id=1847) | FIXED | Columns, MergeCells and ShrinkToFit have implementations and historical tests. |
| [#1850](https://bugs.etersoft.ru/show_bug.cgi?id=1850) | INVALID | Closed in favor of separate client bugs; not an assertion that every client works. |
| [#1873](https://bugs.etersoft.ru/show_bug.cgi?id=1873) | FIXED | Debug logging exists; historical task completed. |
| [#1878](https://bugs.etersoft.ru/show_bug.cgi?id=1878) | FIXED | Discussion/technology decision was completed; no blanket C++ rewrite claim. |
| [#1881](https://bugs.etersoft.ru/show_bug.cgi?id=1881) | FIXED | License notices were added; historical source-maintenance task. |
| [#1957](https://bugs.etersoft.ru/show_bug.cgi?id=1957) | DUPLICATE | Duplicate closure retained; no new client acceptance claim. |
| [#2076](https://bugs.etersoft.ru/show_bug.cgi?id=2076) | FIXED | Insert calls UNO insertCells; historical test recorded. |
| [#2078](https://bugs.etersoft.ru/show_bug.cgi?id=2078) | FIXED | Workbook.Name has an implementation and historical verification. |
| [#2079](https://bugs.etersoft.ru/show_bug.cgi?id=2079) | FIXED | Workbook.Names getter exists; unfinished Names methods remain in #2446/#2447. |
| [#2082](https://bugs.etersoft.ru/show_bug.cgi?id=2082) | FIXED | AutoFit uses UNO column sizing; historical test recorded. |
| [#2099](https://bugs.etersoft.ru/show_bug.cgi?id=2099) | FIXED | EntireRow/EntireColumn implemented; later collection fixes have separate acceptance. |
| [#2136](https://bugs.etersoft.ru/show_bug.cgi?id=2136) | FIXED | Excel IDL interface declarations exist; this does not mean every method is implemented. |
| [#2145](https://bugs.etersoft.ru/show_bug.cgi?id=2145) | FIXED | Historical early-binding test recorded; current native ABI probes add bounded evidence. |
| [#2253](https://bugs.etersoft.ru/show_bug.cgi?id=2253) | FIXED | Worksheet.Outline implemented; remaining Outline properties reopened in #2345. |
| [#2256](https://bugs.etersoft.ru/show_bug.cgi?id=2256) | FIXED | Current MSI removal/lifecycle acceptance also covers removal. |
| [#2257](https://bugs.etersoft.ru/show_bug.cgi?id=2257) | FIXED | Historical path-handling fix/test recorded. |
| [#2261](https://bugs.etersoft.ru/show_bug.cgi?id=2261) | FIXED | FontStyle implemented with historical tests; no universal style-name coverage claim. |
| [#2262](https://bugs.etersoft.ru/show_bug.cgi?id=2262) | FIXED | Historical test/fix task completed; existing suite now has strict execution checks. |
| [#2266](https://bugs.etersoft.ru/show_bug.cgi?id=2266) | FIXED | Historical release record. |
| [#2298](https://bugs.etersoft.ru/show_bug.cgi?id=2298) | INVALID | Request referenced a different API (SpreadsheetGear); not an established Excel requirement. |
| [#2309](https://bugs.etersoft.ru/show_bug.cgi?id=2309) | FIXED | SheetsInNewWorkbook is stored/read and used by workbook creation. |
| [#2315](https://bugs.etersoft.ru/show_bug.cgi?id=2315) | FIXED | Formula access exists; R1C1 and FormulaLocal remain separate unfinished tasks. |
| [#2343](https://bugs.etersoft.ru/show_bug.cgi?id=2343) | FIXED | Worksheet.Visible maps to UNO IsVisible; historical test recorded. |
| [#2451](https://bugs.etersoft.ru/show_bug.cgi?id=2451) | FIXED | Optional Worksheet.Range argument handling exists and has current argument tests. |
| [#2468](https://bugs.etersoft.ru/show_bug.cgi?id=2468) | FIXED | Offset implemented; historical test recorded. |
| [#2480](https://bugs.etersoft.ru/show_bug.cgi?id=2480) | FIXED | Rows/Columns implemented; current collection reports document fixes and limits. |
| [#2511](https://bugs.etersoft.ru/show_bug.cgi?id=2511) | FIXED | Group/Ungroup call UNO outline operations; historical test recorded. |
| [#2513](https://bugs.etersoft.ru/show_bug.cgi?id=2513) | FIXED | Worksheet.Select implemented; historical test recorded. |
| [#2518](https://bugs.etersoft.ru/show_bug.cgi?id=2518) | FIXED | Workbook.Save calls UNO Store; historical test recorded. |
| [#2557](https://bugs.etersoft.ru/show_bug.cgi?id=2557) | FIXED | Type library is generated and installed; native type-info checks passed. |
| [#2562](https://bugs.etersoft.ru/show_bug.cgi?id=2562) | FIXED | Resize implemented; historical test recorded. |
| [#2569](https://bugs.etersoft.ru/show_bug.cgi?id=2569) | FIXED | Type-info based dispatch exists; current ABI/lifetime probes add bounded coverage. |
| [#2595](https://bugs.etersoft.ru/show_bug.cgi?id=2595) | FIXED | Sheets enumeration exists; historical test recorded. |
| [#2600](https://bugs.etersoft.ru/show_bug.cgi?id=2600) | FIXED | Historical OpenOffice 3.0 test task; current 4.1.17 evidence is separate. |
| [#2616](https://bugs.etersoft.ru/show_bug.cgi?id=2616) | FIXED | Historical installer changes; current MSI acceptance supersedes deployment instructions. |
| [#2630](https://bugs.etersoft.ru/show_bug.cgi?id=2630) | FIXED | Historical release record. |
| [#2633](https://bugs.etersoft.ru/show_bug.cgi?id=2633) | FIXED | Version detection exists; the reported later-version problem remains in #4471. |
| [#2634](https://bugs.etersoft.ru/show_bug.cgi?id=2634) | FIXED | Worksheet.Range assignment implemented; current assignment probes add evidence. |
| [#2638](https://bugs.etersoft.ru/show_bug.cgi?id=2638) | FIXED | Application.Names forwards to the active workbook. |
| [#2671](https://bugs.etersoft.ru/show_bug.cgi?id=2671) | FIXED | Historical release-test record. |
| [#2677](https://bugs.etersoft.ru/show_bug.cgi?id=2677) | FIXED | Borders enumeration implemented and covered by current independent-cursor tests. |
| [#3150](https://bugs.etersoft.ru/show_bug.cgi?id=3150) | FIXED | Find has a UNO search implementation and a historical successful client test; no new exhaustive Find test claimed. |
| [#3151](https://bugs.etersoft.ru/show_bug.cgi?id=3151) | FIXED | Workbooks default/item access implemented; historical test recorded. |
| [#3153](https://bugs.etersoft.ru/show_bug.cgi?id=3153) | INVALID | History attributes the issue to DLL override configuration; current controlled COM activation works. |
| [#3579](https://bugs.etersoft.ru/show_bug.cgi?id=3579) | FIXED | Workbooks enumeration implemented; historical test recorded. |
| [#3888](https://bugs.etersoft.ru/show_bug.cgi?id=3888) | FIXED | Historical IDL work is recorded as completed; no new complete Excel ABI audit claimed. |
| [#3977](https://bugs.etersoft.ru/show_bug.cgi?id=3977) | FIXED | Published dual-interface early binding has current native ABI probes. |
| [#4043](https://bugs.etersoft.ru/show_bug.cgi?id=4043) | FIXED | Current MSI installation, repair, coexistence and removal have explicit acceptance. |
| [#4103](https://bugs.etersoft.ru/show_bug.cgi?id=4103) | DUPLICATE | Duplicate of #5870, whose Wine-side fix was recorded and verified. |
| [#6358](https://bugs.etersoft.ru/show_bug.cgi?id=6358) | FIXED | MSI build restored; current accepted packages have recorded hashes. |
| [#8253](https://bugs.etersoft.ru/show_bug.cgi?id=8253) | FIXED | Historical installer testing completed; current MSI lifecycle has separate results. |
| [#9805](https://bugs.etersoft.ru/show_bug.cgi?id=9805) | MOVED | Installation defect moved to #9806; do not turn it into a new Windows 7 acceptance claim. |
| [#17446](https://bugs.etersoft.ru/show_bug.cgi?id=17446) | FIXED | Current controlled Excel.Application activation tests substantiate the narrow COM-creation goal; no customer-client acceptance inferred. |

## Follow-up and reproduction

The tables cover all **110 public product issues** and the **one public related
issue** in the reviewed set. Restricted issues were reviewed without publishing
their identifiers or contents here. Duplicate closures were preserved rather
than creating competing technical tasks.

Each changed issue contains its own audit comment and completion criteria.
Prioritize the Linux LibreOffice vertical slice described in the
[bridge design](linux-office-bridge.md) and [roadmap](roadmap.md), while preserving
the accepted Excel packages. Old client cases need an accessible application,
a minimal input and a clear expected output before implementation changes.

For property gaps, test the resulting document and readback, not just S_OK:
row/column visibility for Hidden, readback for PrintArea and OutlineFont, and
computed cell values for R1C1 formulas. The
[Excel Hidden contract](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden)
describes row/column visibility; CellProtection.IsHidden is a different operation.
No newly discovered method was implemented during this audit, and accepted MSI
artifacts were not rebuilt. No unconfirmed Wine defect was filed.
