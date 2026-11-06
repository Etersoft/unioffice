# Documentation guide

## Start here

1. [Working notes](development-notes.md): current state, build, checks and
   limitations.
2. [Issue audit](issue-audit-2026-11-06.md): status decisions and remaining work.
3. [Roadmap](roadmap.md): the order of the next development stages.

## Setup, status and architecture

- [Supported functions and Windows/Wine installation](supported-usage.md).
- [Native Windows results, fixes and reproduction](native-windows-results.md)
  and [accepted package checksums](native-windows-packages.sha256).
- [Earlier installed-package Wine acceptance](installed-acceptance-results.md)
  and its [dated checksums](installed-packages.sha256).
- [OpenOffice 4.1.17 backend setup and original regression history](openoffice-4.1.17-results.md).
- [PE64 integration](pe64-results.md) and [MSI architecture/coexistence](msi64-results.md).
- [Linux COM/UNO bridge findings and proposed design](linux-office-bridge.md).
- [Related projects and public-source research](related-projects.md).

## Functional and COM regressions

These reports describe individual changes and their checks. Later native
Windows acceptance supersedes older statements that Windows had not yet been
tested; the reports retain their original environments and coverage limits.

| Area | Reports |
| --- | --- |
| Borders enumeration and indexing | [Independent cursors/COM contract](borders-enumerator-results.md), [eight keys and legacy aliases](borders-membership-results.md) |
| Border formatting | [Weight and UNO dimensions](border-weight-results.md), [mixed Null and native ABI](border-weight-mixed-results.md), [supported styles/error propagation](border-style-results.md) |
| Range arguments and ownership | [Preserved string arguments](range-argument-results.md), [cell inputs and owned results](range-cell-ownership-results.md), [retained worksheet parents](cells-parent-lifetime-results.md) |
| Cell and range indexing | [Range.Item and assignment](range-item-results.md), [Application/Worksheet.Cells dispatch](cells-indexing-results.md) |
| Row/column views | [Bounded Range collections](range-rows-columns-results.md), [Worksheet/Application collections](sheet-rows-columns-results.md) |
| Range enumeration | [Lazy traversal, independent cursors and COM validation](range-enumeration-results.md) |
| Application interface | [Published dual-interface ABI](application-interface-results.md) |
| Reopened function gaps | [Range, sheet, workbook, application, names and shapes members](function-gaps-results.md) |

The [working notes](development-notes.md) also preserve the detailed work log,
including superseded Winelib experiments. The current outputs are Windows PE
DLLs; Linux backend integration remains a separate open task.
