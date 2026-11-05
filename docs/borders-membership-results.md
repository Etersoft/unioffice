# Borders membership and index validation

## Contract and compatibility decision

[Microsoft XlBordersIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.xlbordersindex)
defines eight distinct range borders, with keys 5 through 12. The
[Borders object documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.borders)
also describes four borders, while listing all eight keys. It distinguishes
Range/Style indices from FormatCondition indices. These pages alone do not
establish exact Excel enumeration order or runtime Count behavior.

[LibreOffice's VBA implementation](https://github.com/LibreOffice/core/blob/master/sc/source/ui/vba/vbaborders.cxx)
uses eight supported keys and derives Count from that table; its source also
notes that matching Excel enumeration order remains unfinished. This supports
using eight distinct members here, without claiming a native Excel comparison.
The source was inspected on 2026-11-06. No native Excel/Windows oracle was
available in this environment.

UniOffice formerly reported Count=12 and enumerated keys 1 through 12. Keys
1/2/3/4 were duplicates of 7/10/8/9 (left/right/top/bottom). It also accepted
any other integer and returned a Border object that had no supported line key.

Count and enumeration now use a single table of eight distinct keys. The four
legacy aliases remain accepted by Item and the default property for existing
clients, but are not separately enumerated. Current order is ascending keys
5 through 12; callers should use named indices when a specific border matters.
This is a deliberate compatibility change from the previous 12-item traversal.
It is not a claim of full Excel object-model conformance.

Item/default reject keys outside 1 through 12 with DISP_E_BADINDEX and a null
output, before allocating an object. Null output arguments return E_POINTER.
Count likewise checks its output pointer. Aggregate Color/LineStyle setters
visit the six distinct non-diagonal keys, preserving diagonal formatting.

## Validation

Wine@Etersoft 11.20.1-eter1 on builder64, Windows Apache OpenOffice 4.1.17 x86.
MinGW PE32 was built on Fedora; winegcc PE64 was built on builder64.

The expanded native tests/borders-enumerator.c reported 21 failures on the
previous DLL: Count plus HRESULT/output assertions for five invalid keys
through both Item and default. After the fix it reports zero failures on
both architectures, including the existing cursor/Clone/lifetime cases.

tests/borders-enumeration.vbs assigns a distinct ColorIndex to each named key,
requires each to be returned once, checks the four alias mappings, tests
aggregate Color/LineStyle and diagonal preservation, and rejects invalid keys.
Nested loops visit 8x8 members; repeated traversal visits eight again.
The existing border-diagonal.vbs still exercises all twelve accepted keys.

Both scripts passed on PE32 and PE64. The unchanged legacy test_3 now reports
22 successes, zero failures and 22 SUCCESS markers for each architecture.
Its previous 30 comparisons included two checks for each of the four duplicate
members; the new distinct-member test checks their actual replacement coverage.
Other legacy tests were not repeated. Suite outputs:
~/tmp/claude/unioffice-excel-suite-8y0zs2lr (PE32) and
unioffice-excel-suite-l7kaqotw (PE64).

Both MSI packages were rebuilt from the updated DLLs; architecture/Unicode
metadata checks returned zero failures for each. Installer lifecycle was not
repeated because installer/registration logic was unchanged.

## Known separate defect: Weight

An added aggregate-formatting check failed on both architectures when reading
Border.Weight. Inspection confirms that its getter returns E_NOTIMPL and its
setter returns S_OK without changing the backend. This predates membership
changes. It is retained as the explicit failing probe tests/border-weight.vbs,
outside the passing membership test; Weight support was not implemented here.

Run it with the usual ready-backend harness:

```sh
sh tests/run-openoffice-backend.sh border-weight.vbs 'Border weight integration: PASS'
```

Expected current result: failure when reading Weight, harness exit1 and no PASS.
Do not treat the successful put as evidence of working formatting. The next
implementation should validate actual UNO line widths, individual and aggregate
Weight, all supported weight constants, and interactions with line style.

## Reproduction and logs

Use the compile commands and disposable fixture setup in
[enumerator results](borders-enumerator-results.md); native executable and
script names remain the same. Set CLIENT_BITS and --bits to the tested bitness.

Logs are under ~/tmp/claude on builder64:

- unioffice-border-membership-red.log: old native implementation.
- unioffice-border-membership-{32,64}-{native,collection,diagonal}.log.
- unioffice-border-weight-{32,64}-discovery.log: initial aggregate failure.
- unioffice-border-weight-{32,64}-known-failure.log: dedicated Weight probe.

The failures described here are UniOffice defects; no new Wine issue was filed.
