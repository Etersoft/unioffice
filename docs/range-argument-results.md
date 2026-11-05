# Preserve Application.Range input arguments

The mixed Weight probe exposed PE64 string corruption: after obtaining
Excel.Range(address).Borders and reading a property, the caller's address
variable changed to unrelated text. A dedicated Range/Borders.Count probe
reproduced this without calling Weight. Merely obtaining the Range object and
printing its address, and a standalone nested-array VBScript control, passed.

Application.get_Range normalized its by-value VARIANT parameters in place with
MSO_TO_OO_CorrectArg. That helper dereferences arguments using shallow copies.
The public entry now uses VariantCopyInd into owned local variants, forwards
those values and clears them afterwards. It also checks the output pointer and
ActiveSheet errors. Its published native signature is unchanged.

On builder64 with Wine@Etersoft 11.20.1-eter1 and Windows OpenOffice 4.1.17 x86,
the PE64 address-preservation probe failed before this change and passed after
it. The same final probe passed on MinGW PE32. The mixed Weight test likewise
passed its explicit address checks on both architectures.

The diagnostic changes to scalar result assignment and temporary range-address
VARIANT initialization did not eliminate the corruption; they were not evidence
of its cause. The temporary initialization change was discarded. No Wine defect
was established, so no Wine issue was created.

Regression: tests/range-address-variable.vbs, run in a disposable OpenOffice
fixture with the usual CLIENT_BITS/WINE/WINEPREFIX settings inside Xvfb:

```sh
sh tests/run-openoffice-backend.sh range-address-variable.vbs 'Range address lifetime: PASS'
```

It retains the input variable while querying Borders.Count, allocates/prints
strings repeatedly and checks that the argument remains unchanged. Use the
matching updated DLL for each architecture.

Logs under ~/tmp/claude on builder64:

- unioffice-range-address-variable-pe64-diagnostic.log: failing Range/Count case.
- unioffice-range-address-variable-pe64-owned-args.log: fixed PE64 case.
- unioffice-weight-mixed-{32,64}-range-final.log: final regression.
- unioffice-nested-arrays-{system32,syswow64}.log: passing standalone controls.

The control source is retained as ~/tmp/claude/unioffice-nested-array-enumeration.vbs.
This fix covers Application.Range; similar normalization patterns elsewhere
require separate evidence and tests before changing their ownership contracts.
