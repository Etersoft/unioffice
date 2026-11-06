# Supported UniOffice usage

The delivery target is the existing Excel Automation subset over Windows Apache
OpenOffice. Further development should fix failures in this subset and complete
its build/install/use/remove workflow. New COM methods require a concrete usage
requirement; an unimplemented interface member is not automatically a task.

## Verified functional scope

| Area | Existing coverage |
| --- | --- |
| Application and documents | Excel.Application activation, workbooks, worksheets, collection access/enumeration, sheet add/delete/rename, document close and application quit |
| Cells and ranges | String/numeric values, formulas and array scenarios in test5; Item/Cells addressing, selection, offsets, resize, row/column views and single-area enumeration |
| Formatting | Legacy font properties, subscript/superscript, solid/double/absent borders, colors, supported weights and mixed-weight results |
| Sheet layout | Group/ungroup and outline levels; page margins, orientation, zoom, centering and fit-to-page properties |
| COM contracts | Published Application interface, owned arguments/results, retained parents and independent enumerator cursors |
| Distribution | PE32/MinGW and PE64/winegcc MSI packages, registration, repair, architecture coexistence and removal |

The original seven scripts are the functional baseline. Targeted probes cover
previous defects and the added range/collection behavior. This table describes
tested scenarios, not every possible combination of these properties.

## Runtime and package selection

The tested runtimes are native Windows10 Pro22H2 x64 (Microsoft WSH) and
Wine@Etersoft11.20.1-eter1, with **Windows Apache OpenOffice4.1.17 x86** registered
as `com.sun.star.ServiceManager`. Complete OpenOffice's first-start setup in the
user profile before Automation. The Linux OpenOffice
installation alone does not supply this Windows COM backend. See the
[backend setup results](openoffice-4.1.17-results.md).
Linux office integration remains open; see the
[local bridge findings and proposed design](linux-office-bridge.md).

Choose the UniOffice package by the **client process** architecture:

| Client | Package | Test selector |
| --- | --- | --- |
| 32-bit | MSI/build/mingw-32/unioffice.msi | CLIENT_BITS=32; --bits 32 |
| 64-bit | MSI/build/winegcc-64/unioffice.msi | CLIENT_BITS=64; --bits 64 |

Both clients use the same tested x86 backend. Both packages can coexist.
Installation is machine-wide (ALLUSERS=1) and requires administrator privileges;
normal Automation runs without elevation. Before replacing an older per-user
package, remove it under the identity that installed it.
Both original suites also passed under a separate account with no administrator
group membership and its own fresh OpenOffice profile.
Registration is performed by MSI; manual DLL copying/registration is unnecessary.
Build and packaging commands are in the [README](../README.md).

## Native Windows: install and verify

Use the accepted MSI files listed in
[the package manifest](native-windows-packages.sha256). Copy `scripts/` and
`tests/` from the same revision to retain the functional checks. Choose the
package by client bitness, even though the backend is x86.

Install Windows Apache OpenOffice first and complete its first-start setup in
the account that will run Automation. From an elevated Command Prompt at the
repository root, install the selected UniOffice package. These commands install
both variants, waiting for each installer to finish:

```bat
start /wait msiexec.exe /i MSI\build\mingw-32\unioffice.msi /qn /norestart
start /wait msiexec.exe /i MSI\build\winegcc-64\unioffice.msi /qn /norestart
```

Then use ordinary **64-bit PowerShell** in a dedicated test account:

```powershell
# Both installed client architectures:
.\scripts\run-windows-tests.ps1 -Mode Suite
# If only the 32-bit package is installed:
.\scripts\run-windows-tests.ps1 -Mode Suite -Bits 32
```

Suite mode needs the repository scripts, not compiled C probes. It runs the
seven original tests, prints `Native Windows tests: PASS` only after all checks,
and saves reports under `~/tmp/claude/unioffice-windows-*`. Tests create documents
and quit OpenOffice; close other office work before running them. The full
probe and destructive MSI-maintenance commands are in the
[native validation report](native-windows-results.md).

## Wine: install and verify

From the repository root, after installing Windows OpenOffice into the dedicated
prefix, this installs and checks the 32-bit client in one graphical session:

```sh
export TMPDIR="$HOME/tmp/claude"
export WINEPREFIX="$HOME/tmp/claude/unioffice-runtime/prefix"
export WINE=/usr/lib/wine-etersoft/wine64 WINEDEBUG=-all CLIENT_BITS=32
xvfb-run -a sh -c '
set -eu
"$WINE" msiexec /i "Z:$PWD/MSI/build/mingw-32/unioffice.msi" /qn /norestart
sh tests/run-openoffice-backend.sh excel-backend.vbs "UniOffice integration: PASS"
python3 scripts/run-excel-suite.py --bits 32
'
```

For the 64-bit client, select winegcc-64, CLIENT_BITS=64 and --bits 64.
Tests create documents and call Quit; use a dedicated prefix. The runner prints
its result directory. Success requires all expected assertions, valid completion
counters, no failed markers and a zero process status. The current scripts emit
156 successful markers in total; their legacy counters total158 because tests4
and6 each count one extra success. Changing the assertions requires reviewing
the baseline in scripts/run-excel-suite.py.

Repair uses `/i <package> REINSTALL=ALL REINSTALLMODE=amusv`. Remove with
`msiexec /x <ProductCode> /qn`: the IDs and coexistence procedure are documented
in [MSI validation](msi64-results.md). The automated coexistence test removes
both packages and must run only in its disposable test prefix.

## Boundaries

Word is an unfinished component: its class factory returns E_NOTIMPL, and the
MSI packages contain Excel only. Dashed/dotted borders are unsupported by the
tested AOO backend and return an error. Multi-area ranges, structural edits
during enumeration, exhaustive full-row value/format/edit combinations and
other Windows versions have not been validated. Native Windows10 results and
current package checksums are in the [native report](native-windows-results.md).
The tested integration
workflow runs clients sequentially; a historical parallel PE32 test5 fault was
recorded without an isolated cause. No broader concurrency guarantee is made.
