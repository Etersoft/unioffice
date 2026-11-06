# Native Windows 10 validation

Validated on 2026-11-06: the existing Excel subset and both MSI packages passed
native Windows tests after two fixes. No COM functionality was added.

## Environment and final state

Dedicated Proxmox VM 744, lav-unioffice-win10, node gefest, instance Etersoft;
IP 192.168.8.248. It is a linked clone of Windows template 425; original VMs were
not modified. Windows 10 Pro 22H2 x64, build 19045.6456; Microsoft cscript
5.812.10240.16384 in both SysWOW64 and System32. Backend: Windows Apache
OpenOffice 4.1.17 x86. Tests ran in the interactive guest account, including
both complete legacy suites with a limited token (Administrator=False).

The temporary automatic-login password/settings were removed after initial
login. OpenOffice's first-start wizard was completed in this profile. Both final
UniOffice packages were reinstalled after the successful removal test so the VM
is ready for manual review. Final cell round trips passed for both client
architectures without elevation. The temporary artifact HTTP server was stopped.

## Fixed application crash: borrowed ITypeInfo

The original PE32 MSI installed successfully, but excel-backend.vbs crashed
with 0xc0000005 after printing Excel.Application: OK. Native COM probes and a
direct OpenOffice cell round trip passed. tests/typeinfo-lifetime.c reproduced
the crash independently: obtain two GetTypeInfo results, release one, then use
the other. Its original log stops at TypeInfo lifetime: Application.

Seventeen public getters returned a borrowed cached ITypeInfo without AddRef.
Microsoft's script engine released its result and left the cache dangling.
Commit 4c951d8 gives each caller its own reference and validates output/index;
internal cache users still borrow the retained reference. Borders already
followed this ownership rule and was preserved. No IDL or vtable layout changed.
The C probe now passes on Windows and Wine for both architectures; it exercises
repeated references across Application, workbook/sheet collections, ranges,
fonts, interiors, borders, PageSetup and Outline.

## Fixed MSI installation context

Original packages wrote system DLLs and machine COM registration but omitted
ALLUSERS. Installation as SYSTEM succeeded; a desktop user then saw product
state 2 (installed for another user), and even elevated repair returned 1605.
This was an installer defect, not an Automation or Wine failure.

Commit 0559f6b stages ALLUSERS=1 for both packages and tests that property in
msi-smoke.c. After removing the old per-user products under SYSTEM and installing
the fixed packages, the limited desktop user sees both products in state 5.
Native repair by that user's elevated token succeeds for both architectures.
See Microsoft's [ALLUSERS semantics](https://learn.microsoft.com/en-us/windows/win32/msi/allusers)
and [product states](https://learn.microsoft.com/en-us/windows/win32/api/msi/nf-msi-msiqueryproductstatew).
Older per-user installs must be removed under their original installing identity
before changing installation context; automatic context migration is not added.

## Results

| Check | PE32/MinGW | PE64/winegcc |
| --- | --- | --- |
| Initial installation, then repair with ITypeInfo fix | PASS | PASS |
| DLL/factory/TLB, Application ABI, type-info lifetime, enumerators | 0 failures | 0 failures |
| All 18 existing VBS integration probes | PASS | PASS |
| All 7 legacy scripts, elevated and limited tokens | PASS | PASS |
| Separate non-administrator account and fresh profile | PASS | PASS |
| Final MSI metadata/Unicode and machine context | PASS | PASS |
| Repair with both clients retained | PASS | PASS |
| Both removal orders and surviving client | PASS | PASS |
| Final removal of files, product records and shared ProgID | PASS | PASS |
| Reinstallation for review, limited-user cell round trip | PASS | PASS |

Legacy results per architecture: counters 158/0, 156 successful assertion markers,
no failed markers. Tests 4/6 retain their documented extra counter increments.
Native DLL/TLB hashes matched Linux build outputs. After the MSI-only scope
change, payload hashes were unchanged and both surviving-client workflows passed.
Both final packages also passed installation/repair/removal under Wine; both
strict legacy suites and the new C probe passed there after the code fix.

## Separate standard-user check

A subsequent run on the same VM tested a genuinely separate local account,
unioffice_std (SID ending 1002), rather than the administrator's filtered token.
System-side group enumeration confirmed that it did not belong to Administrators;
its process also reported Administrator=False. Its USERPROFILE and ApplicationData
known folders pointed to C:\Users\unioffice_std, independently of the guest account.
Both unchanged machine-installed MSI products returned state 5 for this identity.

The existing runner executed all seven original scripts for PE32 and PE64 in
this fresh profile. Both suites passed: counters 158/0, 156 successful markers,
no failed markers. No DLL copying, COM re-registration, package rebuild or
additional privileges were used. The existing headless startup initialized this
user's OpenOffice profile. Results: standard-user/results/runs/
unioffice-windows-47d97e58ab204e14bbb88e42a83406fd under the guest evidence base.

The temporary account was disabled after the run and its encrypted credential
file deleted; the existing guest session and installed packages were retained.
Verified logs on builder64: ~/tmp/claude/unioffice-standard-user-validation/results.zip,
SHA256 d0dbde2a3d6cbc0bcdb62632e1a62557f408f77aab1e83b984e7ae42b6daa8e3.
The archive includes identity.json, account-cleanup.json, suite.log, complete
reports and a preserved launcher setup error that occurred before the suites.
No new application or MSI defect was found in this check.

## Reproduction

Build on the established hosts: MinGW on Fedora, winegcc on builder64. Copy the
source tests/scripts and matching native probe EXEs to the Windows checkout.
With both MSIs and the backend installed, run in 64-bit PowerShell:

```powershell
# Elevated interactive session for registration probes:
.\scripts\run-windows-tests.ps1
# Ordinary interactive session; no registry-writing probes:
.\scripts\run-windows-tests.ps1 -Mode Suite
# Elevated, disposable VM only; removes both installed products:
.\scripts\test-windows-msi.ps1 -Msi32 C:\tests\unioffice32.msi -Msi64 C:\tests\unioffice64.msi
```

The functional runner preserves CP1251 assertions, substitutes only completion
dialogs, selects native 32/64-bit cscript and enforces complete counters/markers.
It checks installed registration before a probe can rewrite it. The legacy
basename DLL registration was accepted after confirming native activation.
The MSI runner verifies both desktop-visible product states, repair from the
user context, unchanged payloads, survivor activation and final cleanup.

Final accepted packages: MSI/build/mingw-32/unioffice.msi and
MSI/build/winegcc-64/unioffice.msi. Verify them from the repository root with
`sha256sum -c docs/native-windows-packages.sha256`. Repackaging creates a new
PackageCode and changes the hashes.

## Evidence

Guest base: C:\Users\Public\tmp\claude\unioffice-native. Logs:

- native-fixed.log; full results under guest ~/tmp/claude/unioffice-windows-b1a3982cc750491e9ef5ee9c08d24b17.
- native-standard-user.log; results under unioffice-windows-ae0300f14d8742c3a2abe256be915169.
- native-msi-lifecycle-3.log; results under unioffice-msi-windows-28ec6a3539c242039e4797339e189e42.
- typeinfo-lifetime-32-red.log and repair-from-user-before.log preserve both defects.

Verified archive on builder64: ~/tmp/claude/unioffice-native-validation/native-results.zip
(SHA256 2be6f05cf3ea73a46e255c923ab17ac70de3e205fd52a5b67a445ebbfdc378c9),
239 entries, extracted under logs/. Wine suites: unioffice-excel-suite-m_qf01g8
(32), -grvhlc2r (64); final Wine lifecycle: unioffice-native-scope-wine-lifecycle.log.

Only this Windows 10/backend/client workflow was validated. Other Windows
versions, broader concurrency and unsupported APIs remain outside this result.
The earlier parallel Wine PE32 test5 fault remains unisolated. Word is still a
stub; this stage does not expand its scope. Neither confirmed defect is a Wine
bug, so no Wine issue was filed.
