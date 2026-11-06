# MSI32/MSI64 packaging and coexistence

Validated on builder64 with Wine@Etersoft 11.20.1-eter1, using PE32 MinGW output
built on Fedora and PE64 winegcc output built on builder64. Native Windows
installation was not exercised in this environment.
Native Windows10 was validated subsequently; see the
[native results and machine-scope fix](native-windows-results.md).

## Package construction

The packager detects the PE machine (i386 or AMD64), derives staged tables from
the legacy source tables, and imports architecture-specific SummaryInformation.
The checked-in Component/Directory/Property tables remain the 32-bit source.

| Item | MSI32 | MSI64 |
| --- | --- | --- |
| Template Summary | Intel;1033 | x64;1033 |
| Minimum Installer schema | 100 | 200 |
| Component Attributes | 2 | 258 (includes 256) |
| System directory property | SystemFolder | System64Folder |
| ProductCode | {643315BA-DEF8-4CBE-AB26-C54B71535DFD} | {4C7393FD-6C4A-5B3F-8EC2-B7AD16262864} |
| ComponentId | {36D68DB8-D5FE-4342-B983-D189234CABC7} | {F2E409F1-18B1-5FAD-ADAD-184D12608F75} |
| UpgradeCode | {CD87087B-70BE-4BE2-B7F3-987AB335659D} | {AB670464-4388-5009-B76D-7D4FA37EE2CF} |

A fresh PackageCode is generated for every build instead of reusing the old
2008 template package identity. Product identities remain stable per architecture.
The MSI64 product name includes x64. Paths for its registration actions use
System64Folder. Both packages retain the CP1251 import workaround.

[Microsoft's requirements](https://learn.microsoft.com/en-us/windows/win32/msi/using-64-bit-windows-installer-packages)
cover the platform Summary, component bit, directory properties and distinct IDs.

## Registration and repair fixes

The first coexistence attempt installed both variants successfully, but removing
MSI64 deleted shared Excel.Application ProgIDs and broke the retained PE32 client.
UniOffice unregister now preserves shared ProgIDs while the opposite registry
view still has a nonempty CLSID/InprocServer32 registration. Removing the last
architecture deletes the shared ProgIDs. Commit: 85a457b.

The native typelib-registration regression test failed on the previous DLL and
passes on both updated bitnesses. It now checks the shared ProgID as well as own
and peer typelib paths.

MSI registration actions follow component action state: register when Ooffice is
local ($Ooffice=3); unregister only on removal of an installed component
(Installed AND $Ooffice=2). This prevents repair from unregistering without
restoring registration, which the old Installed/NOT Installed conditions allowed.

## Validation

The enhanced tests/msi-smoke.c checks Unicode manufacturer text, ProductCode,
UpgradeCode, ComponentId/64-bit attribute, directory/action source, platform,
minimum schema and PackageCode shape. Both MSI32 and MSI64 returned zero failures.
Pass the expected bitness as the second argument (default32).

A separate test prefix was copied from the existing PE32/OpenOffice fixture:
~/tmp/claude/unioffice-msi64/prefix. The previous PE32 DLL was upgraded through
MSI repair; MSI64 was then installed alongside it. Payload files were compared
byte for byte to the build outputs.

scripts/test-msi-coexistence.py was run inside Xvfb in that disposable prefix,
starting with both updated products installed. It verified:

1. Both installed payloads match and both Excel.Application cell round trips pass.
2. MSI32 and MSI64 repair succeed; after each, both payloads match and both
   Automation clients work.
3. Removing MSI64 removes its files and preserves PE32 files and Automation.
4. Reinstalling MSI64 succeeds; removing MSI32 preserves PE64 files and Automation.
5. Removing the final MSI64 removes its files and the shared ProgID.

Result: MSI coexistence: PASS. Logs are retained under
~/tmp/claude/unioffice-msi64/coexistence-logs/ (resolved on the host as
/tmp/.private/lav/claude/unioffice-msi64/coexistence-logs/).
The test uninstalls both products from its prefix; it must never be run in a
normal user prefix. It requires WINEPREFIX under ~/tmp/claude and both packages
already installed from the current outputs. The other Automation test prefixes
were preserved.

An occupied Xvfb display was avoided with xvfb-run -a -n 210. No shared X lock
files were deleted. All installation results here are Wine results; broader
Windows/installer-version certification is outside this validation.
