# Application dual-interface ABI

QueryInterface(IID__Application) now returns the published _Application vtable.
Previously it returned the private Disp_Application adapter, whose methods use
direct return values instead of the IDL's HRESULT/output parameters. For example,
the adapter's Creator returns XlCreator, whereas _Application.Creator returns
HRESULT and writes XlCreator through an output pointer. Returning that adapter
under the dual IID made native COM calls use incompatible signatures.

The issue surfaced while testing native Application.Rows/Columns getters:
PE32 calls returned S_OK or an arbitrary positive value without filling outputs,
then the test faulted. The source confirmed the incorrect QueryInterface branch.
The log is ~/tmp/claude/unioffice-sheet-axes-32-native-before-qi.log. This is an
adapter defect; no Wine defect was established.

Only the IID__Application branch changes. Canonical IUnknown identity stays on
the existing private adapter, whose QueryInterface/AddRef/Release delegate to
the same object. IDispatch still exposes the published implementation. IDL,
IIDs, method order, native signatures and registration do not change.

tests/application-interface.c verifies HRESULT and guarded Creator output
through QueryInterface, a round trip through IDispatch, canonical IUnknown
identity, repeated QueryInterface, and direct CoCreateInstance(IID__Application).
Run the matching executable with the existing backend harness and require
the explicit `Application interface: 0 failure(s)` marker.

```sh
sh tests/run-openoffice-backend.sh unioffice_excel/build/mingw-32/application-interface.exe 'Application interface: 0 failure(s)'
```

Use winegcc-64 for PE64, a matching disposable prefix/client configuration and
Xvfb. Build PE32 with MinGW on Fedora and PE64 with winegcc on builder64, linking
unioffice_excel_i.o and ole32/oleaut32/uuid as for the other native COM tests.

Both PE32 and PE64 tests returned zero failures on 2026-11-06 with
Wine@Etersoft 11.20.1-eter1 and Windows Apache OpenOffice 4.1.17 x86. The expanded
native Rows/Columns/Range/Borders COM test also returned zero failures for both.
Logs: ~/tmp/claude/unioffice-sheet-axes-{32,64}-interface-final.log and
-native-final.log. Native Windows and remote COM clients were not tested.
