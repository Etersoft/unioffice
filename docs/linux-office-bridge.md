# Linux office backend: findings and proposed design

Inspected on 2026-11-06. The accepted Windows/Wine results use Windows Apache
OpenOffice. A Linux office backend is still an open delivery requirement;
this investigation does not establish runtime support for it.

## Existing local work

- `../unioffice-host` is a separate Git repository. Its only commit,
  `667a4c0`, contains four files: `DocumentLoader`, `DocumentLoader.rdb`,
  `start.sh` and `test.odt`. There are no tracked source files or COM frontend.
  `DocumentLoader` is an unstripped Linux i386 executable linked to native
  `libuno_cppuhelpergcc3`, `libuno_cppu` and `libuno_sal`. Symbols/strings show
  UNO bootstrapping, UnoUrlResolver and XComponentLoader. Its launcher connects
  to `uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager`.
  This is useful evidence of an earlier native UNO client, not a completed
  COM-to-UNO bridge. The historical executable was inspected, not run.
- `../unioffice-0.6/OOWrappers/oo_servicemanager.cpp` still calls
  `CLSIDFromProgID` and `CoCreateInstance` for `com.sun.star.ServiceManager`.
  The classes under `OOWrappers/com/sun/star/` wrap `IDispatch`; their names do
  not make them native UNO SDK bindings.
- `../mingw-widl` is tooling for IDL/proxy generation, not the Linux UNO backend.
- The modern local example is
  `/srv/lav/Projects/CRYPTO/cpcsp_proxy_unixcall`: `proxy.c` calls
  `WINE_UNIX_CALL`, `unixlib.c` exports native and WOW64 function tables, and
  its build separates `cpcsp_proxy.dll` from `cpcsp_proxy.so`.

## Recommended component boundary

Develop the bridge in the existing `unioffice-host` repository as a separately
packaged Wine component. Keep Excel compatibility in this repository. Proposed
flow (component names are illustrative):

```text
Windows client -> UniOffice Excel DLL -> COM IDispatch frontend (PE DLL)
                                             |
                                       Wine unix calls
                                             |
                                native UNO adapter (ELF .so)
                                             |
                                  UNO URL resolver / URP
                                             |
                                  Linux LibreOffice process
```

The PE frontend implements COM activation, dispatch and COM-owned results. The
native C++ side uses the Linux office SDK/runtime, keeps UNO references and
performs reflection/invocation. Separate processes for client and office are
already accommodated by the UNO resolver; initially no extra custom host daemon
is needed. This is a design recommendation, not a validated implementation.

Use opaque object/struct handles, fixed-width fields and explicitly described
buffers at the Unix boundary. Do not pass UNO C++ objects/vtables across it or
reinterpret `VARIANT` as UNO `Any`. Define ownership, identity and release rules;
catch native exceptions and return COM errors with useful descriptions.

The existing code requires more than scalar calls: `special_functions.c` uses
`Bridge_GetStruct`, mutable struct fields and SAFEARRAYs of dispatch objects.
Sequences, nested arrays, argument direction/order, property get/put and numeric
type conversion therefore belong in the initial compatibility design.

## Reuse and build choices

The LibreOffice [Invocation service](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1script_1_1Invocation.html)
constructs adapters exposing [XInvocation](https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1script_1_1XInvocation.html)
for methods and properties. The [UNO URL resolver](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1bridge_1_1UnoUrlResolver.html)
connects to remote UNO objects. These are candidates for the native adapter.
The existing [OLE conversion code](https://github.com/LibreOffice/core/blob/master/extensions/source/ole/unoconversionutilities.hxx)
is a reference for Any/VARIANT, sequence and struct behavior. It is not a ready
Linux library: [the build selects oleautobridge for WNT](https://github.com/LibreOffice/core/blob/master/extensions/Module_extensions.mk),
and [its wrapper uses ATL/Windows headers](https://github.com/LibreOffice/core/blob/master/extensions/source/ole/unoobjw.cxx).

Build the Wine-specific PE frontend with winegcc on builder64 and the ELF side
with a native C++ compiler and matching Linux UNO SDK. Preserve the established
Fedora-only direct MinGW workflow for ordinary Windows artifacts. Start with
PE64/ELF64; PE32/WOW64 needs explicit argument/pointer translation, as in cpcsp.
Use the installed Wine version's Unix-call interface rather than copying an
old private header unquestioningly. An ELF Unix library is intentional here;
the frontend remains PE, not a legacy Winelib `.dll.so`.

## First implementation checkpoints

1. Rebuild a native UNO client from source; connect to a dedicated Linux office
   profile and create a Calc document, write/read a cell, then close it.
2. Implement the minimum COM/Unix adapter for that same round trip, including
   the PropertyValue structs needed to load a document. Run the existing direct
   backend probe before involving Excel compatibility.
3. Connect Excel.Application and run its cell probe, then all seven legacy tests.
   Add conversions needed by those existing scenarios before adding new APIs.
4. Validate PE32/WOW64, reference lifetimes and installation/removal separately.
   Backend selection/registration must be explicit and preserve the working
   Windows backend; do not silently overwrite its ServiceManager registration.
