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

## What the historical mingw-widl port provided

The repository has seven commits, all by Vitaly Lipatov on 2008-12-29. Its
Makefile builds `widl.exe` with `i386-mingw32msvc-gcc`; the existing untracked
binary is PE32. The observable purpose was a standalone Windows build of Wine's
IDL compiler, able to generate COM headers, IID definitions, type libraries
and RPC stubs. A particular UniOffice workflow using this adjacent executable
was not established: the inspected historical Excel Makefile selects `widl`
without referring to this directory.

| Commit | Change |
| --- | --- |
| `5efab3c` | Imports the compiler and standalone Makefile |
| `e1d6076`, `89dd195` | Supply Wine headers/configuration and the include path |
| `b9338be` | Corrects cleanup and Windows output suffixes |
| `726577c` | Comments out language IDs missing from the then-current MinGW headers; the message says `hasp.c`, but changes `hash.c` |
| `0bfddb6` | Adds Wine's WPP preprocessor and mkstemps source and links libwpp |
| `93da5cc` | Builds libwpp before recursively building widl.exe |

Compared the initial import against local upstream Wine tag `wine-1.1.11`
(commit `0b3b6e67ea663e853cc6bb93f4da447ab934e50d`): all 23 compiler source/header/
lexer/parser files match byte for byte. At the final revision, `hash.c` is the
only changed compiler source. Added support sources/headers also match that tag
except `include/wine/port.h`, whose differences concern mkdir, poll and symlink
portability declarations. No custom UNO transport or IDL-generation fix was
found in the import or subsequent commits.

The useful result is the standalone build recipe, especially `0bfddb6` plus
`93da5cc`. There is no demonstrated fix to cherry-pick into the current UniOffice
toolchain: it already uses modern WIDL and has passed both PE builds. Preserve
the history as a reference; do not revive the old compiler or its language-ID
workaround without a concrete compatibility failure. This inspection did not
rebuild/run the historical tool or modify its existing generated files.

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
