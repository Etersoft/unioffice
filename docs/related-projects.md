# Related projects and public bridge search

Checked on 2026-11-06 using web search and authenticated GitHub repository/code
search. No ready-to-use Wine COM/IDispatch-to-Linux-UNO bridge was identified.
This is a search result, not proof that no implementation exists: GitHub code
indexing and the reviewed default branches are not exhaustive.

## Sources worth reusing

### LibreOffice and Apache OpenOffice Automation Bridges

Both projects contain the actual Windows Automation implementation:

- [LibreOffice: extensions/source/ole](https://github.com/LibreOffice/core/tree/master/extensions/source/ole).
- [Apache OpenOffice: main/extensions/source/ole](https://github.com/apache/openoffice/tree/trunk/main/extensions/source/ole).
- [Automation Bridge requirements and behavior](https://wiki.openoffice.org/wiki/Documentation/DevGuide/ProUNO/Bridge/Automation_Bridge).

This is the closest source for dispatch semantics and COM/UNO type conversion.
LibreOffice's `unoobjw.cxx` wraps UNO objects for COM;
`unoconversionutilities.hxx` handles Any/VARIANT, sequences and structures.
Its [build configuration](https://github.com/LibreOffice/core/blob/master/extensions/Module_extensions.mk)
selects `Library_oleautobridge` under WNT, and the wrapper uses ATL/Windows
headers. The documented Automation interface is Windows-only. These sources
provide implementation material, not an already packaged Linux/Wine component.

Before writing a replacement conversion layer, evaluate which behavior/code
can be reused behind the proposed PE/Unix boundary. Porting dependencies and
runtime behavior still need a prototype; no build/runtime claim is made here.

### Collabora COLEAT

[COLEAT](https://github.com/CollaboraOnline/COLEAT) is a C++ application launcher
that redirects a client's Word/Excel COM use to Collabora Office. The
[project description](https://www.collaboraonline.com/coleat/) explains DLL
injection/IAT interception and generated proxies for early-bound clients such
as compiled VB6 programs. Its trace-only mode helps determine the APIs a real
client actually uses.

The inspected head is `66e4eb0103ea14bd2f184f85945a3315d4c9b799`, dated
2020-12-08. Its [README](https://github.com/CollaboraOnline/COLEAT/blob/66e4eb0103ea14bd2f184f85945a3315d4c9b799/README.txt)
says the project is inactive and requires Collabora Office 6.0 or newer for
redirection. [BUILD.txt](https://github.com/CollaboraOnline/COLEAT/blob/66e4eb0103ea14bd2f184f85945a3315d4c9b799/BUILD.txt)
uses Visual Studio and Office type libraries for proxy generation. This is
Windows compatibility/tracing tooling, not a native Linux UNO backend.
`TODO.txt` retains interface/proxy limitations. The inspected injected-DLL
source carries an MPL-2.0 notice; repository metadata alone reported no license.

Potential use: reference early-binding adapters and trace existing customer
workflows to constrain compatibility work. No replacement of UniOffice with
COLEAT, port or execution has been validated.

### Native SDK examples: DocumentLoader

Public source is available for the same kind of native UNO client as the old
binary in `../unioffice-host`:

- [LibreOffice DocumentLoader.cxx](https://github.com/LibreOffice/core/blob/master/odk/examples/cpp/DocumentLoader/DocumentLoader.cxx).
- [Apache OpenOffice DocumentLoader.cxx](https://github.com/apache/openoffice/blob/trunk/main/odk/examples/cpp/DocumentLoader/DocumentLoader.cxx).

The examples bootstrap a local context, resolve a socket/URP endpoint, obtain
the office context and load a document. They include the port-2083 default also
visible in the historical binary. Exact source equivalence with that binary
was not established. Use the SDK example matching the selected Linux office
to rebuild a client from source; no reverse engineering of the old executable
is needed to start this experiment.

### unoserver and other UNO clients

[unoserver](https://github.com/unoconv/unoserver) manages a LibreOffice listener
and conversion/comparison requests using Python UNO and XML-RPC. Its
[README](https://github.com/unoconv/unoserver/blob/master/README.rst) provides
process/profile/readiness ideas. It does not expose the COM object graph
expected by `Excel.Application`, so it cannot replace the missing bridge.

Similarly, language bindings and conversion utilities do not by themselves
supply COM activation or Automation value conversions to a Windows client.

## Our public history and search scope

The [Etersoft/unioffice mirror](https://github.com/Etersoft/unioffice) currently
ends at `128339ddffa654bbafa7e9ab5c1e39007b17c5f9` (2010-11-12), already present
in this checkout's history. It is not a newer Linux-backend implementation.
The [historical project wiki](https://wiki.etersoft.ru/UniOffice) describes host
OpenOffice integration as a plan. The similarly named `unidoc/unioffice` is an
unrelated Go document-format library.

Representative searches included `Wine Linux LibreOffice COM UNO bridge`,
`oleautobridge wine`, and GitHub code queries for `Bridge_GetStruct`,
`com.sun.star` with `unixlib`, and ServiceManager implementations. Reviewed
matches included conversion utilities, Windows Automation clients and the
projects above. The Unixlib/UNO query produced no match; that alone cannot
establish absence of a bridge. No candidate was installed or runtime-tested.

The remaining work is still the [Linux bridge prototype](linux-office-bridge.md):
native UNO cell round trip first, then a compatible COM frontend, existing
Excel tests, and explicit PE32/WOW64 validation.
