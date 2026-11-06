# Multi-office and Word roadmap

The long-term direction, confirmed on 2026-11-06, is to make UniOffice an office
compatibility layer supporting multiple office suites and both Excel and Word
Automation. Broad office coverage is a development goal; the current verified
delivery remains the Excel subset over Windows Apache OpenOffice 4.1.17.
See [supported usage](supported-usage.md) for present capabilities.

## Architecture direction

Separate the Microsoft-compatible client interfaces from the office-specific
implementation. Excel and Word need their own document models, with shared
infrastructure for activation, values, errors and object lifetimes. The planned
backend boundary should accommodate:

- The working Windows OpenOffice Automation backend.
- Native Linux LibreOffice/OpenOffice through the proposed COM/UNO bridge.
- Further office suites through adapters to their available APIs.

The current Excel code calls UNO through IDispatch directly; a general backend
abstraction has not been implemented. Introduce that boundary incrementally as
real backends require it, preserving the published Excel COM interfaces and
existing regression behavior.

Keep the Linux bridge independent of Calc-specific behavior so that Writer can
use the same object wrapping, value conversion and transport. Native UNO code
should use the C++ SDK and its resource-owning types. New C++ components should
manage COM references and Automation values through small ownership wrappers;
the PE/Unix boundary remains explicit C data and opaque handles. The existing
C Excel adapter does not need a wholesale language rewrite.

Backend choice must be explicit when several offices are installed. Describe
capabilities per backend and return useful errors for unavailable operations.
Record suite/version, host platform, client architecture, tested Excel/Word
workflows and installer lifecycle independently. A suite's file-format support
alone does not establish Automation compatibility.

## Implementation sequence

1. Preserve the accepted Excel PE32/PE64 packages and tests as the baseline.
   Continue fixing reproducible defects in the existing workflows.
2. Bring up the Linux UNO backend: rebuild the SDK client, validate a Calc cell
   round trip, add the COM frontend and run the existing Excel suites. Check a
   basic Writer document through the shared bridge before treating it as a
   general office bridge. Validate PE32/WOW64 and deployment separately.
3. Implement the first Word Automation workflow in `unioffice_word/`: create or
   open a document, write text, save, close and quit. Review the existing IDL
   and registration first; preserve published contracts where applicable.
   Add focused tests for that workflow and object lifetimes. Broader formatting,
   tables and template operations should follow concrete client requirements.
4. Package Word only after its activation and functional workflow work. Check
   Excel/Word installation, repair, coexistence and removal together, without
   regressing the accepted Excel packages.
5. Add further office adapters in bounded stages, selecting suites/API paths
   from actual deployment needs and running the same relevant client scenarios.

This sequence is a proposed implementation order, not an implementation or
acceptance claim. Word currently has a stub class factory returning E_NOTIMPL
and is excluded from MSI; neither the Linux bridge nor additional suite
adapters are present yet.

## Starting material

- [Linux bridge findings, SDK sources and prototype checkpoints](linux-office-bridge.md).
- [Related projects](related-projects.md), including the standard Automation
  Bridge and COLEAT's compatibility/tracing approach.
- [Native Windows baseline and accepted artifacts](native-windows-results.md).

The earlier request to finish the existing feature set still guides each
delivery stage. It does not exclude the explicitly requested future Word and
multi-office work. Keep those additions as separate, testable changes rather
than implementing every unimplemented interface member at once.
