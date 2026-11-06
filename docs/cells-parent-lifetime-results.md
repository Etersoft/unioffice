# Cells ranges retain their worksheet

Worksheet.Cells previously returned its private cached Range. That cache uses
a weak worksheet pointer to avoid a reference cycle. Application.Cells creates
a temporary worksheet wrapper and releases it before returning the cached range,
leaving the returned range's Parent pointer dangling.

The dedicated tests/cells-parent-lifetime.vbs reproduced a PE64 execute-access
fault at address2 on Cells.Parent. Wine entered its debugger and the bounded
runner timed out with exit124. This test contains no string indices, separating
the lifetime defect from the argument-copying failures found in the same stage.

Worksheet.Cells now returns a separate Range initialized with a strong worksheet
reference. The worksheet's private weak cache is retained internally, so the
returned object preserves its parent without creating a worksheet/cache cycle.
Existing constructors and AddRef/Release rules are reused; no public interface
or type-library signature changed.

On MinGW PE32 (built on Fedora) and winegcc PE64 (built on builder64), the probe
passes49 checks each: Application.Cells.Parent, repeated Worksheet.Cells after
releasing the worksheet variable, a child cell after releasing the Cells range,
and its Parent. The cell's actual value is checked through a separate Range.
The same final DLLs also pass the argument-ownership probes.

Environment: builder64, Wine@Etersoft 11.20.1-eter1, Windows Apache OpenOffice
4.1.17 x86. Within Xvfb, using the dedicated WINEPREFIX and CLIENT_BITS:

```sh
sh tests/run-openoffice-backend.sh cells-parent-lifetime.vbs 'Cells parent lifetime: PASS'
```

Logs: ~/tmp/claude/unioffice-cells-parent-lifetime-64-red.log and
unioffice-range-ownership-{32,64}-lifetime-final.log. An earlier PE64 string-index
probe also faulted before this lifetime fix; its log is retained. The previous
intermittent legacy test5 crash had no matching isolated backtrace and is not
retroactively attributed to this defect. No Wine defect was established.
