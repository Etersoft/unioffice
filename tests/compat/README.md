# Wine/Windows behaviour differences

Standalone examples for differences found while testing UniOffice on Wine
and native Windows 10. They do not need UniOffice code at run time.

- `typeinfo-invoke-repro.c`: `ITypeInfo::Invoke` on `Range.AddressLocal`
  from `unioffice_excel.tlb`. Windows passes garbage for the enum parameter
  with `[defaultvalue]` (PE32 shifts the following arguments and crashes);
  Wine passes 1. `Range.Address`, which also has `[lcid]`, works on both.

  ```sh
  i686-w64-mingw32-gcc -Wall typeinfo-invoke-repro.c -loleaut32 -lole32 -o repro32.exe
  winegcc -b x86_64-windows -m64 -Wall typeinfo-invoke-repro.c -loleaut32 -lole32 -o repro64.exe
  wine repro64.exe 'C:\windows\system32\unioffice_excel.tlb'
  ```

- `vbscript-self-assign.vbs`: `a = a(0)` assigns the inner array on Wine;
  on Windows the variable is left unchanged (an error with code 0 without
  `On Error Resume Next`).
- `vbscript-nested-array-repro.vbs`: reading a UNO `sequence<sequence<Point>>`
  (`LineShape.PolyPolygon`) works on both when the elements are read into a
  separate variable. It needs a running OpenOffice.
