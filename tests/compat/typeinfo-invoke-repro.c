/* ITypeInfo::Invoke on a dual-interface property getter with optional and
   defaultvalue arguments and no [lcid] (Range.AddressLocal from the UniOffice
   type library), compared with a similar getter that has [lcid] (Range.Address).
   No office is needed: the object is a fake whose vtable slot for the member
   points at a function that records the call.
   Build: i686-w64-mingw32-gcc or winegcc; run: repro.exe path\to\unioffice_excel.tlb */
#define COBJMACROS
#include <windows.h>
#include <ole2.h>
#include <stdio.h>

static const GUID IID_I_Range = {0x00020846, 0x0001, 0x0000, {0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static int called;

static void show(const char *name, VARIANT *v)
{
    printf("    %s: vt=%d", name, V_VT(v));
    if (V_VT(v) == VT_ERROR) printf(" scode=%08lx", (unsigned long)V_ERROR(v));
    if (V_VT(v) == VT_I2) printf(" value=%d", V_I2(v));
    if (V_VT(v) == VT_BOOL) printf(" value=%d", V_BOOL(v));
    printf("\n");
}

static HRESULT WINAPI fake_AddressLocal(void *This, VARIANT RowAbsolute, VARIANT ColumnAbsolute,
        LONG ReferenceStyle, VARIANT External, VARIANT RelativeTo, BSTR *RHS)
{
    called = 1;
    printf("  AddressLocal called: ReferenceStyle=%ld\n", (long)ReferenceStyle);
    show("RowAbsolute", &RowAbsolute);
    show("ColumnAbsolute", &ColumnAbsolute);
    show("External", &External);
    show("RelativeTo", &RelativeTo);
    *RHS = SysAllocString(L"$A$1");
    return S_OK;
}

static HRESULT WINAPI fake_Address(void *This, VARIANT RowAbsolute, VARIANT ColumnAbsolute,
        LONG ReferenceStyle, VARIANT External, VARIANT RelativeTo, LCID lcid, BSTR *RHS)
{
    called = 1;
    printf("  Address called: ReferenceStyle=%ld lcid=%lx\n", (long)ReferenceStyle, (unsigned long)lcid);
    show("RowAbsolute", &RowAbsolute);
    *RHS = SysAllocString(L"$A$1");
    return S_OK;
}

static HRESULT WINAPI fake_other(void *This)
{
    printf("  unexpected vtable slot called\n");
    return E_UNEXPECTED;
}

static void *vtable[1024];
static struct { void **lpVtbl; } object = { vtable };

static void call(ITypeInfo *ti, const WCHAR *name, void *impl, UINT argc, VARIANT *args)
{
    LPOLESTR member = (LPOLESTR)name;
    MEMBERID id;
    FUNCDESC *fd;
    UINT i, index, bad = (UINT)-1;
    DISPPARAMS params = {args, NULL, argc, 0};
    VARIANT result;
    EXCEPINFO info;
    HRESULT hr;

    hr = ITypeInfo_GetIDsOfNames(ti, &member, 1, &id);
    if (FAILED(hr)) { printf("%ls: GetIDsOfNames %08lx\n", name, (unsigned long)hr); return; }
    for (i = 0; ITypeInfo_GetFuncDesc(ti, i, &fd) == S_OK; i++) {
        if (fd->memid == id && fd->invkind == INVOKE_PROPERTYGET) {
            index = fd->oVft / sizeof(void*);
            printf("%ls: memid=%ld oVft=%d (slot %u) cParams=%d cParamsOpt=%d callconv=%d funckind=%d\n",
                    name, (long)id, fd->oVft, index, fd->cParams, fd->cParamsOpt, fd->callconv, fd->funckind);
            for (i = 0; i < fd->cParams; i++) {
                PARAMDESC *pd = &fd->lprgelemdescParam[i].paramdesc;
                printf("  param %u: vt=%d flags=%x", i, fd->lprgelemdescParam[i].tdesc.vt, pd->wParamFlags);
                if (pd->wParamFlags & PARAMFLAG_FHASDEFAULT) {
                    if (!pd->pparamdescex) printf(" default: no PARAMDESCEX");
                    else {
                        VARIANT *def = &pd->pparamdescex->varDefaultValue;
                        printf(" default: cBytes=%lu vt=%d", (unsigned long)pd->pparamdescex->cBytes, V_VT(def));
                        if (V_VT(def) == VT_I4) printf(" value=%ld", (long)V_I4(def));
                        if (V_VT(def) == VT_I2) printf(" value=%d", V_I2(def));
                    }
                }
                printf("\n");
            }
            vtable[index] = impl;
            ITypeInfo_ReleaseFuncDesc(ti, fd);
            break;
        }
        ITypeInfo_ReleaseFuncDesc(ti, fd);
    }
    VariantInit(&result);
    memset(&info, 0, sizeof(info));
    called = 0;
    hr = ITypeInfo_Invoke(ti, &object, id, DISPATCH_PROPERTYGET, &params, &result, &info, &bad);
    printf("  Invoke(%u args) -> hr=%08lx called=%d result vt=%d argerr=%d\n\n",
            argc, (unsigned long)hr, called, V_VT(&result), (int)bad);
    VariantClear(&result);
}

int __cdecl main(int argc, char **argv)
{
    WCHAR path[MAX_PATH];
    ITypeLib *tl;
    ITypeInfo *ti;
    VARIANT args[2];
    HRESULT hr;
    int i;

    setvbuf(stdout, NULL, _IONBF, 0);
    if (argc < 2) { printf("usage: %s unioffice_excel.tlb\n", argv[0]); return 2; }
    MultiByteToWideChar(CP_ACP, 0, argv[1], -1, path, MAX_PATH);
    for (i = 0; i < 1024; i++) vtable[i] = fake_other;
    hr = LoadTypeLibEx(path, REGKIND_NONE, &tl);
    if (FAILED(hr)) { printf("LoadTypeLibEx %08lx\n", (unsigned long)hr); return 2; }
    hr = ITypeLib_GetTypeInfoOfGuid(tl, &IID_I_Range, &ti);
    if (FAILED(hr)) { printf("GetTypeInfoOfGuid %08lx\n", (unsigned long)hr); return 2; }

    call(ti, L"AddressLocal", fake_AddressLocal, 0, NULL);
    V_VT(&args[0]) = VT_BOOL; V_BOOL(&args[0]) = VARIANT_FALSE;   /* ColumnAbsolute */
    V_VT(&args[1]) = VT_BOOL; V_BOOL(&args[1]) = VARIANT_TRUE;    /* RowAbsolute */
    call(ti, L"AddressLocal", fake_AddressLocal, 2, args);
    call(ti, L"Address", fake_Address, 0, NULL);
    call(ti, L"Address", fake_Address, 2, args);

    ITypeInfo_Release(ti);
    ITypeLib_Release(tl);
    return 0;
}
