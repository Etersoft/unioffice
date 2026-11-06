/* Range.Value of an error formula is VT_ERROR with Excel's CVErr code;
   VBScript cannot read the code, so this probe checks it natively. */
#define COBJMACROS
#include <windows.h>
#include <ole2.h>
#include <stdio.h>

static int failures;
#define CHECK(condition) do { if (!(condition)) { \
    fprintf(stderr, "FAIL line %d: %s\n", __LINE__, #condition); ++failures; \
} } while (0)

static HRESULT invoke(IDispatch *object, const WCHAR *name, WORD flags,
        VARIANT *arguments, UINT count, VARIANT *result)
{
    LPOLESTR member = (LPOLESTR)name;
    DISPID id, put = DISPID_PROPERTYPUT;
    DISPPARAMS params = {arguments, NULL, count, 0};
    HRESULT hr = IDispatch_GetIDsOfNames(object, &IID_NULL, &member, 1, 0, &id);
    if (FAILED(hr)) return hr;
    if (flags == DISPATCH_PROPERTYPUT) {
        params.rgdispidNamedArgs = &put;
        params.cNamedArgs = 1;
    }
    return IDispatch_Invoke(object, id, &IID_NULL, 0, flags, &params, result, NULL, NULL);
}

static void check_formula(IDispatch *sheet, const WCHAR *cell, const WCHAR *formula, LONG code)
{
    VARIANT address, range, text, value;
    VariantInit(&range);
    VariantInit(&value);
    V_VT(&address) = VT_BSTR;
    V_BSTR(&address) = SysAllocString(cell);
    CHECK(invoke(sheet, L"Range", DISPATCH_PROPERTYGET, &address, 1, &range) == S_OK);
    if (V_VT(&range) == VT_DISPATCH) {
        V_VT(&text) = VT_BSTR;
        V_BSTR(&text) = SysAllocString(formula);
        CHECK(invoke(V_DISPATCH(&range), L"Formula", DISPATCH_PROPERTYPUT, &text, 1, NULL) == S_OK);
        CHECK(invoke(V_DISPATCH(&range), L"Value", DISPATCH_PROPERTYGET, NULL, 0, &value) == S_OK);
        CHECK(V_VT(&value) == VT_ERROR);
        CHECK(V_ERROR(&value) == MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, code));
        if (V_VT(&value) != VT_ERROR || V_ERROR(&value) != MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, code))
            fprintf(stderr, "  %ls: vt=%d scode=%08lx\n", formula, V_VT(&value), (unsigned long)V_ERROR(&value));
        VariantClear(&text);
    }
    VariantClear(&value);
    VariantClear(&range);
    VariantClear(&address);
}

int __cdecl main(void)
{
    CLSID clsid;
    IDispatch *excel = NULL;
    VARIANT books, book, sheet;
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return 2;
    VariantInit(&books); VariantInit(&book); VariantInit(&sheet);
    CHECK(CLSIDFromProgID(L"Excel.Application", &clsid) == S_OK);
    hr = CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER, &IID_IDispatch, (void**)&excel);
    CHECK(hr == S_OK && excel);
    if (!excel) goto done;
    CHECK(invoke(excel, L"Workbooks", DISPATCH_PROPERTYGET, NULL, 0, &books) == S_OK);
    if (V_VT(&books) != VT_DISPATCH) goto done;
    CHECK(invoke(V_DISPATCH(&books), L"Add", DISPATCH_METHOD, NULL, 0, &book) == S_OK);
    CHECK(invoke(excel, L"ActiveSheet", DISPATCH_PROPERTYGET, NULL, 0, &sheet) == S_OK);
    if (V_VT(&sheet) != VT_DISPATCH) goto done;
    check_formula(V_DISPATCH(&sheet), L"A1", L"=1/0", 2007);   /* xlErrDiv0 */
    check_formula(V_DISPATCH(&sheet), L"A2", L"=NA()", 2042);  /* xlErrNA */
    check_formula(V_DISPATCH(&sheet), L"A3", L"=SQRT(-1)", 2036); /* xlErrNum */
done:
    VariantClear(&sheet);
    if (V_VT(&book) == VT_DISPATCH) {
        VARIANT save;
        V_VT(&save) = VT_BOOL;
        V_BOOL(&save) = VARIANT_FALSE;
        invoke(V_DISPATCH(&book), L"Close", DISPATCH_METHOD, &save, 1, NULL);
    }
    VariantClear(&book); VariantClear(&books);
    if (excel) { invoke(excel, L"Quit", DISPATCH_METHOD, NULL, 0, NULL); IDispatch_Release(excel); }
    CoUninitialize();
    printf("Range value errors: %d failure(s)\n", failures);
    return failures ? 1 : 0;
}
