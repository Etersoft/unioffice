/* A GetTypeInfo result must own a reference independent of the cached pointer. */
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
    DISPID id;
    DISPPARAMS params = {arguments, NULL, count, 0};
    HRESULT hr = IDispatch_GetIDsOfNames(object, &IID_NULL, &member, 1, 0, &id);
    if (FAILED(hr)) return hr;
    return IDispatch_Invoke(object, id, &IID_NULL, 0, flags, &params, result, NULL, NULL);
}

static void contract(IDispatch *object, const char *label)
{
    ITypeInfo *first = NULL, *second = NULL, *invalid = (ITypeInfo*)1;
    TYPEATTR *attributes = NULL;
    UINT round;
    printf("TypeInfo lifetime: %s\n", label);
    for (round = 0; round < 3; ++round) {
        CHECK(IDispatch_GetTypeInfo(object, 0, 0, &first) == S_OK && first);
        CHECK(IDispatch_GetTypeInfo(object, 0, 0, &second) == S_OK && second);
        if (!first || !second) break;
        ITypeInfo_Release(first);
        first = NULL;
        /* A borrowed result can now point at freed type information. */
        CHECK(ITypeInfo_GetTypeAttr(second, &attributes) == S_OK && attributes);
        if (attributes) { ITypeInfo_ReleaseTypeAttr(second, attributes); attributes = NULL; }
        ITypeInfo_Release(second);
        second = NULL;
    }
    CHECK(IDispatch_GetTypeInfo(object, 0, 0, NULL) == E_POINTER);
    CHECK(IDispatch_GetTypeInfo(object, 1, 0, &invalid) == DISP_E_BADINDEX && !invalid);
    if (first) ITypeInfo_Release(first);
    if (second) ITypeInfo_Release(second);
}

int __cdecl main(void)
{
    CLSID clsid;
    IDispatch *excel = NULL;
    VARIANT books, book, sheet, range, value, address, collection;
    const WCHAR *members[] = {L"Font", L"Interior", L"Borders"};
    const WCHAR *sheet_members[] = {L"PageSetup", L"Outline"};
    unsigned int i;
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return 2;
    setvbuf(stdout, NULL, _IONBF, 0);
    VariantInit(&books); VariantInit(&book); VariantInit(&sheet);
    VariantInit(&range); VariantInit(&value); VariantInit(&address); VariantInit(&collection);
    CHECK(CLSIDFromProgID(L"Excel.Application", &clsid) == S_OK);
    CHECK(CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER, &IID_IDispatch, (void**)&excel) == S_OK && excel);
    if (!excel) goto done;
    contract(excel, "Application");
    CHECK(invoke(excel, L"Workbooks", DISPATCH_PROPERTYGET, NULL, 0, &books) == S_OK);
    if (V_VT(&books) != VT_DISPATCH) goto done;
    contract(V_DISPATCH(&books), "Workbooks");
    CHECK(invoke(V_DISPATCH(&books), L"Add", DISPATCH_METHOD, NULL, 0, &book) == S_OK);
    if (V_VT(&book) == VT_DISPATCH) contract(V_DISPATCH(&book), "Workbook");
    CHECK(invoke(excel, L"ActiveSheet", DISPATCH_PROPERTYGET, NULL, 0, &sheet) == S_OK);
    if (V_VT(&sheet) != VT_DISPATCH) goto done;
    contract(V_DISPATCH(&sheet), "Worksheet");
    CHECK(invoke(excel, L"Sheets", DISPATCH_PROPERTYGET, NULL, 0, &collection) == S_OK);
    if (V_VT(&collection) == VT_DISPATCH) contract(V_DISPATCH(&collection), "Sheets");
    VariantClear(&collection);
    V_VT(&address) = VT_BSTR;
    V_BSTR(&address) = SysAllocString(L"B2:C3");
    CHECK(invoke(V_DISPATCH(&sheet), L"Range", DISPATCH_PROPERTYGET, &address, 1, &range) == S_OK);
    if (V_VT(&range) == VT_DISPATCH) {
        contract(V_DISPATCH(&range), "Range");
        for (i = 0; i < sizeof(members) / sizeof(members[0]); ++i) {
            CHECK(invoke(V_DISPATCH(&range), members[i], DISPATCH_PROPERTYGET, NULL, 0, &value) == S_OK);
            if (V_VT(&value) == VT_DISPATCH) {
                contract(V_DISPATCH(&value), "Range child");
                if (i == 2) {
                    VARIANT index;
                    VariantInit(&index); V_VT(&index) = VT_I4; V_I4(&index) = 7;
                    CHECK(invoke(V_DISPATCH(&value), L"Item", DISPATCH_PROPERTYGET, &index, 1, &collection) == S_OK);
                    if (V_VT(&collection) == VT_DISPATCH) contract(V_DISPATCH(&collection), "Border");
                    VariantClear(&collection);
                }
            }
            VariantClear(&value);
        }
    }
    for (i = 0; i < sizeof(sheet_members) / sizeof(sheet_members[0]); ++i) {
        CHECK(invoke(V_DISPATCH(&sheet), sheet_members[i], DISPATCH_PROPERTYGET, NULL, 0, &value) == S_OK);
        if (V_VT(&value) == VT_DISPATCH) contract(V_DISPATCH(&value), "Sheet child");
        VariantClear(&value);
    }
done:
    VariantClear(&collection); VariantClear(&address); VariantClear(&value);
    VariantClear(&range); VariantClear(&sheet); VariantClear(&book); VariantClear(&books);
    if (excel) { invoke(excel, L"Quit", DISPATCH_METHOD, NULL, 0, NULL); IDispatch_Release(excel); }
    CoUninitialize();
    printf("TypeInfo lifetime: %d failure(s)\n", failures);
    return failures ? 1 : 0;
}
