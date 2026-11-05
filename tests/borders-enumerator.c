/* IEnumVARIANT regression test; requires the disposable OpenOffice fixture. */
#define COBJMACROS
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "unioffice_excel.h"

static int failures;
#define CHECK(condition) do { if (!(condition)) { \
    fprintf(stderr, "FAIL line %d: %s\n", __LINE__, #condition); ++failures; \
} } while (0)
#define REQUIRE(condition) do { if (!(condition)) { \
    fprintf(stderr, "FAIL line %d: %s\n", __LINE__, #condition); \
    ++failures; goto done; \
} } while (0)

static HRESULT invoke(IDispatch *object, const WCHAR *name, WORD flags,
                      VARIANT *arguments, UINT count, VARIANT *result)
{
    DISPID id;
    LPOLESTR member = (LPOLESTR)name;
    DISPPARAMS params = {arguments, NULL, count, 0};
    HRESULT hr = IDispatch_GetIDsOfNames(object, &IID_NULL, &member, 1, 0, &id);
    if (FAILED(hr)) return hr;
    return IDispatch_Invoke(object, id, &IID_NULL, 0, flags, &params, result, NULL, NULL);
}

static IEnumVARIANT *enumerator(I_Borders *borders)
{
    IUnknown *unknown = NULL;
    IEnumVARIANT *result = NULL;
    CHECK(I_Borders_GetEnumerator(borders, &unknown) == S_OK);
    if (unknown) {
        CHECK(IUnknown_QueryInterface(unknown, &IID_IEnumVARIANT, (void **)&result) == S_OK);
        IUnknown_Release(unknown);
    }
    return result;
}

static void next(IEnumVARIANT *enumerator, ULONG requested, ULONG expected, HRESULT status)
{
    VARIANT values[65];
    ULONG fetched = 99, i;
    HRESULT hr;
    for (i = 0; i < 65; ++i) VariantInit(&values[i]);
    hr = IEnumVARIANT_Next(enumerator, requested, values, &fetched);
    CHECK(hr == status);
    CHECK(fetched == expected);
    for (i = 0; i < expected; ++i)
        CHECK(V_VT(&values[i]) == VT_DISPATCH && V_DISPATCH(&values[i]));
    for (i = 0; i < 65; ++i) VariantClear(&values[i]);
}

static void invalid_indices(I_Borders *borders)
{
    const LONG indices[] = {0, -1, 13, (-2147483647L - 1), 2147483647L};
    IDispatch *item;
    HRESULT hr;
    unsigned int i;
    for (i = 0; i < sizeof(indices) / sizeof(indices[0]); ++i) {
        item = (IDispatch *)1;
        hr = I_Borders_get_Item(borders, (XlBordersIndex)indices[i], &item);
        CHECK(hr == DISP_E_BADINDEX);
        CHECK(item == NULL);
        if (SUCCEEDED(hr) && item) IDispatch_Release(item);
        item = (IDispatch *)1;
        hr = I_Borders_get__Default(borders, (XlBordersIndex)indices[i], &item);
        CHECK(hr == DISP_E_BADINDEX);
        CHECK(item == NULL);
        if (SUCCEEDED(hr) && item) IDispatch_Release(item);
    }
}

static void weight_contract(I_Borders *borders)
{
    IDispatch *dispatch = NULL;
    I_Border *border = NULL;
    CHECK(I_Borders_get_Weight(borders, NULL) == E_POINTER);
    CHECK(I_Borders_put_Weight(borders, (XlBorderWeight)3) == E_INVALIDARG);
    CHECK(I_Borders_get_Item(borders, xlEdgeLeft, &dispatch) == S_OK && dispatch);
    if (!dispatch) return;
    CHECK(IDispatch_QueryInterface(dispatch, &IID_I_Border, (void **)&border) == S_OK);
    if (border) {
        CHECK(I_Border_get_Weight(border, NULL) == E_POINTER);
        CHECK(I_Border_get_LineStyle(border, NULL) == E_POINTER);
        CHECK(I_Border_put_Weight(border, (XlBorderWeight)3) == E_INVALIDARG);
        CHECK(I_Border_put_LineStyle(border, (XlLineStyle)12345) == E_INVALIDARG);
        I_Border_Release(border);
    }
    IDispatch_Release(dispatch);
}

int __cdecl main(void)
{
    CLSID clsid;
    IDispatch *excel = NULL;
    I_Borders *borders = NULL;
    IEnumVARIANT *first = NULL, *second = NULL, *clone = NULL;
    VARIANT books, workbook, range, collection, argument;
    LONG count = 0;
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return 2;
    VariantInit(&books);
    VariantInit(&workbook);
    VariantInit(&range);
    VariantInit(&collection);
    VariantInit(&argument);
    REQUIRE(CLSIDFromProgID(L"Excel.Application", &clsid) == S_OK);
    REQUIRE(CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER, &IID_IDispatch,
                             (void **)&excel) == S_OK);
    REQUIRE(invoke(excel, L"Workbooks", DISPATCH_PROPERTYGET, NULL, 0, &books) == S_OK);
    REQUIRE(invoke(V_DISPATCH(&books), L"Add", DISPATCH_METHOD, NULL, 0, &workbook) == S_OK);
    V_VT(&argument) = VT_BSTR;
    V_BSTR(&argument) = SysAllocString(L"A1:D4");
    REQUIRE(invoke(excel, L"Range", DISPATCH_PROPERTYGET, &argument, 1, &range) == S_OK);
    VariantClear(&argument);
    REQUIRE(invoke(V_DISPATCH(&range), L"Borders", DISPATCH_PROPERTYGET, NULL, 0, &collection) == S_OK);
    REQUIRE(IDispatch_QueryInterface(V_DISPATCH(&collection), &IID_I_Borders, (void **)&borders) == S_OK);
    REQUIRE(I_Borders_get_Count(borders, &count) == S_OK && count > 3 && count < 65);
    CHECK(count == 8);
    invalid_indices(borders);
    weight_contract(borders);
    CHECK(I_Borders_get_Count(borders, NULL) == E_POINTER);
    CHECK(I_Borders_get_Item(borders, xlEdgeLeft, NULL) == E_POINTER);
    CHECK(I_Borders_get__Default(borders, xlEdgeLeft, NULL) == E_POINTER);
    first = enumerator(borders);
    second = enumerator(borders);
    REQUIRE(first && second);
    CHECK(I_Borders_GetEnumerator(borders, NULL) == E_POINTER);
    CHECK(IEnumVARIANT_Clone(first, NULL) == E_POINTER);
    CHECK(I_Borders_QueryInterface(borders, &IID_IUnknown, NULL) == E_POINTER);
    next(first, 3, 3, S_OK);
    next(second, count + 1, count, S_FALSE); /* Independent position. */
    CHECK(IEnumVARIANT_Skip(first, count - 3) == S_OK); /* Exactly to end. */
    next(first, 1, 0, S_FALSE);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, count + 1) == S_FALSE);
    next(first, 1, 0, S_FALSE);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, ~(ULONG)0) == S_FALSE);
    next(first, 1, 0, S_FALSE);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    next(first, count + 1, count, S_FALSE);
    next(first, 1, 0, S_FALSE);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, 2) == S_OK);
    hr = IEnumVARIANT_Clone(first, &clone);
    CHECK(hr == S_OK && clone);
    if (clone) {
        CHECK(IEnumVARIANT_Skip(first, 1) == S_OK);
        next(clone, count + 1, count - 2, S_FALSE);
        next(first, count + 1, count - 3, S_FALSE);
        IEnumVARIANT_Release(clone);
        clone = NULL;
        CHECK(IEnumVARIANT_Clone(first, &clone) == S_OK && clone);
        if (clone) next(clone, 1, 0, S_FALSE);
    }
    /* Enumerators retain their range after releasing the original collection. */
    I_Borders_Release(borders);
    borders = NULL;
    VariantClear(&collection);
    VariantClear(&range);
    CHECK(IEnumVARIANT_Reset(second) == S_OK);
    next(second, count + 1, count, S_FALSE);
done:
    if (clone) IEnumVARIANT_Release(clone);
    if (second) IEnumVARIANT_Release(second);
    if (first) IEnumVARIANT_Release(first);
    if (borders) I_Borders_Release(borders);
    VariantClear(&collection);
    VariantClear(&range);
    VariantClear(&argument);
    if (V_VT(&workbook) == VT_DISPATCH) {
        V_VT(&argument) = VT_BOOL;
        V_BOOL(&argument) = VARIANT_FALSE;
        CHECK(SUCCEEDED(invoke(V_DISPATCH(&workbook), L"Close", DISPATCH_METHOD, &argument, 1, NULL)));
    }
    VariantClear(&workbook);
    VariantClear(&books);
    if (excel) {
        CHECK(invoke(excel, L"Quit", DISPATCH_METHOD, NULL, 0, NULL) == S_OK);
        IDispatch_Release(excel);
    }
    CoUninitialize();
    printf("Borders enumerator: %d failure(s)\n", failures);
    return failures ? 1 : 0;
}
