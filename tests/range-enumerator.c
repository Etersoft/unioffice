/* Native IEnumVARIANT contracts, separate from VBScript For Each integration. */
#define COBJMACROS
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "unioffice_excel.h"

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

static IEnumVARIANT *enumerator(I_Range *range)
{
    IUnknown *unknown = NULL;
    IEnumVARIANT *result = NULL;
    CHECK(I_Range_get__NewEnum(range, &unknown) == S_OK && unknown);
    if (unknown) {
        CHECK(IUnknown_QueryInterface(unknown, &IID_IEnumVARIANT, (void**)&result) == S_OK);
        IUnknown_Release(unknown);
    }
    return result;
}

static void shape(VARIANT *value, LONG row, LONG column, LONG count)
{
    I_Range *range = NULL;
    LONG actual = -1;
    CHECK(V_VT(value) == VT_DISPATCH && V_DISPATCH(value));
    if (V_VT(value) != VT_DISPATCH || !V_DISPATCH(value)) return;
    CHECK(IDispatch_QueryInterface(V_DISPATCH(value), &IID_I_Range, (void**)&range) == S_OK);
    if (!range) return;
    CHECK(I_Range_get_Row(range) == row);
    CHECK(I_Range_get_Column(range) == column);
    CHECK(I_Range_get_Count(range, &actual) == S_OK && actual == count);
    I_Range_Release(range);
}

static void next_one(IEnumVARIANT *cursor, LONG row, LONG column)
{
    VARIANT value;
    VariantInit(&value);
    CHECK(IEnumVARIANT_Next(cursor, 1, &value, NULL) == S_OK);
    shape(&value, row, column, 1);
    VariantClear(&value);
}

static void collection(I_Range *range, int kind)
{
    IEnumVARIANT *first = NULL, *second = NULL, *copy = NULL;
    IUnknown *identity = NULL, *other = NULL;
    IDispatch *unsupported = (IDispatch*)1;
    VARIANT values[8], result, argument;
    ULONG fetched = 99, i, total = kind == 0 ? 6 : kind == 1 ? 2 : 3;
    LONG row, column;
    DISPPARAMS params = {NULL, NULL, 0, 0};
    for (i = 0; i < 8; ++i) VariantInit(&values[i]);
    VariantInit(&result);
    VariantInit(&argument);
    CHECK(I_Range_get__NewEnum(range, NULL) == E_POINTER);
    first = enumerator(range);
    second = enumerator(range);
    if (!first || !second) goto done;
    CHECK(first != second);
    CHECK(IEnumVARIANT_QueryInterface(first, &IID_IUnknown, (void**)&identity) == S_OK);
    CHECK(identity == (IUnknown*)first);
    CHECK(IUnknown_QueryInterface(identity, &IID_IUnknown, (void**)&other) == S_OK && other == identity);
    CHECK(IEnumVARIANT_QueryInterface(first, &IID_IDispatch, (void**)&unsupported) == E_NOINTERFACE);
    CHECK(!unsupported);
    CHECK(IEnumVARIANT_QueryInterface(first, &IID_IUnknown, NULL) == E_POINTER);
    CHECK(IEnumVARIANT_Clone(first, NULL) == E_POINTER);
    CHECK(IEnumVARIANT_Next(first, 1, NULL, &fetched) == E_INVALIDARG && fetched == 0);
    CHECK(IEnumVARIANT_Next(first, 0, values, &fetched) == S_OK && fetched == 0);
    CHECK(IEnumVARIANT_Skip(first, 0) == S_OK);
    CHECK(IEnumVARIANT_Next(first, 1, values, &fetched) == S_OK && fetched == 1);
    shape(values, 2, 2, 1);
    VariantClear(values);
    CHECK(IEnumVARIANT_Clone(first, &copy) == S_OK && copy);
    CHECK(IEnumVARIANT_Next(second, 1, values, NULL) == S_OK);
    shape(values, 2, 2, 1);
    VariantClear(values);
    row = kind == 1 ? 3 : 2;
    column = kind == 1 ? 2 : 3;
    if (copy) next_one(copy, row, column);
    next_one(first, row, column);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    CHECK(IEnumVARIANT_Next(first, total + 1, values, &fetched) == S_FALSE && fetched == total);
    for (i = 0; i < total; ++i) {
        row = kind == 0 ? 2 + i / 3 : kind == 1 ? 2 + i : 2;
        column = kind == 0 ? 2 + i % 3 : kind == 1 ? 2 : 2 + i;
        shape(&values[i], row, column, 1);
        VariantClear(&values[i]);
    }
    CHECK(V_VT(&values[total]) == VT_EMPTY);
    CHECK(IEnumVARIANT_Next(first, 1, values, &fetched) == S_FALSE && fetched == 0);
    CHECK(IEnumVARIANT_Skip(first, 0) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, 1) == S_FALSE);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, total) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, 1) == S_FALSE);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, 0xffffffffUL) == S_FALSE);
    CHECK(IEnumVARIANT_Next(first, 1, values, &fetched) == S_FALSE && fetched == 0);
    CHECK(IEnumVARIANT_Reset(first) == S_OK);
    CHECK(IEnumVARIANT_Skip(first, total - 1) == S_OK);
    CHECK(IEnumVARIANT_Next(first, 2, values, NULL) == S_FALSE);
    CHECK(V_VT(&values[0]) == VT_DISPATCH && V_VT(&values[1]) == VT_EMPTY);
    VariantClear(values);
    CHECK(invoke((IDispatch*)range, L"_NewEnum", DISPATCH_PROPERTYGET, NULL, 0, &result) == S_OK);
    CHECK(V_VT(&result) == VT_UNKNOWN && V_UNKNOWN(&result));
    VariantClear(&result);
    CHECK(invoke((IDispatch*)range, L"_NewEnum", DISPATCH_METHOD | DISPATCH_PROPERTYGET, NULL, 0, NULL) == S_OK);
    CHECK(invoke((IDispatch*)range, L"_NewEnum", DISPATCH_PROPERTYGET, &argument, 1, &result) == DISP_E_BADPARAMCOUNT);
    CHECK(V_VT(&result) == VT_EMPTY);
    CHECK(invoke((IDispatch*)range, L"_NewEnum", DISPATCH_PROPERTYPUT, NULL, 0, &result) == DISP_E_MEMBERNOTFOUND);
    CHECK(I_Range_Invoke(range, DISPID_NEWENUM, &IID_NULL, 0, DISPATCH_PROPERTYGET,
            NULL, &result, NULL, NULL) == E_POINTER);
    CHECK(I_Range_Invoke(range, DISPID_NEWENUM, &IID_NULL, 0, DISPATCH_PROPERTYGET,
            &params, &result, NULL, NULL) == S_OK);
done:
    VariantClear(&result);
    for (i = 0; i < 8; ++i) VariantClear(&values[i]);
    if (other) IUnknown_Release(other);
    if (identity) IUnknown_Release(identity);
    if (copy) IEnumVARIANT_Release(copy);
    if (second) IEnumVARIANT_Release(second);
    if (first) IEnumVARIANT_Release(first);
}

static void last_sheet_cell(IDispatch *sheet)
{
    VARIANT cells, value;
    IDispatch *rows = NULL, *columns = NULL;
    IEnumVARIANT *cursor = NULL;
    I_Range *range = NULL;
    LONG height = 0, width = 0;
    ULONGLONG remaining;
    ULONG step;
    VariantInit(&cells);
    VariantInit(&value);
    CHECK(invoke(sheet, L"Cells", DISPATCH_PROPERTYGET, NULL, 0, &cells) == S_OK);
    if (V_VT(&cells) != VT_DISPATCH) goto done;
    CHECK(IDispatch_QueryInterface(V_DISPATCH(&cells), &IID_I_Range, (void**)&range) == S_OK);
    if (!range) goto done;
    CHECK(I_Range_get_Rows(range, &rows) == S_OK && rows);
    CHECK(I_Range_get_Columns(range, &columns) == S_OK && columns);
    if (!rows || !columns) goto done;
    CHECK(I_Range_get_Count((I_Range*)rows, &height) == S_OK && height > 0);
    CHECK(I_Range_get_Count((I_Range*)columns, &width) == S_OK && width > 0);
    if (height < 1 || width < 1) goto done;
    cursor = enumerator(range);
    if (!cursor) goto done;
    remaining = (ULONGLONG)height * width - 1;
    while (remaining) {
        step = remaining > 0xffffffffUL ? 0xffffffffUL : (ULONG)remaining;
        CHECK(IEnumVARIANT_Skip(cursor, step) == S_OK);
        remaining -= step;
    }
    CHECK(IEnumVARIANT_Next(cursor, 1, &value, NULL) == S_OK);
    shape(&value, height, width, 1);
    VariantClear(&value);
    CHECK(IEnumVARIANT_Next(cursor, 1, &value, NULL) == S_FALSE);
done:
    VariantClear(&value);
    if (cursor) IEnumVARIANT_Release(cursor);
    if (columns) IDispatch_Release(columns);
    if (rows) IDispatch_Release(rows);
    if (range) I_Range_Release(range);
    VariantClear(&cells);
}

int __cdecl main(void)
{
    CLSID clsid;
    IDispatch *excel = NULL;
    I_Range *range = NULL;
    IEnumVARIANT *retained = NULL, *copy = NULL;
    VARIANT books, book, sheet, source, view, address, value;
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return 2;
    VariantInit(&books); VariantInit(&book); VariantInit(&sheet);
    VariantInit(&source); VariantInit(&view); VariantInit(&address); VariantInit(&value);
    CHECK(CLSIDFromProgID(L"Excel.Application", &clsid) == S_OK);
    hr = CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER, &IID_IDispatch, (void**)&excel);
    CHECK(hr == S_OK && excel);
    if (!excel) goto done;
    CHECK(invoke(excel, L"Workbooks", DISPATCH_PROPERTYGET, NULL, 0, &books) == S_OK);
    if (V_VT(&books) != VT_DISPATCH) goto done;
    CHECK(invoke(V_DISPATCH(&books), L"Add", DISPATCH_METHOD, NULL, 0, &book) == S_OK);
    CHECK(invoke(excel, L"ActiveSheet", DISPATCH_PROPERTYGET, NULL, 0, &sheet) == S_OK);
    if (V_VT(&sheet) != VT_DISPATCH) goto done;
    last_sheet_cell(V_DISPATCH(&sheet));
    V_VT(&address) = VT_BSTR;
    V_BSTR(&address) = SysAllocString(L"B2:D3");
    CHECK(invoke(V_DISPATCH(&sheet), L"Range", DISPATCH_PROPERTYGET, &address, 1, &source) == S_OK);
    if (V_VT(&source) != VT_DISPATCH) goto done;
    CHECK(IDispatch_QueryInterface(V_DISPATCH(&source), &IID_I_Range, (void**)&range) == S_OK);
    if (!range) goto done;
    collection(range, 0);
    retained = enumerator(range);
    CHECK(I_Range_get_Rows(range, &V_DISPATCH(&view)) == S_OK);
    if (V_DISPATCH(&view)) { V_VT(&view) = VT_DISPATCH; collection((I_Range*)V_DISPATCH(&view), 1); }
    VariantClear(&view);
    CHECK(I_Range_get_Columns(range, &V_DISPATCH(&view)) == S_OK);
    if (V_DISPATCH(&view)) { V_VT(&view) = VT_DISPATCH; collection((I_Range*)V_DISPATCH(&view), 2); }
    VariantClear(&view);
    I_Range_Release(range); range = NULL;
    VariantClear(&source); VariantClear(&sheet);
    if (retained) {
        next_one(retained, 2, 2);
        CHECK(IEnumVARIANT_Clone(retained, &copy) == S_OK && copy);
        IEnumVARIANT_Release(retained); retained = NULL;
        if (copy) {
            CHECK(IEnumVARIANT_Next(copy, 1, &value, NULL) == S_OK);
            shape(&value, 2, 3, 1);
            IEnumVARIANT_Release(copy); copy = NULL;
            /* The yielded Range must outlive the last enumerator reference. */
            shape(&value, 2, 3, 1);
        }
    }
done:
    VariantClear(&value);
    if (copy) IEnumVARIANT_Release(copy);
    if (retained) IEnumVARIANT_Release(retained);
    if (range) I_Range_Release(range);
    VariantClear(&address); VariantClear(&view); VariantClear(&source); VariantClear(&sheet);
    VariantClear(&book); VariantClear(&books);
    if (excel) { invoke(excel, L"Quit", DISPATCH_METHOD, NULL, 0, NULL); IDispatch_Release(excel); }
    CoUninitialize();
    printf("Range enumerator: %d failure(s)\n", failures);
    return failures ? 1 : 0;
}
