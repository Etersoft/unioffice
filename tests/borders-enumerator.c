/* COM contracts: Borders metadata/enumeration and Range/Cells dispatch. */
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
#define CHECK_HR(expression, expected) do { HRESULT actual = (expression); \
    if (actual != (expected)) { \
        fprintf(stderr, "FAIL line %d: %s returned 0x%08lx, expected 0x%08lx\n", \
                __LINE__, #expression, (unsigned long)actual, (unsigned long)(expected)); \
        ++failures; \
    } \
} while (0)

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
    const XlLineStyle unsupported[] = {xlDash, xlDashDot, xlDashDotDot, xlDot, xlSlantDashDot};
    unsigned int index;
    IDispatch *dispatch = NULL;
    I_Border *border = NULL;
    CHECK(I_Borders_get_LineStyle(borders, NULL) == E_POINTER);
    CHECK(I_Borders_get_Value(borders, NULL) == E_POINTER);
    for (index = 0; index < sizeof(unsupported) / sizeof(unsupported[0]); ++index) {
        CHECK(I_Borders_put_LineStyle(borders, unsupported[index]) == E_NOTIMPL);
        CHECK(I_Borders_put_Value(borders, unsupported[index]) == E_NOTIMPL);
    }
    CHECK(I_Borders_put_LineStyle(borders, (XlLineStyle)12345) == E_INVALIDARG);
    CHECK(I_Borders_put_Value(borders, (XlLineStyle)12345) == E_INVALIDARG);
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

static void mixed_weight_contract(I_Borders *borders)
{
    struct { LONG before; XlBorderWeight weight; LONG after; } legacy = {12345, xlHairline, 67890};
    ITypeInfo *info = NULL, *invalid = (ITypeInfo *)1;
    TYPEATTR *attributes = NULL;
    FUNCDESC *function;
    IDispatch *dispatch = NULL;
    I_Border *border = NULL;
    IUnknown *automation = NULL;
    VARIANT result;
    UINT index;
    BOOL found = FALSE;
    CHECK(I_Borders_GetTypeInfo(borders, 1, 0, &invalid) == DISP_E_BADINDEX && !invalid);
    CHECK(I_Borders_GetTypeInfo(borders, 0, 0, NULL) == E_POINTER);
    CHECK(I_Borders_GetTypeInfoCount(borders, NULL) == E_POINTER);
    CHECK(I_Borders_QueryInterface(borders, &DIID_D_Borders, (void **)&automation) == S_OK && automation);
    if (automation) IUnknown_Release(automation);
    CHECK(I_Borders_GetTypeInfo(borders, 0, 0, &info) == S_OK && info);
    if (info) {
        CHECK(ITypeInfo_GetTypeAttr(info, &attributes) == S_OK);
        if (attributes) {
            CHECK(IsEqualGUID(&attributes->guid, &DIID_D_Borders));
            CHECK(attributes->typekind == TKIND_DISPATCH);
            CHECK(!(attributes->wTypeFlags & TYPEFLAG_FDUAL));
            for (index = 0; index < attributes->cFuncs; ++index) {
                if (FAILED(ITypeInfo_GetFuncDesc(info, index, &function))) { CHECK(0); continue; }
                if (function->memid == 11 && function->invkind == INVOKE_PROPERTYGET) {
                    found = TRUE;
                    CHECK(function->elemdescFunc.tdesc.vt == VT_VARIANT);
                }
                ITypeInfo_ReleaseFuncDesc(info, function);
            }
            CHECK(found);
            ITypeInfo_ReleaseTypeAttr(info, attributes);
        }
        ITypeInfo_Release(info);
    }
    CHECK_HR(I_Borders_put_Weight(borders, xlThin), S_OK);
    VariantInit(&result);
    CHECK_HR(invoke((IDispatch *)borders, L"Weight", DISPATCH_PROPERTYGET, NULL, 0, &result), S_OK);
    CHECK(V_VT(&result) == VT_I4 && V_I4(&result) == xlThin);
    VariantClear(&result);
    CHECK(I_Borders_get_Item(borders, xlEdgeLeft, &dispatch) == S_OK && dispatch);
    if (dispatch) {
        CHECK(IDispatch_QueryInterface(dispatch, &IID_I_Border, (void **)&border) == S_OK);
        if (border) {
            CHECK(I_Border_put_Weight(border, xlThick) == S_OK);
            I_Border_Release(border);
        }
        IDispatch_Release(dispatch);
    }
    CHECK(invoke((IDispatch *)borders, L"Weight", DISPATCH_PROPERTYGET, NULL, 0, &result) == S_OK);
    CHECK(V_VT(&result) == VT_NULL);
    VariantClear(&result);
    /* This unchanged legacy slot must still write one enum, even for mixed borders. */
    CHECK(I_Borders_get_Weight(borders, &legacy.weight) == S_OK);
    CHECK(legacy.weight == xlThin && legacy.before == 12345 && legacy.after == 67890);
    CHECK(I_Borders_put_Weight(borders, xlThin) == S_OK);
}

static void range_argument_contract(IDispatch *dispatch)
{
    I_Range *range = NULL;
    I_Worksheet *worksheet = NULL;
    IDispatch *parent = NULL, *result = (IDispatch *)1;
    VARIANT row, column, value, missing, reference, arguments[3];
    DOUBLE original_row = 2;
    VariantInit(&row);
    VariantInit(&column);
    V_VT(&row) = VT_I4;
    V_I4(&row) = 2;
    V_VT(&column) = VT_BSTR;
    V_BSTR(&column) = SysAllocString(L"invalid-column");
    CHECK(IDispatch_QueryInterface(dispatch, &IID_I_Range, (void **)&range) == S_OK && range);
    if (!range) goto done;
    CHECK(I_Range_get__Default(range, row, column, NULL) == E_POINTER);
    CHECK(I_Range_get__Default(range, row, column, &result) == DISP_E_TYPEMISMATCH);
    CHECK(result == NULL);
    CHECK(V_VT(&column) == VT_BSTR && !lstrcmpW(V_BSTR(&column), L"invalid-column"));
    VariantInit(&value);
    VariantInit(&missing);
    VariantInit(&reference);
    V_VT(&missing) = VT_ERROR;
    V_ERROR(&missing) = DISP_E_PARAMNOTFOUND;
    CHECK(I_Range_get_Item(range, row, column, 0, NULL) == E_POINTER);
    V_VT(&value) = VT_I4;
    V_I4(&value) = 123;
    CHECK(I_Range_get_Item(range, row, column, 0, &value) == DISP_E_TYPEMISMATCH);
    CHECK(V_VT(&value) == VT_EMPTY);
    CHECK(I_Range_put_Item(range, row, column, 0, row) == DISP_E_TYPEMISMATCH);
    CHECK(V_VT(&column) == VT_BSTR && !lstrcmpW(V_BSTR(&column), L"invalid-column"));
    VariantClear(&column);
    V_VT(&column) = VT_BSTR;
    V_BSTR(&column) = SysAllocString(L"AA");
    V_VT(&reference) = VT_R8 | VT_BYREF;
    V_R8REF(&reference) = &original_row;
    CHECK(I_Range_get_Item(range, reference, column, 0, &value) == S_OK);
    CHECK(V_VT(&value) == VT_DISPATCH && V_DISPATCH(&value));
    VariantClear(&value);
    CHECK(original_row == 2 && !lstrcmpW(V_BSTR(&column), L"AA"));
    CHECK(I_Range_get_Item(range, reference, missing, 0, &value) == S_OK);
    CHECK(V_VT(&value) == VT_DISPATCH && V_DISPATCH(&value));
    VariantClear(&value);
    CHECK(I_Range_get__Default(range, reference, missing, &result) == S_OK && result);
    if (result) IDispatch_Release(result);
    CHECK(I_Range_put_Item(range, reference, column, 0, row) == S_OK);
    CHECK(I_Range_put__Default(range, reference, column, 0, row) == S_OK);
    CHECK(original_row == 2 && !lstrcmpW(V_BSTR(&column), L"AA"));
    arguments[0] = column;
    arguments[1] = reference;
    arguments[2] = row;
    CHECK(invoke(dispatch, L"Item", DISPATCH_PROPERTYGET, arguments, 0, &value) == DISP_E_BADPARAMCOUNT);
    CHECK(invoke(dispatch, L"Item", DISPATCH_PROPERTYGET, arguments, 3, &value) == DISP_E_BADPARAMCOUNT);
    CHECK(invoke(dispatch, L"Item", DISPATCH_PROPERTYGET, arguments, 2, NULL) == E_POINTER);
    CHECK(invoke(dispatch, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, arguments, 2, &value) == S_OK);
    CHECK(V_VT(&value) == VT_DISPATCH && V_DISPATCH(&value));
    VariantClear(&value);
    VariantClear(&column);
    V_VT(&column) = VT_BSTR;
    V_BSTR(&column) = SysAllocString(L"ZZZZZZZZZZZZZZZZ");
    CHECK(I_Range_get_Item(range, row, column, 0, &value) == DISP_E_OVERFLOW);
    CHECK(V_VT(&value) == VT_EMPTY);
    CHECK(I_Range_get_Parent(range, &parent) == S_OK && parent);
    if (!parent) goto done;
    CHECK(IDispatch_QueryInterface(parent, &IID_I_Worksheet, (void **)&worksheet) == S_OK && worksheet);
    if (!worksheet) goto done;
    CHECK(I_Worksheet_get_Cells(worksheet, NULL) == E_POINTER);
    CHECK(I_Worksheet_get_Range(worksheet, row, row, NULL) == E_POINTER);
    result = (IDispatch *)1;
    CHECK(I_Worksheet_get_Range(worksheet, row, row, &result) == E_INVALIDARG);
    CHECK(result == NULL);
done:
    if (worksheet) I_Worksheet_Release(worksheet);
    if (parent) IDispatch_Release(parent);
    if (range) I_Range_Release(range);
    VariantClear(&column);
    VariantClear(&row);
}

/* The supplied source is A1:D4. Exercise native slots as well as Invoke. */
static void range_collection_contract(IDispatch *dispatch)
{
    I_Range *range = NULL, *view = NULL, *selected = NULL;
    IDispatch *collection = NULL, *cells = NULL;
    VARIANT index, missing, result, arguments[2];
    DOUBLE offset = 2;
    LONG count = 0;
    UINT kind;
    VariantInit(&index);
    VariantInit(&missing);
    VariantInit(&result);
    V_VT(&missing) = VT_ERROR;
    V_ERROR(&missing) = DISP_E_PARAMNOTFOUND;
    CHECK(IDispatch_QueryInterface(dispatch, &IID_I_Range, (void**)&range) == S_OK);
    if (!range) return;
    CHECK(I_Range_get_Rows(range, NULL) == E_POINTER);
    CHECK(I_Range_get_Columns(range, NULL) == E_POINTER);
    CHECK(I_Range_get_Cells(range, NULL) == E_POINTER);
    CHECK(I_Range_get_Count(range, NULL) == E_POINTER);
    for (kind = 0; kind < 2; ++kind) {
        if (kind == 0) CHECK(I_Range_get_Rows(range, &collection) == S_OK);
        else CHECK(I_Range_get_Columns(range, &collection) == S_OK);
        if (!collection) continue;
        CHECK(IDispatch_QueryInterface(collection, &IID_I_Range, (void**)&view) == S_OK);
        if (!view) { IDispatch_Release(collection); collection = NULL; continue; }
        CHECK(I_Range_get_Count(view, &count) == S_OK && count == 4);
        V_VT(&index) = VT_R8 | VT_BYREF;
        V_R8REF(&index) = &offset;
        CHECK_HR(I_Range_get_Item(view, index, missing, 0, &result), S_OK);
        CHECK(offset == 2);
        if (V_VT(&result) == VT_DISPATCH && V_DISPATCH(&result)) {
            CHECK(IDispatch_QueryInterface(V_DISPATCH(&result), &IID_I_Range, (void**)&selected) == S_OK);
            if (selected) {
                CHECK(I_Range_get_Row(selected) == (kind == 0 ? 2 : 1));
                CHECK(I_Range_get_Column(selected) == (kind == 0 ? 1 : 2));
                CHECK(I_Range_get_Count(selected, &count) == S_OK && count == 1);
                CHECK(I_Range_get_Cells(selected, &cells) == S_OK && cells);
                I_Range_Release(selected);
                selected = NULL;
                if (cells) {
                    CHECK(IDispatch_QueryInterface(cells, &IID_I_Range, (void**)&selected) == S_OK);
                    if (selected) {
                        CHECK(I_Range_get_Count(selected, &count) == S_OK && count == 4);
                        I_Range_Release(selected);
                        selected = NULL;
                    }
                    IDispatch_Release(cells);
                    cells = NULL;
                }
            }
        } else CHECK(FALSE);
        VariantClear(&result);
        CHECK_HR(I_Range_get_Item(view, index, index, 0, &result), E_INVALIDARG);
        CHECK(V_VT(&result) == VT_EMPTY && offset == 2);
        V_VT(&index) = VT_BSTR;
        V_BSTR(&index) = SysAllocString(kind == 0 ? L"bad-row" : L"ZZZZZZZZZZZZZZ");
        CHECK_HR(I_Range_get_Item(view, index, missing, 0, &result),
                kind == 0 ? DISP_E_TYPEMISMATCH : DISP_E_OVERFLOW);
        CHECK(V_VT(&result) == VT_EMPTY && V_VT(&index) == VT_BSTR);
        VariantClear(&index);
        V_VT(&index) = VT_I4;
        V_I4(&index) = 0;
        CHECK_HR(I_Range_get_Item(view, index, missing, 0, &result), E_INVALIDARG);
        CHECK(V_VT(&result) == VT_EMPTY);
        arguments[0] = arguments[1] = index;
        CHECK_HR(invoke(dispatch, kind == 0 ? L"Rows" : L"Columns",
                DISPATCH_PROPERTYGET, arguments, 2, &result), E_INVALIDARG);
        CHECK(V_VT(&result) == VT_EMPTY);
        CHECK_HR(invoke(dispatch, kind == 0 ? L"Rows" : L"Columns",
                DISPATCH_METHOD | DISPATCH_PROPERTYGET, NULL, 0, NULL), S_OK);
        I_Range_Release(view);
        view = NULL;
        IDispatch_Release(collection);
        collection = NULL;
    }
    CHECK(I_Range_get_Count(range, &count) == S_OK && count == 16);
    VariantClear(&index);
    I_Range_Release(range);
}

static void cells_dispatch_contract(IDispatch *owner)
{
    LPOLESTR name = L"Cells";
    DISPID id, put = DISPID_PROPERTYPUT;
    VARIANT arguments[4], result;
    DISPPARAMS params = {arguments, NULL, 0, 0};
    DOUBLE index = 4;
    I_Range *cell = NULL;
    UINT item;
    for (item = 0; item < 4; ++item) VariantInit(&arguments[item]);
    VariantInit(&result);
    CHECK_HR(IDispatch_GetIDsOfNames(owner, &IID_NULL, &name, 1, 0, &id), S_OK);
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYGET,
                             NULL, &result, NULL, NULL), E_POINTER);
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYGET,
                             &params, NULL, NULL, NULL), S_OK);
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYGET,
                             &params, &result, NULL, NULL), S_OK);
    CHECK(V_VT(&result) == VT_DISPATCH && V_DISPATCH(&result));
    VariantClear(&result);
    V_VT(&arguments[0]) = VT_R8 | VT_BYREF;
    V_R8REF(&arguments[0]) = &index;
    params.cArgs = 1;
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                             &params, &result, NULL, NULL), S_OK);
    if (V_VT(&result) == VT_DISPATCH && V_DISPATCH(&result)) {
        CHECK_HR(IDispatch_QueryInterface(V_DISPATCH(&result), &IID_I_Range, (void **)&cell), S_OK);
        if (cell) {
            CHECK(I_Range_get_Row(cell) == 1 && I_Range_get_Column(cell) == 4);
            I_Range_Release(cell);
        }
    } else CHECK(0);
    VariantClear(&result);
    CHECK(index == 4);
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYGET,
                             &params, NULL, NULL, NULL), S_OK);
    arguments[1] = arguments[0];
    V_VT(&arguments[0]) = VT_BSTR;
    V_BSTR(&arguments[0]) = SysAllocString(L"native-owned-value");
    params.cArgs = 2;
    params.cNamedArgs = 1;
    params.rgdispidNamedArgs = &put;
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYPUT,
                             &params, NULL, NULL, NULL), S_OK);
    CHECK(index == 4 && !lstrcmpW(V_BSTR(&arguments[0]), L"native-owned-value"));
    VariantClear(&arguments[0]);
    V_VT(&arguments[0]) = VT_BSTR;
    V_BSTR(&arguments[0]) = SysAllocString(L"invalid-index");
    params.cArgs = 1;
    params.cNamedArgs = 0;
    params.rgdispidNamedArgs = NULL;
    V_VT(&result) = VT_I4;
    V_I4(&result) = 123;
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYGET,
                             &params, &result, NULL, NULL), DISP_E_TYPEMISMATCH);
    CHECK(V_VT(&result) == VT_EMPTY);
    CHECK(!lstrcmpW(V_BSTR(&arguments[0]), L"invalid-index"));
    params.cArgs = 3;
    V_VT(&result) = VT_I4;
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYGET,
                             &params, &result, NULL, NULL), DISP_E_BADPARAMCOUNT);
    CHECK(V_VT(&result) == VT_EMPTY);
    params.cArgs = 1;
    params.cNamedArgs = 1;
    params.rgdispidNamedArgs = &put;
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYPUT,
                             &params, NULL, NULL, NULL), DISP_E_BADPARAMCOUNT);
    params.cArgs = 4;
    CHECK_HR(IDispatch_Invoke(owner, id, &IID_NULL, 0, DISPATCH_PROPERTYPUT,
                             &params, NULL, NULL, NULL), DISP_E_BADPARAMCOUNT);
    VariantClear(&result);
    for (item = 0; item < 4; ++item) VariantClear(&arguments[item]);
}

int __cdecl main(void)
{
    CLSID clsid;
    IDispatch *excel = NULL;
    I_Borders *borders = NULL;
    IEnumVARIANT *first = NULL, *second = NULL, *clone = NULL;
    VARIANT books, workbook, range, collection, argument, sheet;
    LONG count = 0;
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return 2;
    VariantInit(&books);
    VariantInit(&workbook);
    VariantInit(&range);
    VariantInit(&collection);
    VariantInit(&argument);
    VariantInit(&sheet);
    REQUIRE(CLSIDFromProgID(L"Excel.Application", &clsid) == S_OK);
    REQUIRE(CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER, &IID_IDispatch,
                             (void **)&excel) == S_OK);
    REQUIRE(invoke(excel, L"Workbooks", DISPATCH_PROPERTYGET, NULL, 0, &books) == S_OK);
    REQUIRE(invoke(V_DISPATCH(&books), L"Add", DISPATCH_METHOD, NULL, 0, &workbook) == S_OK);
    REQUIRE(invoke(excel, L"ActiveSheet", DISPATCH_PROPERTYGET, NULL, 0, &sheet) == S_OK);
    cells_dispatch_contract(excel);
    cells_dispatch_contract(V_DISPATCH(&sheet));
    VariantClear(&sheet);
    V_VT(&argument) = VT_BSTR;
    V_BSTR(&argument) = SysAllocString(L"A1:D4");
    REQUIRE(invoke(excel, L"Range", DISPATCH_PROPERTYGET, &argument, 1, &range) == S_OK);
    VariantClear(&argument);
    range_argument_contract(V_DISPATCH(&range));
    range_collection_contract(V_DISPATCH(&range));
    REQUIRE(invoke(V_DISPATCH(&range), L"Borders", DISPATCH_PROPERTYGET, NULL, 0, &collection) == S_OK);
    REQUIRE(IDispatch_QueryInterface(V_DISPATCH(&collection), &IID_I_Borders, (void **)&borders) == S_OK);
    REQUIRE(I_Borders_get_Count(borders, &count) == S_OK && count > 3 && count < 65);
    CHECK(count == 8);
    invalid_indices(borders);
    weight_contract(borders);
    mixed_weight_contract(borders);
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
    VariantClear(&sheet);
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
