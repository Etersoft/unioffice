/*
 * IWindows and IWindow interface functions
 *
 * Copyright (C) 2008 Sinitsin Ivan (Etersoft) <ivan@etersoft.ru>
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301, USA
 */

#include "mso_to_oo_private.h"
#include "special_functions.h"

/* A Window is the view of one workbook; Windows lists the application's
   workbook windows, or only one workbook's window. */
#define WINDOWS_THIS(iface) ((WindowsImpl*)(iface))
#define WINDOW_THIS(iface) ((WindowImpl*)(iface))

ITypeInfo *ti_windows = NULL;
ITypeInfo *ti_window = NULL;

static HRESULT load_typeinfo(REFIID iid, ITypeInfo **cache, ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (*cache) {
        *typeinfo = *cache;
        return S_OK;
    }
    hres = LoadTypeLib(file_name, &typelib);
    if (FAILED(hres)) return hres;
    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, iid, cache);
    typelib->lpVtbl->Release(typelib);
    *typeinfo = *cache;
    return hres;
}

HRESULT get_typeinfo_windows(ITypeInfo **typeinfo)
{
    return load_typeinfo(&IID_I_Windows, &ti_windows, typeinfo);
}

HRESULT get_typeinfo_window(ITypeInfo **typeinfo)
{
    return load_typeinfo(&IID_I_Window, &ti_window, typeinfo);
}

/* Snapshot enumerator over Window objects. */
typedef struct
{
    const IEnumVARIANTVtbl *lpVtbl;
    LONG ref;
    LONG position;
    LONG count;
    IDispatch *items[1];
} WindowsEnum;

static const IEnumVARIANTVtbl windows_enum_vtbl;

static HRESULT WINAPI windows_enum_QueryInterface(IEnumVARIANT *iface, REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualGUID(riid, &IID_IUnknown) || IsEqualGUID(riid, &IID_IEnumVARIANT)) {
        *ppv = iface;
        IEnumVARIANT_AddRef(iface);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG WINAPI windows_enum_AddRef(IEnumVARIANT *iface)
{
    return InterlockedIncrement(&((WindowsEnum*)iface)->ref);
}

static ULONG WINAPI windows_enum_Release(IEnumVARIANT *iface)
{
    WindowsEnum *This = (WindowsEnum*)iface;
    ULONG ref = InterlockedDecrement(&This->ref);
    LONG i;
    if (!ref) {
        for (i = 0; i < This->count; i++)
            if (This->items[i]) IDispatch_Release(This->items[i]);
        HeapFree(GetProcessHeap(), 0, This);
        InterlockedDecrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI windows_enum_Next(IEnumVARIANT *iface, ULONG celt, VARIANT *rgVar, ULONG *fetched)
{
    WindowsEnum *This = (WindowsEnum*)iface;
    ULONG n = 0;
    if (!rgVar) return E_POINTER;
    while (n < celt && This->position < This->count) {
        V_VT(&rgVar[n]) = VT_DISPATCH;
        V_DISPATCH(&rgVar[n]) = This->items[This->position++];
        IDispatch_AddRef(V_DISPATCH(&rgVar[n]));
        n++;
    }
    if (fetched) *fetched = n;
    return n == celt ? S_OK : S_FALSE;
}

static HRESULT WINAPI windows_enum_Skip(IEnumVARIANT *iface, ULONG celt)
{
    WindowsEnum *This = (WindowsEnum*)iface;
    LONG left = This->count - This->position;
    if ((LONG)celt > left) {
        This->position = This->count;
        return S_FALSE;
    }
    This->position += celt;
    return S_OK;
}

static HRESULT WINAPI windows_enum_Reset(IEnumVARIANT *iface)
{
    ((WindowsEnum*)iface)->position = 0;
    return S_OK;
}

static HRESULT WINAPI windows_enum_Clone(IEnumVARIANT *iface, IEnumVARIANT **ppEnum)
{
    WindowsEnum *This = (WindowsEnum*)iface, *copy;
    LONG i;
    if (!ppEnum) return E_POINTER;
    *ppEnum = NULL;
    copy = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(*copy) + This->count * sizeof(IDispatch*));
    if (!copy) return E_OUTOFMEMORY;
    *copy = *This;
    copy->ref = 1;
    for (i = 0; i < This->count; i++) {
        copy->items[i] = This->items[i];
        IDispatch_AddRef(copy->items[i]);
    }
    InterlockedIncrement(&dll_ref);
    *ppEnum = (IEnumVARIANT*)copy;
    return S_OK;
}

static const IEnumVARIANTVtbl windows_enum_vtbl =
{
    windows_enum_QueryInterface,
    windows_enum_AddRef,
    windows_enum_Release,
    windows_enum_Next,
    windows_enum_Skip,
    windows_enum_Reset,
    windows_enum_Clone
};

/* ---- shared helpers ---- */
static IDispatch *window_document(WindowImpl *This)
{
    return This && This->pWorkbook ? ((WorkbookImpl*)This->pWorkbook)->pDoc : NULL;
}

static HRESULT uno_call(IDispatch *object, LPOLESTR name, VARIANT *result)
{
    VariantInit(result);
    if (!object) return E_POINTER;
    return AutoWrap(DISPATCH_METHOD, result, object, name, 0);
}

static HRESULT uno_object(IDispatch *object, LPOLESTR name, IDispatch **value)
{
    VARIANT result;
    HRESULT hr = uno_call(object, name, &result);
    *value = NULL;
    if (SUCCEEDED(hr) && V_VT(&result) == VT_DISPATCH && V_DISPATCH(&result)) {
        *value = V_DISPATCH(&result);
        return S_OK;
    }
    VariantClear(&result);
    return FAILED(hr) ? hr : E_FAIL;
}

static HRESULT window_controller(WindowImpl *This, IDispatch **controller)
{
    return uno_object(window_document(This), L"getCurrentController", controller);
}

static HRESULT window_frame(WindowImpl *This, IDispatch **frame)
{
    IDispatch *controller = NULL;
    HRESULT hr = window_controller(This, &controller);
    *frame = NULL;
    if (SUCCEEDED(hr)) hr = uno_object(controller, L"getFrame", frame);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT window_container(WindowImpl *This, IDispatch **window)
{
    IDispatch *frame = NULL;
    HRESULT hr = window_frame(This, &frame);
    *window = NULL;
    if (SUCCEEDED(hr)) hr = uno_object(frame, L"getContainerWindow", window);
    if (frame) IDispatch_Release(frame);
    return hr;
}

static HRESULT window_get_long(IDispatch *object, LPOLESTR name, BOOL method, LONG *value)
{
    VARIANT result;
    HRESULT hr;
    VariantInit(&result);
    if (!object) return E_POINTER;
    hr = AutoWrap(method ? DISPATCH_METHOD : DISPATCH_PROPERTYGET, &result, object, name, 0);
    if (SUCCEEDED(hr)) hr = VariantChangeTypeEx(&result, &result, 0, 0, VT_I4);
    if (SUCCEEDED(hr)) *value = V_I4(&result);
    VariantClear(&result);
    return hr;
}

static HRESULT window_put_value(IDispatch *object, LPOLESTR name, VARIANT value)
{
    VARIANT result;
    HRESULT hr;
    VariantInit(&result);
    if (!object) return E_POINTER;
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &result, object, name, 1, value);
    VariantClear(&result);
    return hr;
}

/* View settings are properties of the spreadsheet controller. */
static HRESULT window_get_flag(I_Window *iface, LPOLESTR name, VARIANT_BOOL *value)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    VARIANT result;
    HRESULT hr;
    if (!value) return E_POINTER;
    VariantInit(&result);
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_PROPERTYGET, &result, controller, name, 0);
    if (SUCCEEDED(hr)) hr = VariantChangeTypeEx(&result, &result, 0, 0, VT_BOOL);
    if (SUCCEEDED(hr)) *value = V_BOOL(&result) ? VARIANT_TRUE : VARIANT_FALSE;
    VariantClear(&result);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT window_put_flag(I_Window *iface, LPOLESTR name, VARIANT_BOOL value)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    VARIANT flag;
    HRESULT hr;
    V_VT(&flag) = VT_BOOL;
    V_BOOL(&flag) = value ? VARIANT_TRUE : VARIANT_FALSE;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_put_value(controller, name, flag);
    if (controller) IDispatch_Release(controller);
    return hr;
}

/* Excel window geometry is in points; UNO windows use pixels. */
static double window_pixels_per_point(void)
{
    HDC dc = GetDC(NULL);
    int dpi = dc ? GetDeviceCaps(dc, LOGPIXELSY) : 96;
    if (dc) ReleaseDC(NULL, dc);
    return (dpi > 0 ? dpi : 96) / 72.0;
}

/* The office window's HWND (SystemDependent.SYSTEM_WIN32), or NULL. Moving
   through VCL's setPosSize is ignored right after a state change, so
   geometry uses Win32 directly when the backend window is a Win32 one. */
static HWND window_hwnd(WindowImpl *This)
{
    IDispatch *window = NULL;
    SAFEARRAY *process;
    VARIANT id, type, result;
    HWND hwnd = NULL;

    if (FAILED(window_container(This, &window))) return NULL;
    process = SafeArrayCreateVector(VT_UI1, 0, 16);
    VariantInit(&result);
    if (process) {
        V_VT(&id) = VT_ARRAY | VT_UI1;
        V_ARRAY(&id) = process;
        V_VT(&type) = VT_I2;
        V_I2(&type) = 1;
        if (SUCCEEDED(AutoWrap(DISPATCH_METHOD, &result, window, L"getWindowHandle", 2, type, id)) &&
                SUCCEEDED(VariantChangeTypeEx(&result, &result, 0, 0, VT_I4)) && V_I4(&result))
            hwnd = (HWND)(ULONG_PTR)(ULONG)V_I4(&result);
        SafeArrayDestroy(process);
    }
    VariantClear(&result);
    IDispatch_Release(window);
    return hwnd && IsWindow(hwnd) ? hwnd : NULL;
}

static HRESULT window_get_bounds(WindowImpl *This, LONG bounds[4])
{
    static const LPOLESTR names[4] = {L"X", L"Y", L"Width", L"Height"};
    IDispatch *window = NULL, *possize = NULL;
    HWND hwnd = window_hwnd(This);
    RECT rect;
    HRESULT hr;
    int i;

    if (hwnd && GetWindowRect(hwnd, &rect)) {
        bounds[0] = rect.left;
        bounds[1] = rect.top;
        bounds[2] = rect.right - rect.left;
        bounds[3] = rect.bottom - rect.top;
        return S_OK;
    }
    hr = window_container(This, &window);
    if (SUCCEEDED(hr)) hr = uno_object(window, L"getPosSize", &possize);
    for (i = 0; SUCCEEDED(hr) && i < 4; i++) hr = window_get_long(possize, names[i], FALSE, &bounds[i]);
    if (possize) IDispatch_Release(possize);
    if (window) IDispatch_Release(window);
    return hr;
}

static HRESULT window_get_point(I_Window *iface, int index, double *value)
{
    LONG bounds[4];
    HRESULT hr;
    if (!value) return E_POINTER;
    hr = window_get_bounds(WINDOW_THIS(iface), bounds);
    if (SUCCEEDED(hr)) *value = bounds[index] / window_pixels_per_point();
    return hr;
}

static HRESULT window_put_point(I_Window *iface, int index, double value)
{
    /* VCL ignores a move with only PosSize.X or Y set, so always pass the
       complete rectangle with PosSize.POSSIZE. */
    static const SHORT flags[4] = {15, 15, 15, 15};
    IDispatch *window = NULL;
    VARIANT a[5], result;
    LONG bounds[4];
    HRESULT hr;
    int i;
    HWND hwnd;
    if (index >= 2 && value < 0) return E_INVALIDARG;
    hr = window_get_bounds(WINDOW_THIS(iface), bounds);
    bounds[index] = (LONG)(value * window_pixels_per_point() + 0.5);
    hwnd = window_hwnd(WINDOW_THIS(iface));
    if (SUCCEEDED(hr) && hwnd)
        return SetWindowPos(hwnd, NULL, bounds[0], bounds[1], bounds[2], bounds[3],
                SWP_NOZORDER | SWP_NOACTIVATE) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    if (SUCCEEDED(hr)) hr = window_container(WINDOW_THIS(iface), &window);
    VariantInit(&result);
    V_VT(&a[0]) = VT_I2;
    V_I2(&a[0]) = flags[index];
    for (i = 0; i < 4; i++) {
        V_VT(&a[4 - i]) = VT_I4;
        V_I4(&a[4 - i]) = bounds[i];
    }
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_METHOD, &result, window, L"setPosSize", 5, a[0], a[1], a[2], a[3], a[4]);
    VariantClear(&result);
    if (window) IDispatch_Release(window);
    return hr;
}

static HRESULT window_wrap_range(WindowImpl *This, LONG sr, LONG sc, LONG er, LONG ec, IDispatch **result)
{
    IDispatch *sheet = NULL;
    I_Range *range = NULL;
    VARIANT a[4], rect;
    HRESULT hr;
    int i;
    *result = NULL;
    hr = I_Workbook_get_ActiveSheet(This->pWorkbook, &sheet);
    if (FAILED(hr)) return hr;
    for (i = 0; i < 4; i++) V_VT(&a[i]) = VT_I4;
    V_I4(&a[0]) = er; V_I4(&a[1]) = ec; V_I4(&a[2]) = sr; V_I4(&a[3]) = sc;
    VariantInit(&rect);
    hr = AutoWrap(DISPATCH_METHOD, &rect, ((WorksheetImpl*)sheet)->pOOSheet, L"getCellRangeByPosition", 4, a[0], a[1], a[2], a[3]);
    if (SUCCEEDED(hr)) hr = _I_RangeConstructor((void**)&range);
    if (SUCCEEDED(hr)) {
        I_Range_AddRef(range);
        hr = MSO_TO_OO_I_Range_Initialize3(range, V_DISPATCH(&rect), sheet, NULL);
        if (SUCCEEDED(hr)) *result = (IDispatch*)range;
        else I_Range_Release(range);
    }
    VariantClear(&rect);
    IDispatch_Release(sheet);
    return hr;
}

static HRESULT window_address(IDispatch *source, LONG bounds[4])
{
    IDispatch *address = NULL;
    static const LPOLESTR names[4] = {L"StartRow", L"StartColumn", L"EndRow", L"EndColumn"};
    HRESULT hr = uno_object(source, L"getRangeAddress", &address);
    int i;
    /* getVisibleRange() already returns a CellRangeAddress struct. */
    if (FAILED(hr)) {
        address = source;
        IDispatch_AddRef(address);
        hr = S_OK;
    }
    for (i = 0; SUCCEEDED(hr) && i < 4; i++) hr = window_get_long(address, names[i], FALSE, &bounds[i]);
    if (address) IDispatch_Release(address);
    return hr;
}

/* The selection is a cell range; multi-area selections use their first area. */
static HRESULT window_selection_bounds(WindowImpl *This, LONG bounds[4])
{
    IDispatch *controller = NULL, *selection = NULL, *area = NULL;
    VARIANT index, result;
    HRESULT hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = uno_object(controller, L"getSelection", &selection);
    if (SUCCEEDED(hr)) {
        hr = window_address(selection, bounds);
        if (FAILED(hr)) {
            VariantInit(&result);
            V_VT(&index) = VT_I4;
            V_I4(&index) = 0;
            hr = AutoWrap(DISPATCH_METHOD, &result, selection, L"getByIndex", 1, index);
            if (SUCCEEDED(hr) && V_VT(&result) == VT_DISPATCH) {
                area = V_DISPATCH(&result);
                hr = window_address(area, bounds);
                IDispatch_Release(area);
            } else {
                VariantClear(&result);
                if (SUCCEEDED(hr)) hr = E_FAIL;
            }
        }
    }
    if (selection) IDispatch_Release(selection);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static WorkbooksImpl *window_books(WindowImpl *This)
{
    return This && This->pWorkbook ? (WorkbooksImpl*)((WorkbookImpl*)This->pWorkbook)->pworkbooks : NULL;
}

static int window_book_index(WindowImpl *This)
{
    WorkbooksImpl *books = window_books(This);
    int i;
    for (i = 0; books && i < books->count_workbooks; i++)
        if (books->pworkbook[i] == This->pWorkbook) return i;
    return -1;
}

static HRESULT window_scroll(I_Window *iface, VARIANT Down, VARIANT Up, VARIANT ToRight, VARIANT ToLeft, BOOL pages)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL, *visible = NULL;
    VARIANT steps[4], value, result;
    LONG count[4] = {0, 0, 0, 0}, row = 0, column = 0, height = 1, width = 1, bounds[4];
    VARIANT *source[4] = {&Down, &Up, &ToRight, &ToLeft};
    HRESULT hr;
    int i;

    for (i = 0; i < 4; i++) {
        VariantInit(&steps[i]);
        if (!Is_Variant_Null(*source[i])) {
            hr = VariantChangeTypeEx(&steps[i], source[i], 0, 0, VT_I4);
            if (FAILED(hr)) return hr;
            count[i] = V_I4(&steps[i]);
        }
    }
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"getFirstVisibleRow", TRUE, &row);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"getFirstVisibleColumn", TRUE, &column);
    /* LargeScroll moves by the visible page size. */
    if (SUCCEEDED(hr) && pages && SUCCEEDED(uno_object(controller, L"getVisibleRange", &visible))) {
        if (SUCCEEDED(window_address(visible, bounds))) {
            height = bounds[2] - bounds[0] + 1;
            width = bounds[3] - bounds[1] + 1;
        }
        IDispatch_Release(visible);
    }
    row += (count[0] - count[1]) * height;
    column += (count[2] - count[3]) * width;
    if (row < 0) row = 0;
    if (column < 0) column = 0;
    VariantInit(&result);
    V_VT(&value) = VT_I4;
    V_I4(&value) = row;
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_METHOD, &result, controller, L"setFirstVisibleRow", 1, value);
    VariantClear(&result);
    V_I4(&value) = column;
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_METHOD, &result, controller, L"setFirstVisibleColumn", 1, value);
    VariantClear(&result);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT window_freeze(WindowImpl *This, LONG columns, LONG rows)
{
    IDispatch *controller = NULL;
    VARIANT a[2], result;
    HRESULT hr = window_controller(This, &controller);
    VariantInit(&result);
    V_VT(&a[0]) = VT_I4;
    V_I4(&a[0]) = rows;
    V_VT(&a[1]) = VT_I4;
    V_I4(&a[1]) = columns;
    /* freezeAtPosition(0, 0) leaves an empty frozen split; removing the
       split at pixel 0,0 unfreezes the view. */
    if (SUCCEEDED(hr))
        hr = AutoWrap(DISPATCH_METHOD, &result, controller, columns || rows ? L"freezeAtPosition" : L"splitAtPosition", 2, a[0], a[1]);
    VariantClear(&result);
    if (controller) IDispatch_Release(controller);
    return hr;
}

/* ---- I_Windows ---- */
static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Application(
        I_Windows* iface,
        IDispatch **value)
{
    WindowsImpl *This = WINDOWS_THIS(iface);
    if (!This || !value) return E_POINTER;
    *value = (IDispatch*)This->pApplication;
    if (!*value) return E_POINTER;
    IDispatch_AddRef(*value);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Creator(
        I_Windows* iface,
        XlCreator *RHS)
{
    if (!RHS) return E_POINTER;
    *RHS = xlCreatorCode;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Parent(
        I_Windows* iface,
        IDispatch **RHS)
{
    WindowsImpl *This = WINDOWS_THIS(iface);
    if (!This || !RHS) return E_POINTER;
    *RHS = This->pWorkbook ? (IDispatch*)This->pWorkbook : (IDispatch*)This->pApplication;
    if (!*RHS) return E_POINTER;
    IDispatch_AddRef(*RHS);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Count(
        I_Windows* iface,
        LONG *retval)
{
    WindowsImpl *This = WINDOWS_THIS(iface);
    WorkbooksImpl *books;
    if (!This || !retval) return E_POINTER;
    if (This->pWorkbook) {
        *retval = 1;
        return S_OK;
    }
    books = (WorkbooksImpl*)((_ApplicationImpl*)This->pApplication)->pdWorkbooks;
    *retval = books ? books->count_workbooks : 0;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Item(
        I_Windows* iface,
        VARIANT Index,
        IDispatch **RHS)
{
    WindowsImpl *This = WINDOWS_THIS(iface);
    WorkbooksImpl *books;
    VARIANT index;
    HRESULT hr;
    int i, count;

    if (!This || !RHS) return E_POINTER;
    *RHS = NULL;
    books = (WorkbooksImpl*)((_ApplicationImpl*)This->pApplication)->pdWorkbooks;
    count = This->pWorkbook ? 1 : (books ? books->count_workbooks : 0);
    VariantInit(&index);
    hr = VariantCopyInd(&index, &Index);
    if (FAILED(hr)) return hr;
    if (V_VT(&index) == VT_BSTR) {
        /* Windows are also addressed by caption (the workbook name). */
        for (i = 0; i < count; i++) {
            I_Workbook *book = This->pWorkbook ? This->pWorkbook : books->pworkbook[i];
            BSTR name = NULL;
            if (SUCCEEDED(I_Workbook_get_Name(book, &name)) && !lstrcmpiW(name, V_BSTR(&index))) {
                SysFreeString(name);
                VariantClear(&index);
                return MSO_TO_OO_Window_Create(book, RHS);
            }
            SysFreeString(name);
        }
        VariantClear(&index);
        return DISP_E_BADINDEX;
    }
    hr = VariantChangeTypeEx(&index, &index, 0, 0, VT_I4);
    if (FAILED(hr)) return hr;
    if (V_I4(&index) < 1 || V_I4(&index) > count) return DISP_E_BADINDEX;
    return MSO_TO_OO_Window_Create(This->pWorkbook ? This->pWorkbook : books->pworkbook[V_I4(&index) - 1], RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get__Default(
        I_Windows* iface,
        VARIANT Index,
        IDispatch **RHS)
{
    return MSO_TO_OO_I_Windows_get_Item(iface, Index, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get__NewEnum(
        I_Windows* iface,
        IUnknown **RHS)
{
    WindowsImpl *This = WINDOWS_THIS(iface);
    WindowsEnum *cursor;
    LONG count = 0, i;
    VARIANT index;
    HRESULT hr;

    if (!This || !RHS) return E_POINTER;
    *RHS = NULL;
    hr = MSO_TO_OO_I_Windows_get_Count(iface, &count);
    if (FAILED(hr)) return hr;
    cursor = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(*cursor) + count * sizeof(IDispatch*));
    if (!cursor) return E_OUTOFMEMORY;
    cursor->lpVtbl = &windows_enum_vtbl;
    cursor->ref = 1;
    cursor->count = count;
    /* A snapshot keeps the enumeration stable while workbooks change. */
    for (i = 0; SUCCEEDED(hr) && i < count; i++) {
        V_VT(&index) = VT_I4;
        V_I4(&index) = i + 1;
        hr = MSO_TO_OO_I_Windows_get_Item(iface, index, &cursor->items[i]);
    }
    if (FAILED(hr)) {
        IEnumVARIANT_Release((IEnumVARIANT*)cursor);
        return hr;
    }
    InterlockedIncrement(&dll_ref);
    *RHS = (IUnknown*)cursor;
    return S_OK;
}

/* ---- I_Window ---- */
static HRESULT WINAPI MSO_TO_OO_I_Window_get_Application(
        I_Window* iface,
        IDispatch **value)
{
    WindowImpl *This = WINDOW_THIS(iface);
    if (!This || !value) return E_POINTER;
    return I_Workbook_get_Application(This->pWorkbook, value);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Creator(
        I_Window* iface,
        XlCreator *RHS)
{
    if (!RHS) return E_POINTER;
    *RHS = xlCreatorCode;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Parent(
        I_Window* iface,
        IDispatch **RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    if (!This || !RHS) return E_POINTER;
    *RHS = (IDispatch*)This->pWorkbook;
    I_Workbook_AddRef(This->pWorkbook);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_Activate(
        I_Window* iface,
        VARIANT *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    HRESULT hr;
    if (!This) return E_POINTER;
    hr = MSO_TO_OO_ActivateWorkbook(This->pWorkbook);
    if (RHS) {
        VariantInit(RHS);
        if (SUCCEEDED(hr)) {
            V_VT(RHS) = VT_BOOL;
            V_BOOL(RHS) = VARIANT_TRUE;
        }
    }
    return hr;
}

static HRESULT window_activate_relative(I_Window *iface, int step, VARIANT *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    WorkbooksImpl *books = window_books(This);
    int index = window_book_index(This);
    HRESULT hr;
    if (!books || index < 0) return E_FAIL;
    index = (index + step + books->count_workbooks) % books->count_workbooks;
    hr = MSO_TO_OO_ActivateWorkbook(books->pworkbook[index]);
    if (RHS) {
        VariantInit(RHS);
        if (SUCCEEDED(hr)) {
            V_VT(RHS) = VT_BOOL;
            V_BOOL(RHS) = VARIANT_TRUE;
        }
    }
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_ActivateNext(
        I_Window* iface,
        VARIANT *RHS)
{
    return window_activate_relative(iface, 1, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_ActivatePrevious(
        I_Window* iface,
        VARIANT *RHS)
{
    return window_activate_relative(iface, -1, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActiveCell(
        I_Window* iface,
        IDispatch **RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    LONG bounds[4];
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    *RHS = NULL;
    hr = window_selection_bounds(This, bounds);
    if (SUCCEEDED(hr)) hr = window_wrap_range(This, bounds[0], bounds[1], bounds[0], bounds[1], RHS);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActiveSheet(
        I_Window* iface,
        IDispatch **RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    if (!This || !RHS) return E_POINTER;
    return I_Workbook_get_ActiveSheet(This->pWorkbook, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Caption(
        I_Window* iface,
        VARIANT *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *frame = NULL;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    VariantInit(RHS);
    hr = window_frame(This, &frame);
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_PROPERTYGET, RHS, frame, L"Title", 0);
    if (frame) IDispatch_Release(frame);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Caption(
        I_Window* iface,
        VARIANT RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *frame = NULL;
    VARIANT text;
    HRESULT hr;
    if (!This) return E_POINTER;
    VariantInit(&text);
    hr = VariantChangeTypeEx(&text, &RHS, 0, 0, VT_BSTR);
    if (SUCCEEDED(hr)) hr = window_frame(This, &frame);
    if (SUCCEEDED(hr)) hr = window_put_value(frame, L"Title", text);
    if (frame) IDispatch_Release(frame);
    VariantClear(&text);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_Close(
        I_Window* iface,
        VARIANT SaveChanges,
        VARIANT Filename,
        VARIANT RouteWorkbook,
        VARIANT_BOOL *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    HRESULT hr;
    if (!This) return E_POINTER;
    /* A workbook has one window, so closing it closes the workbook. */
    hr = I_Workbook_Close(This->pWorkbook, SaveChanges, Filename, RouteWorkbook, 0);
    if (RHS) *RHS = SUCCEEDED(hr) ? VARIANT_TRUE : VARIANT_FALSE;
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayFormulas(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"ShowFormulas", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayFormulas(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"ShowFormulas", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayGridlines(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"ShowGrid", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayGridlines(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"ShowGrid", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayHeadings(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"HasColumnRowHeaders", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayHeadings(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"HasColumnRowHeaders", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayHorizontalScrollBar(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"HasHorizontalScrollBar", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayHorizontalScrollBar(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"HasHorizontalScrollBar", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayOutline(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"IsOutlineSymbolsSet", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayOutline(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"IsOutlineSymbolsSet", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayVerticalScrollBar(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"HasVerticalScrollBar", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayVerticalScrollBar(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"HasVerticalScrollBar", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayWorkbookTabs(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"HasSheetTabs", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayWorkbookTabs(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"HasSheetTabs", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayZeros(I_Window* iface, VARIANT_BOOL *RHS)
{
    return window_get_flag(iface, L"ShowZeroValues", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayZeros(I_Window* iface, VARIANT_BOOL RHS)
{
    return window_put_flag(iface, L"ShowZeroValues", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_FreezePanes(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    LONG frozen = 0;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"hasFrozenPanes", TRUE, &frozen);
    if (SUCCEEDED(hr)) *RHS = frozen ? VARIANT_TRUE : VARIANT_FALSE;
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_FreezePanes(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    LONG bounds[4], row = 0, column = 0;
    HRESULT hr;
    if (!This) return E_POINTER;
    if (!RHS) return window_freeze(This, 0, 0);
    /* Excel freezes the rows above and the columns left of the active cell. */
    hr = window_selection_bounds(This, bounds);
    if (SUCCEEDED(hr)) hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"getFirstVisibleRow", TRUE, &row);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"getFirstVisibleColumn", TRUE, &column);
    if (controller) IDispatch_Release(controller);
    if (SUCCEEDED(hr)) {
        LONG columns = bounds[1] > column ? bounds[1] - column : 0;
        LONG rows = bounds[0] > row ? bounds[0] - row : 0;
        if (!rows && !columns) return E_FAIL;
        hr = window_freeze(This, columns, rows);
    }
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_GridlineColor(
        I_Window* iface,
        LONG *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    LONG color = 0;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"GridColor", FALSE, &color);
    /* UNO uses 0xRRGGBB, Excel RGB() values are 0xBBGGRR. */
    if (SUCCEEDED(hr)) *RHS = ((color & 0xff) << 16) | (color & 0xff00) | ((color >> 16) & 0xff);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_GridlineColor(
        I_Window* iface,
        LONG RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    VARIANT color;
    HRESULT hr;
    if (!This) return E_POINTER;
    V_VT(&color) = VT_I4;
    V_I4(&color) = ((RHS & 0xff) << 16) | (RHS & 0xff00) | ((RHS >> 16) & 0xff);
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_put_value(controller, L"GridColor", color);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Height(I_Window* iface, double *RHS)
{
    return window_get_point(iface, 3, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Height(I_Window* iface, double RHS)
{
    return window_put_point(iface, 3, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Index(
        I_Window* iface,
        LONG *RHS)
{
    int index = window_book_index(WINDOW_THIS(iface));
    if (!RHS) return E_POINTER;
    if (index < 0) return E_FAIL;
    *RHS = index + 1;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_LargeScroll(
        I_Window* iface,
        VARIANT Down,
        VARIANT Up,
        VARIANT toRight,
        VARIANT toLeft,
        VARIANT *RHS)
{
    if (RHS) VariantInit(RHS);
    return window_scroll(iface, Down, Up, toRight, toLeft, TRUE);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Left(I_Window* iface, double *RHS)
{
    return window_get_point(iface, 0, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Left(I_Window* iface, double RHS)
{
    return window_put_point(iface, 0, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PrintPreview(
        I_Window* iface,
        VARIANT EnableChanges,
        VARIANT *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    if (RHS) VariantInit(RHS);
    if (!This) return E_POINTER;
    return MSO_TO_OO_PrintPreview(This->pWorkbook);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_RangeSelection(
        I_Window* iface,
        IDispatch **RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    LONG bounds[4];
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    *RHS = NULL;
    hr = window_selection_bounds(This, bounds);
    if (SUCCEEDED(hr)) hr = window_wrap_range(This, bounds[0], bounds[1], bounds[2], bounds[3], RHS);
    return hr;
}

static HRESULT window_get_first(I_Window *iface, LPOLESTR name, LONG *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    LONG value = 0;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, name, TRUE, &value);
    if (SUCCEEDED(hr)) *RHS = value + 1;
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT window_put_first(I_Window *iface, LPOLESTR name, LONG value)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    VARIANT number, result;
    HRESULT hr;
    if (!This) return E_POINTER;
    if (value < 1) return E_INVALIDARG;
    VariantInit(&result);
    V_VT(&number) = VT_I4;
    V_I4(&number) = value - 1;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_METHOD, &result, controller, name, 1, number);
    VariantClear(&result);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ScrollColumn(I_Window* iface, LONG *RHS)
{
    return window_get_first(iface, L"getFirstVisibleColumn", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_ScrollColumn(I_Window* iface, LONG RHS)
{
    return window_put_first(iface, L"setFirstVisibleColumn", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ScrollRow(I_Window* iface, LONG *RHS)
{
    return window_get_first(iface, L"getFirstVisibleRow", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_ScrollRow(I_Window* iface, LONG RHS)
{
    return window_put_first(iface, L"setFirstVisibleRow", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Selection(
        I_Window* iface,
        IDispatch **RHS)
{
    return MSO_TO_OO_I_Window_get_RangeSelection(iface, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_SmallScroll(
        I_Window* iface,
        VARIANT Down,
        VARIANT Up,
        VARIANT toRight,
        VARIANT toLeft,
        VARIANT *RHS)
{
    if (RHS) VariantInit(RHS);
    return window_scroll(iface, Down, Up, toRight, toLeft, FALSE);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Split(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    LONG split = 0, frozen = 0;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"getIsWindowSplit", TRUE, &split);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"hasFrozenPanes", TRUE, &frozen);
    if (SUCCEEDED(hr)) *RHS = split || frozen ? VARIANT_TRUE : VARIANT_FALSE;
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Split(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    VARIANT a[2], result;
    HRESULT hr;
    if (!This) return E_POINTER;
    /* Only removing a split is defined without a position. */
    if (RHS) return E_NOTIMPL;
    hr = window_freeze(This, 0, 0);
    VariantInit(&result);
    V_VT(&a[0]) = VT_I4;
    V_I4(&a[0]) = 0;
    V_VT(&a[1]) = VT_I4;
    V_I4(&a[1]) = 0;
    if (SUCCEEDED(hr)) hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_METHOD, &result, controller, L"splitAtPosition", 2, a[0], a[1]);
    VariantClear(&result);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT window_get_split(I_Window *iface, LPOLESTR name, LONG *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, name, TRUE, RHS);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT window_put_split(I_Window *iface, BOOL row, LONG value)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    LONG frozen = 0, rows = 0, columns = 0;
    HRESULT hr;
    if (!This) return E_POINTER;
    if (value < 0) return E_INVALIDARG;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"hasFrozenPanes", TRUE, &frozen);
    if (SUCCEEDED(hr) && frozen) {
        hr = window_get_long(controller, L"getSplitRow", TRUE, &rows);
        if (SUCCEEDED(hr)) hr = window_get_long(controller, L"getSplitColumn", TRUE, &columns);
    }
    if (controller) IDispatch_Release(controller);
    if (FAILED(hr)) return hr;
    /* UNO positions unfrozen splits in pixels; cell-based splits are frozen panes. */
    if (!frozen && value) return E_NOTIMPL;
    if (row) rows = value;
    else columns = value;
    return window_freeze(This, columns, rows);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitColumn(I_Window* iface, LONG *RHS)
{
    return window_get_split(iface, L"getSplitColumn", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitColumn(I_Window* iface, LONG RHS)
{
    return window_put_split(iface, FALSE, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitRow(I_Window* iface, LONG *RHS)
{
    return window_get_split(iface, L"getSplitRow", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitRow(I_Window* iface, LONG RHS)
{
    return window_put_split(iface, TRUE, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Top(I_Window* iface, double *RHS)
{
    return window_get_point(iface, 1, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Top(I_Window* iface, double RHS)
{
    return window_put_point(iface, 1, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Type(
        I_Window* iface,
        XlWindowType *RHS)
{
    if (!RHS) return E_POINTER;
    *RHS = xlWorkbook;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_UsableHeight(I_Window* iface, double *RHS)
{
    return window_get_point(iface, 3, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_UsableWidth(I_Window* iface, double *RHS)
{
    return window_get_point(iface, 2, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Visible(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *window = NULL;
    LONG visible = 0;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    hr = window_container(This, &window);
    if (SUCCEEDED(hr)) hr = window_get_long(window, L"isVisible", TRUE, &visible);
    if (SUCCEEDED(hr)) *RHS = visible ? VARIANT_TRUE : VARIANT_FALSE;
    if (window) IDispatch_Release(window);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Visible(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *window = NULL;
    VARIANT flag, result;
    HRESULT hr;
    if (!This) return E_POINTER;
    VariantInit(&result);
    V_VT(&flag) = VT_BOOL;
    V_BOOL(&flag) = RHS ? VARIANT_TRUE : VARIANT_FALSE;
    hr = window_container(This, &window);
    if (SUCCEEDED(hr)) hr = AutoWrap(DISPATCH_METHOD, &result, window, L"setVisible", 1, flag);
    VariantClear(&result);
    if (window) IDispatch_Release(window);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_VisibleRange(
        I_Window* iface,
        IDispatch **RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL, *visible = NULL;
    LONG bounds[4];
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    *RHS = NULL;
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = uno_object(controller, L"getVisibleRange", &visible);
    if (SUCCEEDED(hr)) hr = window_address(visible, bounds);
    if (SUCCEEDED(hr)) hr = window_wrap_range(This, bounds[0], bounds[1], bounds[2], bounds[3], RHS);
    if (visible) IDispatch_Release(visible);
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Width(I_Window* iface, double *RHS)
{
    return window_get_point(iface, 2, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Width(I_Window* iface, double RHS)
{
    return window_put_point(iface, 2, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_WindowNumber(
        I_Window* iface,
        LONG *RHS)
{
    if (!RHS) return E_POINTER;
    *RHS = 1;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_WindowState(
        I_Window* iface,
        XlWindowState *RHS)
{
    if (!iface || !RHS) return E_POINTER;
    return MSO_TO_OO_GetWorkbookWindowState(WINDOW_THIS(iface)->pWorkbook, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_WindowState(
        I_Window* iface,
        XlWindowState RHS)
{
    if (!iface) return E_POINTER;
    return MSO_TO_OO_SetWorkbookWindowState(WINDOW_THIS(iface)->pWorkbook, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Zoom(
        I_Window* iface,
        VARIANT *RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    LONG zoom = 100;
    HRESULT hr;
    if (!This || !RHS) return E_POINTER;
    VariantInit(RHS);
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_get_long(controller, L"ZoomValue", FALSE, &zoom);
    if (SUCCEEDED(hr)) {
        V_VT(RHS) = VT_R8;
        V_R8(RHS) = zoom;
    }
    if (controller) IDispatch_Release(controller);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Zoom(
        I_Window* iface,
        VARIANT RHS)
{
    WindowImpl *This = WINDOW_THIS(iface);
    IDispatch *controller = NULL;
    VARIANT value, type, zoom;
    HRESULT hr;
    if (!This) return E_POINTER;
    VariantInit(&value);
    hr = VariantCopyInd(&value, &RHS);
    if (FAILED(hr)) return hr;
    V_VT(&type) = VT_I2;
    if (V_VT(&value) == VT_BOOL) {
        /* Zoom = True fits the selection: DocumentZoomType.OPTIMAL. */
        if (!V_BOOL(&value)) return E_INVALIDARG;
        V_I2(&type) = 0;
    } else {
        hr = VariantChangeTypeEx(&value, &value, 0, 0, VT_I2);
        if (FAILED(hr)) return hr;
        if (V_I2(&value) < 10 || V_I2(&value) > 400) return E_INVALIDARG;
        V_I2(&type) = 3; /* DocumentZoomType.BY_VALUE */
    }
    hr = window_controller(This, &controller);
    if (SUCCEEDED(hr)) hr = window_put_value(controller, L"ZoomType", type);
    if (SUCCEEDED(hr) && V_I2(&type) == 3) {
        V_VT(&zoom) = VT_I2;
        V_I2(&zoom) = V_I2(&value);
        hr = window_put_value(controller, L"ZoomValue", zoom);
    }
    if (controller) IDispatch_Release(controller);
    return hr;
}


/*** IUnknown and IDispatch for I_Windows ***/
static HRESULT WINAPI MSO_TO_OO_I_Windows_QueryInterface(I_Windows* iface, REFIID riid, void **ppvObject)
{
    if (!iface || !ppvObject) return E_POINTER;
    *ppvObject = NULL;
    if (IsEqualGUID(riid, &IID_IDispatch) || IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Windows)) {
        *ppvObject = iface;
        I_Windows_AddRef(iface);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Windows_AddRef(I_Windows* iface)
{
    WindowsImpl *This = (WindowsImpl*)iface;
    ULONG ref = InterlockedIncrement(&This->ref);
    if (ref == 1) InterlockedIncrement(&dll_ref);
    return ref;
}

static ULONG WINAPI MSO_TO_OO_I_Windows_Release(I_Windows* iface)
{
    WindowsImpl *This = (WindowsImpl*)iface;
    ULONG ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pWorkbook) I_Workbook_Release(This->pWorkbook);
        if (This->pApplication) _Application_Release(This->pApplication);
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
        DELETE_OBJECT;
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_GetTypeInfoCount(I_Windows* iface, UINT *pctinfo)
{
    if (!pctinfo) return E_POINTER;
    *pctinfo = 1;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_GetTypeInfo(I_Windows* iface, UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    HRESULT hres;
    if (!ppTInfo) return E_POINTER;
    *ppTInfo = NULL;
    if (iTInfo) return DISP_E_BADINDEX;
    hres = get_typeinfo_windows(ppTInfo);
    /* The cache owns one reference; the caller must receive another. */
    if (SUCCEEDED(hres)) ITypeInfo_AddRef(*ppTInfo);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_GetIDsOfNames(I_Windows* iface, REFIID riid, LPOLESTR *rgszNames,
        UINT cNames, LCID lcid, DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres = get_typeinfo_windows(&typeinfo);
    if (FAILED(hres)) return hres;
    return ITypeInfo_GetIDsOfNames(typeinfo, rgszNames, cNames, rgDispId);
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_Invoke(I_Windows* iface, DISPID dispIdMember, REFIID riid, LCID lcid,
        WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
    ITypeInfo *typeinfo;
    HRESULT hres = get_typeinfo_windows(&typeinfo);
    if (FAILED(hres)) return hres;
    return ITypeInfo_Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}


/* Members without an OpenOffice/LibreOffice counterpart. */
static HRESULT WINAPI MSO_TO_OO_I_Windows_Arrange(
        I_Windows* iface,
        XlArrangeStyle ArrangeStyle,
        VARIANT ActiveWorkbook,
        VARIANT SyncHorizontal,
        VARIANT SyncVertical,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}


static const I_WindowsVtbl MSO_TO_OO_I_WindowsVtbl =
{
    MSO_TO_OO_I_Windows_QueryInterface,
    MSO_TO_OO_I_Windows_AddRef,
    MSO_TO_OO_I_Windows_Release,
    MSO_TO_OO_I_Windows_GetTypeInfoCount,
    MSO_TO_OO_I_Windows_GetTypeInfo,
    MSO_TO_OO_I_Windows_GetIDsOfNames,
    MSO_TO_OO_I_Windows_Invoke,
    MSO_TO_OO_I_Windows_get_Application,
    MSO_TO_OO_I_Windows_get_Creator,
    MSO_TO_OO_I_Windows_get_Parent,
    MSO_TO_OO_I_Windows_Arrange,
    MSO_TO_OO_I_Windows_get_Count,
    MSO_TO_OO_I_Windows_get_Item,
    MSO_TO_OO_I_Windows_get__NewEnum,
    MSO_TO_OO_I_Windows_get__Default
};


/*** IUnknown and IDispatch for I_Window ***/
static HRESULT WINAPI MSO_TO_OO_I_Window_QueryInterface(I_Window* iface, REFIID riid, void **ppvObject)
{
    if (!iface || !ppvObject) return E_POINTER;
    *ppvObject = NULL;
    if (IsEqualGUID(riid, &IID_IDispatch) || IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Window)) {
        *ppvObject = iface;
        I_Window_AddRef(iface);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Window_AddRef(I_Window* iface)
{
    WindowImpl *This = (WindowImpl*)iface;
    ULONG ref = InterlockedIncrement(&This->ref);
    if (ref == 1) InterlockedIncrement(&dll_ref);
    return ref;
}

static ULONG WINAPI MSO_TO_OO_I_Window_Release(I_Window* iface)
{
    WindowImpl *This = (WindowImpl*)iface;
    ULONG ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pWorkbook) I_Workbook_Release(This->pWorkbook);
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
        DELETE_OBJECT;
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_GetTypeInfoCount(I_Window* iface, UINT *pctinfo)
{
    if (!pctinfo) return E_POINTER;
    *pctinfo = 1;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_GetTypeInfo(I_Window* iface, UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    HRESULT hres;
    if (!ppTInfo) return E_POINTER;
    *ppTInfo = NULL;
    if (iTInfo) return DISP_E_BADINDEX;
    hres = get_typeinfo_window(ppTInfo);
    /* The cache owns one reference; the caller must receive another. */
    if (SUCCEEDED(hres)) ITypeInfo_AddRef(*ppTInfo);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_GetIDsOfNames(I_Window* iface, REFIID riid, LPOLESTR *rgszNames,
        UINT cNames, LCID lcid, DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres = get_typeinfo_window(&typeinfo);
    if (FAILED(hres)) return hres;
    return ITypeInfo_GetIDsOfNames(typeinfo, rgszNames, cNames, rgDispId);
}

static HRESULT WINAPI MSO_TO_OO_I_Window_Invoke(I_Window* iface, DISPID dispIdMember, REFIID riid, LCID lcid,
        WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
    ITypeInfo *typeinfo;
    HRESULT hres = get_typeinfo_window(&typeinfo);
    if (FAILED(hres)) return hres;
    return ITypeInfo_Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}


/* Members without an OpenOffice/LibreOffice counterpart. */
static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActiveChart(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActivePane(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get__DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put__DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_EnableResize(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_EnableResize(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_GridlineColorIndex(
        I_Window* iface,
        XlColorIndex *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_GridlineColorIndex(
        I_Window* iface,
        XlColorIndex RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_NewWindow(
        I_Window* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_OnWindow(
        I_Window* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_OnWindow(
        I_Window* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Panes(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PrintOut(
        I_Window* iface,
        VARIANT From,
        VARIANT To_,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        VARIANT PrToFileName,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_ScrollWorkbookTabs(
        I_Window* iface,
        VARIANT Sheets,
        VARIANT Position,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SelectedSheets(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitHorizontal(
        I_Window* iface,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitHorizontal(
        I_Window* iface,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitVertical(
        I_Window* iface,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitVertical(
        I_Window* iface,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_TabRatio(
        I_Window* iface,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_TabRatio(
        I_Window* iface,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_View(
        I_Window* iface,
        XlWindowView *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_View(
        I_Window* iface,
        XlWindowView RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PointsToScreenPixelsX(
        I_Window* iface,
        LONG Points,
        LONG *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PointsToScreenPixelsY(
        I_Window* iface,
        LONG Points,
        LONG *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_RangeFromPoint(
        I_Window* iface,
        LONG x,
        LONG y,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_ScrollIntoView(
        I_Window* iface,
        LONG left,
        LONG top,
        LONG width,
        LONG height,
        VARIANT Start)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}


static const I_WindowVtbl MSO_TO_OO_I_WindowVtbl =
{
    MSO_TO_OO_I_Window_QueryInterface,
    MSO_TO_OO_I_Window_AddRef,
    MSO_TO_OO_I_Window_Release,
    MSO_TO_OO_I_Window_GetTypeInfoCount,
    MSO_TO_OO_I_Window_GetTypeInfo,
    MSO_TO_OO_I_Window_GetIDsOfNames,
    MSO_TO_OO_I_Window_Invoke,
    MSO_TO_OO_I_Window_get_Application,
    MSO_TO_OO_I_Window_get_Creator,
    MSO_TO_OO_I_Window_get_Parent,
    MSO_TO_OO_I_Window_Activate,
    MSO_TO_OO_I_Window_ActivateNext,
    MSO_TO_OO_I_Window_ActivatePrevious,
    MSO_TO_OO_I_Window_get_ActiveCell,
    MSO_TO_OO_I_Window_get_ActiveChart,
    MSO_TO_OO_I_Window_get_ActivePane,
    MSO_TO_OO_I_Window_get_ActiveSheet,
    MSO_TO_OO_I_Window_get_Caption,
    MSO_TO_OO_I_Window_put_Caption,
    MSO_TO_OO_I_Window_Close,
    MSO_TO_OO_I_Window_get_DisplayFormulas,
    MSO_TO_OO_I_Window_put_DisplayFormulas,
    MSO_TO_OO_I_Window_get_DisplayGridlines,
    MSO_TO_OO_I_Window_put_DisplayGridlines,
    MSO_TO_OO_I_Window_get_DisplayHeadings,
    MSO_TO_OO_I_Window_put_DisplayHeadings,
    MSO_TO_OO_I_Window_get_DisplayHorizontalScrollBar,
    MSO_TO_OO_I_Window_put_DisplayHorizontalScrollBar,
    MSO_TO_OO_I_Window_get_DisplayOutline,
    MSO_TO_OO_I_Window_put_DisplayOutline,
    MSO_TO_OO_I_Window_get__DisplayRightToLeft,
    MSO_TO_OO_I_Window_put__DisplayRightToLeft,
    MSO_TO_OO_I_Window_get_DisplayVerticalScrollBar,
    MSO_TO_OO_I_Window_put_DisplayVerticalScrollBar,
    MSO_TO_OO_I_Window_get_DisplayWorkbookTabs,
    MSO_TO_OO_I_Window_put_DisplayWorkbookTabs,
    MSO_TO_OO_I_Window_get_DisplayZeros,
    MSO_TO_OO_I_Window_put_DisplayZeros,
    MSO_TO_OO_I_Window_get_EnableResize,
    MSO_TO_OO_I_Window_put_EnableResize,
    MSO_TO_OO_I_Window_get_FreezePanes,
    MSO_TO_OO_I_Window_put_FreezePanes,
    MSO_TO_OO_I_Window_get_GridlineColor,
    MSO_TO_OO_I_Window_put_GridlineColor,
    MSO_TO_OO_I_Window_get_GridlineColorIndex,
    MSO_TO_OO_I_Window_put_GridlineColorIndex,
    MSO_TO_OO_I_Window_get_Height,
    MSO_TO_OO_I_Window_put_Height,
    MSO_TO_OO_I_Window_get_Index,
    MSO_TO_OO_I_Window_LargeScroll,
    MSO_TO_OO_I_Window_get_Left,
    MSO_TO_OO_I_Window_put_Left,
    MSO_TO_OO_I_Window_NewWindow,
    MSO_TO_OO_I_Window_get_OnWindow,
    MSO_TO_OO_I_Window_put_OnWindow,
    MSO_TO_OO_I_Window_get_Panes,
    MSO_TO_OO_I_Window_PrintOut,
    MSO_TO_OO_I_Window_PrintPreview,
    MSO_TO_OO_I_Window_get_RangeSelection,
    MSO_TO_OO_I_Window_get_ScrollColumn,
    MSO_TO_OO_I_Window_put_ScrollColumn,
    MSO_TO_OO_I_Window_get_ScrollRow,
    MSO_TO_OO_I_Window_put_ScrollRow,
    MSO_TO_OO_I_Window_ScrollWorkbookTabs,
    MSO_TO_OO_I_Window_get_SelectedSheets,
    MSO_TO_OO_I_Window_get_Selection,
    MSO_TO_OO_I_Window_SmallScroll,
    MSO_TO_OO_I_Window_get_Split,
    MSO_TO_OO_I_Window_put_Split,
    MSO_TO_OO_I_Window_get_SplitColumn,
    MSO_TO_OO_I_Window_put_SplitColumn,
    MSO_TO_OO_I_Window_get_SplitHorizontal,
    MSO_TO_OO_I_Window_put_SplitHorizontal,
    MSO_TO_OO_I_Window_get_SplitRow,
    MSO_TO_OO_I_Window_put_SplitRow,
    MSO_TO_OO_I_Window_get_SplitVertical,
    MSO_TO_OO_I_Window_put_SplitVertical,
    MSO_TO_OO_I_Window_get_TabRatio,
    MSO_TO_OO_I_Window_put_TabRatio,
    MSO_TO_OO_I_Window_get_Top,
    MSO_TO_OO_I_Window_put_Top,
    MSO_TO_OO_I_Window_get_Type,
    MSO_TO_OO_I_Window_get_UsableHeight,
    MSO_TO_OO_I_Window_get_UsableWidth,
    MSO_TO_OO_I_Window_get_Visible,
    MSO_TO_OO_I_Window_put_Visible,
    MSO_TO_OO_I_Window_get_VisibleRange,
    MSO_TO_OO_I_Window_get_Width,
    MSO_TO_OO_I_Window_put_Width,
    MSO_TO_OO_I_Window_get_WindowNumber,
    MSO_TO_OO_I_Window_get_WindowState,
    MSO_TO_OO_I_Window_put_WindowState,
    MSO_TO_OO_I_Window_get_Zoom,
    MSO_TO_OO_I_Window_put_Zoom,
    MSO_TO_OO_I_Window_get_View,
    MSO_TO_OO_I_Window_put_View,
    MSO_TO_OO_I_Window_get_DisplayRightToLeft,
    MSO_TO_OO_I_Window_put_DisplayRightToLeft,
    MSO_TO_OO_I_Window_PointsToScreenPixelsX,
    MSO_TO_OO_I_Window_PointsToScreenPixelsY,
    MSO_TO_OO_I_Window_RangeFromPoint,
    MSO_TO_OO_I_Window_ScrollIntoView
};


HRESULT MSO_TO_OO_Windows_Create(
        _Application *app,
        I_Workbook *workbook,
        IDispatch **result)
{
    WindowsImpl *windows;
    if (!result) return E_POINTER;
    *result = NULL;
    if (!app) return E_POINTER;
    windows = HeapAlloc(GetProcessHeap(), 0, sizeof(*windows));
    if (!windows) return E_OUTOFMEMORY;
    windows->_windowsVtbl = &MSO_TO_OO_I_WindowsVtbl;
    windows->ref = 0;
    windows->pApplication = app;
    _Application_AddRef(app);
    windows->pWorkbook = workbook;
    if (workbook) I_Workbook_AddRef(workbook);
    CREATE_OBJECT;
    I_Windows_AddRef((I_Windows*)windows);
    *result = (IDispatch*)windows;
    return S_OK;
}

HRESULT MSO_TO_OO_Window_Create(
        I_Workbook *workbook,
        IDispatch **result)
{
    WindowImpl *window;
    if (!result) return E_POINTER;
    *result = NULL;
    if (!workbook) return E_POINTER;
    window = HeapAlloc(GetProcessHeap(), 0, sizeof(*window));
    if (!window) return E_OUTOFMEMORY;
    window->_windowVtbl = &MSO_TO_OO_I_WindowVtbl;
    window->ref = 0;
    window->pWorkbook = workbook;
    I_Workbook_AddRef(workbook);
    CREATE_OBJECT;
    I_Window_AddRef((I_Window*)window);
    *result = (IDispatch*)window;
    return S_OK;
}
