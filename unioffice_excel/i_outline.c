/*
 * IOutline interface functions
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

ITypeInfo *ti_outline = NULL;

HRESULT get_typeinfo_outline(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if(ti_outline) {
        *typeinfo = ti_outline;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Outline, &ti_outline);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_outline;
    return hres;
}

#define OUTLINE_THIS(iface) DEFINE_THIS(OutlineImpl, outline, iface)

    /*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Outline_AddRef(
        I_Outline* iface)
{
    OutlineImpl *This = OUTLINE_THIS(iface);
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (!This) {
        ERR("object is NULL \n");         
        return E_POINTER;
    }

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_QueryInterface(
        I_Outline* iface,
        REFIID riid,
        void **ppvObject)
{
    OutlineImpl *This = OUTLINE_THIS(iface);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    if (!ppvObject) {
        ERR("Object2 is NULL\n");
        return E_POINTER;
    }

    *ppvObject = NULL;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Outline)) {
        *ppvObject = OUTLINE_OUTLINE(This);
        I_Outline_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Outline_Release(
        I_Outline* iface)
{
    OutlineImpl *This = OUTLINE_THIS(iface);
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pWorksheet) {
            I_Worksheet_Release(This->pWorksheet);
            This->pWorksheet = NULL;
        }
        if (This->pOOSheet) {
            IDispatch_Release(This->pOOSheet);
            This->pOOSheet = NULL;
        }        
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
        DELETE_OBJECT;
    }
    return ref;
}

/*** IOutline methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Outline_get_Application(
        I_Outline* iface,
        IDispatch **RHS)
{
    OutlineImpl *This = OUTLINE_THIS(iface);
    HRESULT hres;
    TRACE_IN;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    if (!RHS) {
        ERR("Object RHS is NULL \n");
        return E_POINTER;
    }

    hres = I_Worksheet_get_Application(This->pWorksheet, RHS);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_Creator(
        I_Outline* iface,
        XlCreator *RHS)
{
    TRACE_IN;
    *RHS = xlCreatorCode;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_Parent(
        I_Outline* iface,
        IDispatch **RHS)
{
    OutlineImpl *This = OUTLINE_THIS(iface);
    TRACE_IN;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    if (!RHS) {
        ERR("Object RHS is NULL \n");
        return E_POINTER;
    }

    *RHS = (IDispatch*)(This->pWorksheet);
    I_Worksheet_AddRef(This->pWorksheet);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_AutomaticStyles(
        I_Outline* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_IN;
    /*Always return VARIANT_FALSE*/
    *RHS = VARIANT_FALSE;
    TRACE_OUT;
    return S_OK;
}

#define RANGE_THIS(iface) DEFINE_THIS(RangeImpl, range, iface)
static HRESULT WINAPI MSO_TO_OO_I_Outline_put_AutomaticStyles(
        I_Outline* iface,
        VARIANT_BOOL RHS)
{
    OutlineImpl *This = OUTLINE_THIS(iface);
    HRESULT hres;
    VARIANT param1, cols, vret;
    I_Range *tmp_range;
    TRACE_IN;

    VariantClear(&param1);
    VariantClear(&cols);
    VariantClear(&vret);

    if (RHS == VARIANT_TRUE) {
        V_VT(&cols) = VT_BSTR;
        {
            WCHAR all_columns[24];
            wsprintfW(all_columns, L"1:%d", (int)MSO_TO_OO_SheetColumnCount(This->pOOSheet));
            V_BSTR(&cols) = SysAllocString(all_columns);
        }
        I_Worksheet_get_Columns(This->pWorksheet, cols, (IDispatch**)&tmp_range);
        
        RangeImpl *rangeimpl = RANGE_THIS(tmp_range);
        
        hres = AutoWrap(DISPATCH_METHOD, &param1, rangeimpl->pOORange, L"getRangeAddress", 0);
        if (FAILED(hres)) {
            ERR("getRangeAddress\n");
        }

        hres = AutoWrap(DISPATCH_METHOD, &vret, This->pOOSheet, L"autoOutline", 1, param1);
        if (FAILED(hres)) 
            ERR("autoOutline\n");
            
        if (tmp_range) {
            I_Range_Release(tmp_range);
            tmp_range = NULL;
        }
        
        VariantClear(&param1);
        VariantClear(&cols);
    } else {
        hres = AutoWrap(DISPATCH_METHOD, &vret, This->pOOSheet, L"clearOutline", 0);
        if (FAILED(hres)) 
            ERR("autoOutline\n");
    }

    VariantClear(&vret);

    TRACE_OUT;
    return hres;
}
#undef RANGE_THIS

static HRESULT WINAPI MSO_TO_OO_I_Outline_ShowLevels(
        I_Outline *iface, VARIANT RowLevels, VARIANT ColumnLevels, VARIANT *RHS)
{
    OutlineImpl *This = OUTLINE_THIS(iface);
    VARIANT inputs[2], level, orientation, result;
    HRESULT hr = S_OK;
    int index;

    if (!RHS) return E_POINTER;
    VariantInit(RHS);
    MSO_TO_OO_CorrectArg(RowLevels, &inputs[0]);
    MSO_TO_OO_CorrectArg(ColumnLevels, &inputs[1]);
    if (Is_Variant_Null(inputs[0]) && Is_Variant_Null(inputs[1])) return E_INVALIDARG;
    VariantInit(&level);
    VariantInit(&orientation);
    VariantInit(&result);
    V_VT(&orientation) = VT_I4;
    for (index = 0; index < 2 && SUCCEEDED(hr); ++index) {
        if (Is_Variant_Null(inputs[index])) continue;
        hr = VariantChangeTypeEx(&level, &inputs[index], 0, 0, VT_I2);
        if (FAILED(hr)) break;
        if (V_I2(&level) < 0) {
            hr = E_INVALIDARG;
            break;
        }
        if (!V_I2(&level)) continue;
        /* Excel labels outline levels from 1; Calc uses a zero-based depth. */
        --V_I2(&level);
        V_I4(&orientation) = index == 0 ? toROWS : toCOLUMNS;
        hr = AutoWrap(DISPATCH_METHOD, &result, This->pOOSheet, L"showLevel", 2, orientation, level);
        VariantClear(&result);
    }
    VariantClear(&level);
    if (SUCCEEDED(hr)) {
        V_VT(RHS) = VT_BOOL;
        V_BOOL(RHS) = VARIANT_TRUE;
    }
    return hr;
}

static HRESULT outline_setting_name(OutlineImpl *This, const WCHAR *kind, VARIANT *name)
{
    VARIANT sheet;
    WCHAR text[300];
    HRESULT hr;

    /* UNO outlines have no summary position, so the Excel setting is kept
       per sheet in the document's user-defined properties. */
    VariantInit(&sheet);
    hr = AutoWrap(DISPATCH_METHOD, &sheet, This->pOOSheet, L"getName", 0);
    if (SUCCEEDED(hr) && (V_VT(&sheet) != VT_BSTR || SysStringLen(V_BSTR(&sheet)) > 250)) hr = E_FAIL;
    if (SUCCEEDED(hr)) {
        wsprintfW(text, L"UniOffice.Outline.%s.%s", kind, V_BSTR(&sheet));
        V_VT(name) = VT_BSTR;
        V_BSTR(name) = SysAllocString(text);
        if (!V_BSTR(name)) hr = E_OUTOFMEMORY;
    }
    VariantClear(&sheet);
    return hr;
}

static HRESULT outline_settings(OutlineImpl *This, IDispatch **properties)
{
    WorksheetImpl *sheet = (WorksheetImpl*)This->pWorksheet;
    WorkbookImpl *book = sheet ? (WorkbookImpl*)sheet->pwb : NULL;
    VARIANT info, result;
    HRESULT hr;

    *properties = NULL;
    if (!book || !book->pDoc) return E_POINTER;
    VariantInit(&info);
    VariantInit(&result);
    hr = AutoWrap(DISPATCH_METHOD, &info, book->pDoc, L"getDocumentProperties", 0);
    if (SUCCEEDED(hr) && V_VT(&info) == VT_DISPATCH)
        hr = AutoWrap(DISPATCH_METHOD, &result, V_DISPATCH(&info), L"getUserDefinedProperties", 0);
    else if (SUCCEEDED(hr)) hr = E_FAIL;
    if (SUCCEEDED(hr) && V_VT(&result) == VT_DISPATCH) {
        *properties = V_DISPATCH(&result);
        VariantInit(&result);
    } else if (SUCCEEDED(hr)) hr = E_FAIL;
    VariantClear(&result);
    VariantClear(&info);
    return hr;
}

static HRESULT outline_has_setting(IDispatch *properties, VARIANT *name, BOOL *exists)
{
    VARIANT info, result;
    HRESULT hr;
    VariantInit(&info);
    VariantInit(&result);
    *exists = FALSE;
    hr = AutoWrap(DISPATCH_METHOD, &info, properties, L"getPropertySetInfo", 0);
    if (SUCCEEDED(hr) && V_VT(&info) == VT_DISPATCH)
        hr = AutoWrap(DISPATCH_METHOD, &result, V_DISPATCH(&info), L"hasPropertyByName", 1, *name);
    else if (SUCCEEDED(hr)) hr = E_FAIL;
    if (SUCCEEDED(hr)) hr = VariantChangeTypeEx(&result, &result, 0, 0, VT_BOOL);
    if (SUCCEEDED(hr)) *exists = V_BOOL(&result) != VARIANT_FALSE;
    VariantClear(&result);
    VariantClear(&info);
    return hr;
}

static HRESULT outline_get_setting(OutlineImpl *This, const WCHAR *kind, LONG fallback, LONG *value)
{
    IDispatch *properties = NULL;
    VARIANT name, result;
    BOOL exists = FALSE;
    HRESULT hr;

    *value = fallback;
    VariantInit(&name);
    VariantInit(&result);
    hr = outline_settings(This, &properties);
    if (SUCCEEDED(hr)) hr = outline_setting_name(This, kind, &name);
    /* A missing property means the Excel default. */
    if (SUCCEEDED(hr)) hr = outline_has_setting(properties, &name, &exists);
    if (SUCCEEDED(hr) && exists) {
        hr = AutoWrap(DISPATCH_METHOD, &result, properties, L"getPropertyValue", 1, name);
        if (SUCCEEDED(hr)) hr = VariantChangeTypeEx(&result, &result, 0, 0, VT_I4);
        if (SUCCEEDED(hr)) *value = V_I4(&result);
    }
    VariantClear(&result);
    VariantClear(&name);
    if (properties) IDispatch_Release(properties);
    return hr;
}

static HRESULT outline_put_setting(OutlineImpl *This, const WCHAR *kind, LONG value)
{
    IDispatch *properties = NULL;
    VARIANT name, number, attributes, result;
    BOOL exists = FALSE;
    HRESULT hr;

    VariantInit(&name);
    VariantInit(&result);
    /* Document user properties accept short but not long integers. */
    V_VT(&number) = VT_I2;
    V_I2(&number) = (SHORT)value;
    hr = outline_settings(This, &properties);
    if (SUCCEEDED(hr)) hr = outline_setting_name(This, kind, &name);
    if (SUCCEEDED(hr)) hr = outline_has_setting(properties, &name, &exists);
    if (SUCCEEDED(hr) && exists) {
        hr = AutoWrap(DISPATCH_METHOD, &result, properties, L"setPropertyValue", 2, number, name);
    } else if (SUCCEEDED(hr)) {
        V_VT(&attributes) = VT_I2;
        V_I2(&attributes) = 128; /* PropertyAttribute.REMOVABLE */
        hr = AutoWrap(DISPATCH_METHOD, &result, properties, L"addProperty", 3, number, attributes, name);
    }
    VariantClear(&result);
    VariantClear(&name);
    if (properties) IDispatch_Release(properties);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_SummaryColumn(
        I_Outline* iface,
        XlSummaryColumn *RHS)
{
    LONG value;
    HRESULT hr;
    if (!iface || !RHS) return E_POINTER;
    hr = outline_get_setting(OUTLINE_THIS(iface), L"SummaryColumn", xlSummaryOnRight, &value);
    if (SUCCEEDED(hr)) *RHS = value;
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_put_SummaryColumn(
        I_Outline* iface,
        XlSummaryColumn RHS)
{
    if (!iface) return E_POINTER;
    if (RHS != xlSummaryOnLeft && RHS != xlSummaryOnRight) return E_INVALIDARG;
    return outline_put_setting(OUTLINE_THIS(iface), L"SummaryColumn", RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_SummaryRow(
        I_Outline* iface,
        XlSummaryRow *RHS)
{
    LONG value;
    HRESULT hr;
    if (!iface || !RHS) return E_POINTER;
    hr = outline_get_setting(OUTLINE_THIS(iface), L"SummaryRow", xlSummaryBelow, &value);
    if (SUCCEEDED(hr)) *RHS = value;
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_put_SummaryRow(
        I_Outline* iface,
        XlSummaryRow RHS)
{
    if (!iface) return E_POINTER;
    if (RHS != xlSummaryAbove && RHS != xlSummaryBelow) return E_INVALIDARG;
    return outline_put_setting(OUTLINE_THIS(iface), L"SummaryRow", RHS);
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Outline_GetTypeInfoCount(
        I_Outline* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_GetTypeInfo(
        I_Outline* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres;
    if (!ppTInfo) return E_POINTER;
    *ppTInfo = NULL;
    if (iTInfo) return DISP_E_BADINDEX;
    hres = get_typeinfo_outline(ppTInfo);
    /* The cache owns one reference; the caller must receive another. */
    if (SUCCEEDED(hres)) ITypeInfo_AddRef(*ppTInfo);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_GetIDsOfNames(
        I_Outline* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_outline(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WERR(L"name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_Invoke(
        I_Outline* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_outline(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
    if (FAILED(hres)) {
        ERR("wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
    }

    TRACE_OUT;
    return hres;
}

#undef OUTLINE_THIS

const I_OutlineVtbl MSO_TO_OO_I_Outline_Vtbl =
{
    MSO_TO_OO_I_Outline_QueryInterface,
    MSO_TO_OO_I_Outline_AddRef,
    MSO_TO_OO_I_Outline_Release,
    MSO_TO_OO_I_Outline_GetTypeInfoCount,
    MSO_TO_OO_I_Outline_GetTypeInfo,
    MSO_TO_OO_I_Outline_GetIDsOfNames,
    MSO_TO_OO_I_Outline_Invoke,
    MSO_TO_OO_I_Outline_get_Application,
    MSO_TO_OO_I_Outline_get_Creator,
    MSO_TO_OO_I_Outline_get_Parent,
    MSO_TO_OO_I_Outline_get_AutomaticStyles,
    MSO_TO_OO_I_Outline_put_AutomaticStyles,
    MSO_TO_OO_I_Outline_ShowLevels,
    MSO_TO_OO_I_Outline_get_SummaryColumn,
    MSO_TO_OO_I_Outline_put_SummaryColumn,
    MSO_TO_OO_I_Outline_get_SummaryRow,
    MSO_TO_OO_I_Outline_put_SummaryRow
};

extern HRESULT _I_OutlineConstructor(LPVOID *ppObj)
{
    OutlineImpl *outline;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    outline = HeapAlloc(GetProcessHeap(), 0, sizeof(*outline));
    if (!outline)
    {
        return E_OUTOFMEMORY;
    }

    outline->poutlineVtbl = &MSO_TO_OO_I_Outline_Vtbl;
    outline->ref = 0;
    outline->pWorksheet = NULL;
    outline->pOOSheet = NULL;

    *ppObj = OUTLINE_OUTLINE(outline);
        
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}
