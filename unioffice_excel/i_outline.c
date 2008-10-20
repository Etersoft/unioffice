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

    /*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Outline_AddRef(
        I_Outline* iface)
{
    OutlineImpl *This = (OutlineImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

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
    OutlineImpl *This = (OutlineImpl*)iface;

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Outline)) {
        *ppvObject = &This->_outlineVtbl;
        MSO_TO_OO_I_Outline_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Outline_Release(
        I_Outline* iface)
{
    OutlineImpl *This = (OutlineImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pwsh!=NULL) {
            IDispatch_Release(This->pwsh);
            This->pwsh = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** IOutline methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Outline_get_Application(
        I_Outline* iface,
        IDispatch **RHS)
{
    OutlineImpl *This = (OutlineImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (RHS==NULL)
        return E_POINTER;

    I_Worksheet_get_Application((I_Worksheet*)(This->pwsh),RHS);

    TRACE_OUT;
    return S_OK;
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
    OutlineImpl *This = (OutlineImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (RHS==NULL)
        return E_POINTER;

    *RHS = This->pwsh;
    I_Worksheet_AddRef((I_Worksheet*)(This->pwsh));

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

static HRESULT WINAPI MSO_TO_OO_I_Outline_put_AutomaticStyles(
        I_Outline* iface,
        VARIANT_BOOL RHS)
{
    OutlineImpl *This = (OutlineImpl*)iface;
    WorksheetImpl *wsh = (WorksheetImpl*)This->pwsh;
    HRESULT hres;
    VARIANT param1, cols, vret;
    IDispatch *tmp_range;
    TRACE_IN;

    VariantClear(&param1);
    VariantClear(&cols);
    VariantClear(&vret);

    if (RHS == VARIANT_TRUE) {
        V_VT(&cols) = VT_BSTR;
        switch (OOVersion) {
        case VER_2:
            V_BSTR(&cols) = SysAllocString(L"1:256");
            break;
        case VER_3:
            V_BSTR(&cols) = SysAllocString(L"1:1024");
            break;
        }
        I_Worksheet_get_Columns((I_Worksheet*)(This->pwsh),cols,&tmp_range);
        RangeImpl *rangeimp = (RangeImpl*)tmp_range;
        V_VT(&param1) = VT_DISPATCH;
        V_DISPATCH(&param1) = rangeimp->pOORange;
        IDispatch_AddRef(rangeimp->pOORange);

        hres = AutoWrap(DISPATCH_METHOD, &vret, wsh->pOOSheet, L"autoOutline", 1, param1);
        if (FAILED(hres)) TRACE("ERROR when autoOutline\n");
        VariantClear(&param1);
        IDispatch_Release(tmp_range);
        tmp_range = NULL;
        VariantClear(&cols);
    } else {
        hres = AutoWrap(DISPATCH_METHOD, &vret, wsh->pOOSheet, L"clearOutline", 0);
        if (FAILED(hres)) TRACE("ERROR when autoOutline\n");
    }

    VariantClear(&vret);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_ShowLevels(
        I_Outline* iface,
        VARIANT RowLevels,
        VARIANT ColumnLevels,
        VARIANT *RHS)
{
    OutlineImpl *This = (OutlineImpl*)iface;
    WorksheetImpl *wsh = (WorksheetImpl*)This->pwsh;
    HRESULT hres;
    VARIANT param1, param2, vret;
    TRACE_IN;

    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&vret);

    MSO_TO_OO_CorrectArg(RowLevels, &RowLevels);
    MSO_TO_OO_CorrectArg(ColumnLevels, &ColumnLevels);

    if (!Is_Variant_Null(RowLevels)) {
        hres = VariantChangeTypeEx(&param1, &RowLevels, 0, 0, VT_I4);
        if (FAILED(hres)) {
            TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
            return hres;
        }
        V_VT(&param2) = VT_I4;
        V_I4(&param2) = toROWS;
    } else {
        hres = VariantChangeTypeEx(&param1, &ColumnLevels, 0, 0, VT_I4);
        if (FAILED(hres)) {
            TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
            return hres;
        }
        V_VT(&param2) = VT_I4;
        V_I4(&param2) = toROWS;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vret, wsh->pOOSheet, L"showLevel", 2, param2, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when showLevel\n");
        return hres;
    }

    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&vret);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_SummaryColumn(
        I_Outline* iface,
        XlSummaryColumn *RHS)
{
    TRACE_IN;
    *RHS = xlSummaryOnLeft;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_put_SummaryColumn(
        I_Outline* iface,
        XlSummaryColumn RHS)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_SummaryRow(
        I_Outline* iface,
        XlSummaryRow *RHS)
{
    TRACE_IN;
    *RHS = xlSummaryAbove;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_put_SummaryRow(
        I_Outline* iface,
        XlSummaryRow RHS)
{
    TRACE_NOTIMPL;
    return S_OK;
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
    HRESULT hres = get_typeinfo_outline(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

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
        WTRACE(L"ERROR name = %s \n", *rgszNames);
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
        TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
    }

    TRACE_OUT;
    return hres;
}


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

    outline->_outlineVtbl = &MSO_TO_OO_I_Outline_Vtbl;
    outline->ref = 0;
    outline->pwsh = NULL;

    *ppObj = &outline->_outlineVtbl;
    TRACE_OUT;
    return S_OK;
}
