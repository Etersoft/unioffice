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

#define toCOLUMNS 0
#define toROWS 1

static WCHAR const str_showlevels[] = {
    'S','h','o','w','L','e','v','e','l','s',0};
static WCHAR const str_summarycolumn[] = {
    'S','u','m','m','a','r','y','C','o','l','u','m','n',0};
static WCHAR const str_summaryrow[] = {
    'S','u','m','m','a','r','y','R','o','w',0};

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

    TRACE("\n");

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
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_Creator(
        I_Outline* iface,
        XlCreator *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_Parent(
        I_Outline* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_AutomaticStyles(
        I_Outline* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_put_AutomaticStyles(
        I_Outline* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
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

    TRACE("\n");

    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&vret);

    if ((V_VT(&RowLevels)!=VT_NULL) && (V_VT(&ColumnLevels)!=VT_EMPTY)) {
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_get_SummaryColumn(
        I_Outline* iface,
        XlSummaryColumn *RHS)
{
    TRACE("\n");
    *RHS = xlSummaryOnLeft;
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
    TRACE("\n");
    *RHS = xlSummaryAbove;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_put_SummaryRow(
        I_Outline* iface,
        XlSummaryRow RHS)
{
    TRACE("\n");
    return S_OK;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Outline_GetTypeInfoCount(
        I_Outline* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_GetTypeInfo(
        I_Outline* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Outline_GetIDsOfNames(
        I_Outline* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    if (!lstrcmpiW(*rgszNames, str_showlevels)) {
        *rgDispId = dispid_outline_showlevels;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_summarycolumn)) {
        *rgDispId = dispid_outline_summarycolumn;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_summaryrow)) {
        *rgDispId = dispid_outline_summaryrow;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L" NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
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
    VARIANT param1, param2, vNull;
    HRESULT hres;
    long lret = 0;

    TRACE("\n");

    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;

    switch (dispIdMember) {
        case dispid_outline_showlevels:
            switch (pDispParams->cArgs) {
                case 1:
                    MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &param1);
                    hres = MSO_TO_OO_I_Outline_ShowLevels(iface, param1, vNull, pVarResult);
                    if (FAILED(hres)) {
                        pExcepInfo->bstrDescription=SysAllocString(str_error);
                        TRACE("(case 2) ERROR ShowLevels hres = %08x\n",hres);
                    }
                    return hres;
                case 2:
                    MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &param1);
                    MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &param2);
                    hres = MSO_TO_OO_I_Outline_ShowLevels(iface, param1, param2, pVarResult);
                    if (FAILED(hres)) {
                        pExcepInfo->bstrDescription=SysAllocString(str_error);
                        TRACE("(case 2) ERROR ShowLevels hres = %08x\n",hres);
                    }
                    return hres;
                default:
                    TRACE("Error parameters \n");
                    return E_FAIL;
            }
            break;
        case dispid_outline_summarycolumn:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                if (pDispParams->cArgs!=1) return E_FAIL;
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &param1);
                hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
                if (FAILED(hres)) {
                    TRACE("(case 4) ERROR VariantChangeTypeEx   %08x\n",hres);
                    return hres;
                }
                lret = V_I4(&param1);
                return MSO_TO_OO_I_Outline_put_SummaryColumn(iface, (XlSummaryColumn)lret);
            } else {
                hres = MSO_TO_OO_I_Outline_get_SummaryColumn(iface,(XlSummaryColumn*)&lret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_I4;
                    V_I4(pVarResult) = lret;
                }
                return S_OK;
            }
            break;
        case dispid_outline_summaryrow:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                if (pDispParams->cArgs!=1) return E_FAIL;
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &param1);
                hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
                if (FAILED(hres)) {
                    TRACE("(case 4) ERROR VariantChangeTypeEx   %08x\n",hres);
                    return hres;
                }
                lret = V_I4(&param1);
                return MSO_TO_OO_I_Outline_put_SummaryRow(iface, (XlSummaryRow)lret);
            } else {
                hres = MSO_TO_OO_I_Outline_get_SummaryRow(iface,(XlSummaryRow*)&lret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_I4;
                    V_I4(pVarResult) = lret;
                }
                return S_OK;
            }
            break;
    }
    TRACE("dispid ( %i ) Not realized\n");
    return E_NOTIMPL;
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

    return S_OK;
}
