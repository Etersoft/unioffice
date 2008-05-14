/*
 * IInterior interface functions
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

static WCHAR const str_colorindex[] = {
    'C','o','l','o','r','I','n','d','e','x',0};
static WCHAR const str_color[] = {
    'C','o','l','o','r',0};
static WCHAR const str_application[] = {
    'A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR const str_parent[] = {
    'P','a','r','e','n','t',0};
static WCHAR const str_creator[] = {
    'C','r','e','a','t','o','r',0};

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Interior_AddRef(
        I_Interior* iface)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_QueryInterface(
        I_Interior* iface,
        REFIID riid,
        void **ppvObject)
{
    InteriorImpl *This = (InteriorImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Interior)) {
        *ppvObject = &This->_interiorVtbl;
        MSO_TO_OO_I_Interior_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Interior_Release(
        I_Interior* iface)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->prange!=NULL) {
            IDispatch_Release(This->prange);
            This->prange = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Interior methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Interior_get_Color(
        I_Interior* iface,
        long *plcolor)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    HRESULT hres;
    VARIANT vret;
    VariantInit(&vret);

    TRACE("\n");

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    RangeImpl *cur_range = (RangeImpl*)(I_Range*)(This->prange);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vret, cur_range->pOORange, L"CellBackColor", 0);

    if (FAILED(hres)) {
        TRACE("ERROR when CellBackColor");
    }

    hres = VariantChangeTypeEx(&vret, &vret, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
    return E_FAIL;
    }
    *plcolor = V_I4(&vret);
    TRACE("lcolor=%i\n",*plcolor);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_Color(
        I_Interior* iface,
        long lcolor)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    HRESULT hres;
    VARIANT vret,param1;

    TRACE(" lcolor = %i\n",lcolor);

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    RangeImpl *cur_range = (RangeImpl*)((I_Range*)(This->prange));

    VariantInit(&param1);
    V_VT(&param1) = VT_BOOL;
    V_BOOL(&param1) = VARIANT_TRUE;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"IsCellBackgroundTransparent", 1, param1);

    if (FAILED(hres)) TRACE("ERROR when IsCellBackgroundTransparent");

    VariantClear(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = lcolor;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"cellBackColor", 1, param1);

    if (FAILED(hres)) TRACE("ERROR when cellBackColor");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_ColorIndex(
        I_Interior* iface,
        long *plcolorindex)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    long tmpcolor;
    int i;
    HRESULT hres;
    TRACE(" \n");

    if (This==NULL) return E_POINTER;


    hres = MSO_TO_OO_I_Interior_get_Color(iface,&tmpcolor);
    if (FAILED(hres)) {
        return hres;
    }
    for (i=0;i<56;i++)
        if (color[i]==tmpcolor) {
            *plcolorindex = i+1;
            return S_OK;
        }

    TRACE("ERROR Color don`t have colorindex \n");
    *plcolorindex = 1;/*белый цвет*/
    /*Отправляем что все хорошо, на всякий случай*/
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_ColorIndex(
        I_Interior* iface,
        long lcolorindex)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    long tmpcolor;
    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;

    if ((lcolorindex<1)||(lcolorindex>56)) {
        TRACE("ERROR Incorrect colorindex %i\n", lcolorindex);
        return S_OK;
    } else 
        return MSO_TO_OO_I_Interior_put_Color(iface,color[lcolorindex-1]);
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_Application(
        I_Interior* iface,
        IDispatch **value)
{
    InteriorImpl *This = (InteriorImpl*)iface;

    TRACE("\n");

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    return I_Range_get_Application((I_Range*)(This->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_Parent(
        I_Interior* iface,
        IDispatch **value)
{
    InteriorImpl *This = (InteriorImpl*)iface;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    *value = This->prange;

    if (value==NULL)
        return E_POINTER;

    I_Range_AddRef((I_Range*)*value);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_Creator(
        I_Interior* iface,
        VARIANT *result)
{
    TRACE(" \n");
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    return S_OK;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Interior_GetTypeInfoCount(
        I_Interior* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_GetTypeInfo(
        I_Interior* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_GetIDsOfNames(
        I_Interior* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    if (!lstrcmpiW(*rgszNames, str_colorindex)) {
        *rgDispId = 1;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_color)) {
        *rgDispId = 2;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = 3;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_parent)) {
        *rgDispId = 4;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_creator)) {
        *rgDispId = 5;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L" %s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_Invoke(
        I_Interior* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    HRESULT hres;
    IDispatch *dret;
    VARIANT vresult,vtmp;
    long ltmp;
    VariantInit(&vtmp);
    VariantInit(&vresult);

    if (This==NULL) return E_POINTER;

    switch(dispIdMember)
    {
    case 1:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (case 1) ERROR VariantChangeTypeEx   %08x\n",hres);
                return E_FAIL;
            }
            ltmp = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Interior_put_ColorIndex(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Interior_get_ColorIndex(iface,&ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = ltmp;
            }
            return S_OK;
        }
    case 2:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (case 1) ERROR VariantChangeTypeEx   %08x\n",hres);
                return E_FAIL;
            }
            ltmp = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Interior_put_Color(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Interior_get_Color(iface,&ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = ltmp;
            }
            return S_OK;
        }
    case 3:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Interior_get_Application(iface,&dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case 4:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Interior_get_Parent(iface,&dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case 5:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Interior_get_Creator(iface, &vresult);
            if (pVarResult!=NULL){ 
                *pVarResult = vresult;
            }
            return hres;
        }
    }

    TRACE(" dispIdMember = %i NOT REALIZE\n",dispIdMember);
    return E_NOTIMPL;
}

const I_InteriorVtbl MSO_TO_OO_I_Interior_Vtbl =
{
    MSO_TO_OO_I_Interior_QueryInterface,
    MSO_TO_OO_I_Interior_AddRef,
    MSO_TO_OO_I_Interior_Release ,
    MSO_TO_OO_I_Interior_GetTypeInfoCount,
    MSO_TO_OO_I_Interior_GetTypeInfo,
    MSO_TO_OO_I_Interior_GetIDsOfNames,
    MSO_TO_OO_I_Interior_Invoke,
    MSO_TO_OO_I_Interior_get_ColorIndex,
    MSO_TO_OO_I_Interior_put_ColorIndex,
    MSO_TO_OO_I_Interior_get_Color,
    MSO_TO_OO_I_Interior_put_Color,
    MSO_TO_OO_I_Interior_get_Application,
    MSO_TO_OO_I_Interior_get_Parent,
    MSO_TO_OO_I_Interior_get_Creator
};

InteriorImpl MSO_TO_OO_Interior =
{
    &MSO_TO_OO_I_Interior_Vtbl,
    0,
    NULL
};

extern HRESULT _I_InteriorConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    InteriorImpl *interior;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);
    
    interior = HeapAlloc(GetProcessHeap(), 0, sizeof(*interior));
    if (!interior)
    {
        return E_OUTOFMEMORY;
    }

    interior->_interiorVtbl = &MSO_TO_OO_I_Interior_Vtbl;
    interior->ref = 0;
    interior->prange = NULL;

    *ppObj = &interior->_interiorVtbl;

    return S_OK;
}
