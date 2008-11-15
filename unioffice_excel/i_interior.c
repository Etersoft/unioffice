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

ITypeInfo *ti_interrior = NULL;

HRESULT get_typeinfo_interrior(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_interrior) {
        *typeinfo = ti_interrior;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Interior, &ti_interrior);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_interrior;

    return hres;
}

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
        DELETE_OBJECT;
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
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_Color(
        I_Interior* iface,
        long lcolor)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    HRESULT hres;
    VARIANT vret,param1;
    TRACE_IN;
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

    TRACE_OUT;
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
    TRACE_IN;

    if (This==NULL) return E_POINTER;


    hres = MSO_TO_OO_I_Interior_get_Color(iface,&tmpcolor);
    if (FAILED(hres)) {
        return hres;
    }
    for (i=0;i<56;i++)
        if (color[i]==tmpcolor) {
            *plcolorindex = i+1;
            TRACE_OUT;
            return S_OK;
        }

    TRACE("ERROR Color don`t have colorindex \n");
    *plcolorindex = 1;/*белый цвет*/
    /*Отправляем что все хорошо, на всякий случай*/
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_ColorIndex(
        I_Interior* iface,
        long lcolorindex)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    long tmpcolor;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;
    TRACE_OUT;
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
    TRACE_IN;

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;
    TRACE_OUT;
    return I_Range_get_Application((I_Range*)(This->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_Parent(
        I_Interior* iface,
        IDispatch **value)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    *value = This->prange;

    if (value==NULL)
        return E_POINTER;

    I_Range_AddRef((I_Range*)*value);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_Creator(
        I_Interior* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_InvertIfNegative(
        I_Interior* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_InvertIfNegative(
        I_Interior* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_Pattern(
        I_Interior* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_Pattern(
        I_Interior* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_PatternColor(
        I_Interior* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_PatternColor(
        I_Interior* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_get_PatternColorIndex(
        I_Interior* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_put_PatternColorIndex(
        I_Interior* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Interior_GetTypeInfoCount(
        I_Interior* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_GetTypeInfo(
        I_Interior* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_interrior(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Interior_GetIDsOfNames(
        I_Interior* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_interrior(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
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
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_interrior(&typeinfo);
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

const I_InteriorVtbl MSO_TO_OO_I_Interior_Vtbl =
{
    MSO_TO_OO_I_Interior_QueryInterface,
    MSO_TO_OO_I_Interior_AddRef,
    MSO_TO_OO_I_Interior_Release ,
    MSO_TO_OO_I_Interior_GetTypeInfoCount,
    MSO_TO_OO_I_Interior_GetTypeInfo,
    MSO_TO_OO_I_Interior_GetIDsOfNames,
    MSO_TO_OO_I_Interior_Invoke,
    MSO_TO_OO_I_Interior_get_Application,
    MSO_TO_OO_I_Interior_get_Creator,
    MSO_TO_OO_I_Interior_get_Parent,
    MSO_TO_OO_I_Interior_get_Color,
    MSO_TO_OO_I_Interior_put_Color,
    MSO_TO_OO_I_Interior_get_ColorIndex,
    MSO_TO_OO_I_Interior_put_ColorIndex,
    MSO_TO_OO_I_Interior_get_InvertIfNegative,
    MSO_TO_OO_I_Interior_put_InvertIfNegative,
    MSO_TO_OO_I_Interior_get_Pattern,
    MSO_TO_OO_I_Interior_put_Pattern,
    MSO_TO_OO_I_Interior_get_PatternColor,
    MSO_TO_OO_I_Interior_put_PatternColor,
    MSO_TO_OO_I_Interior_get_PatternColorIndex,
    MSO_TO_OO_I_Interior_put_PatternColorIndex
};

extern HRESULT _I_InteriorConstructor(LPVOID *ppObj)
{
    InteriorImpl *interior;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);
    
    interior = HeapAlloc(GetProcessHeap(), 0, sizeof(*interior));
    if (!interior)
    {
        return E_OUTOFMEMORY;
    }

    interior->_interiorVtbl = &MSO_TO_OO_I_Interior_Vtbl;
    interior->ref = 0;
    interior->prange = NULL;

    *ppObj = &interior->_interiorVtbl;
    
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}
