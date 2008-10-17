/*
 * IBorders interface functions
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

ITypeInfo *ti_borders = NULL;

HRESULT get_typeinfo_borders(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_borders) {
        *typeinfo = ti_borders;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Borders, &ti_borders);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_borders;

    return hres;
}

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Borders_AddRef(
        I_Borders* iface)
{
    BordersImpl *This = (BordersImpl*)iface;
    ULONG ref;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_QueryInterface(
        I_Borders* iface,
        REFIID riid,
        void **ppvObject)
{
    BordersImpl *This = (BordersImpl*)iface;

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Borders)) {
        *ppvObject = &This->_bordersVtbl;
        MSO_TO_OO_I_Borders_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Borders_Release(
        I_Borders* iface)
{
    BordersImpl *This = (BordersImpl*)iface;
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

/*** I_Borders methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Application(
        I_Borders* iface,
        IDispatch **value)
{
    BordersImpl *This = (BordersImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    TRACE_OUT;
    return I_Range_get_Application((I_Range*)(This->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Parent(
        I_Borders* iface,
        IDispatch **value)
{
    BordersImpl *This = (BordersImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    *value = This->prange;
    I_Range_AddRef(This->prange);

    if (value==NULL)
        return E_POINTER;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Color(
        I_Borders* iface,
        long *plcolor)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;
    TRACE_IN;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_get_Color((I_Border*)border_tmp, plcolor);
    IDispatch_Release(border_tmp);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Color(
        I_Borders* iface,
        long lcolor)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;
    int i;
    TRACE_IN;
    TRACE(" lcolor = %i\n",lcolor);

    if (This==NULL) return E_POINTER;

    for (i=1;i<=12;i++) {
        if ((i==5)||(i==6)) continue;
        I_Borders_get_Item(iface, i, &border_tmp);
        I_Border_put_Color((I_Border*)border_tmp, lcolor);
        IDispatch_Release(border_tmp);
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_ColorIndex(
        I_Borders* iface,
        long *plcolorindex)
{
    BordersImpl *This = (BordersImpl*)iface;
    long tmpcolor;
    int i;
    HRESULT hres;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    hres = MSO_TO_OO_I_Borders_get_Color(iface,&tmpcolor);
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

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_ColorIndex(
        I_Borders* iface,
        long lcolorindex)
{
    BordersImpl *This = (BordersImpl*)iface;
    long tmpcolor;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;
    TRACE_OUT;
    if ((lcolorindex<1)||(lcolorindex>56)) {
        TRACE(" ERROR Incorrect colorindex %i\n", lcolorindex);
        return S_OK;
    } else 
        return MSO_TO_OO_I_Borders_put_Color(iface,color[lcolorindex-1]);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Creator(
        I_Borders* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_LineStyle(
        I_Borders* iface,
        XlLineStyle *plinestyle)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_get_LineStyle((I_Border*)border_tmp, plinestyle);
    IDispatch_Release(border_tmp);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_LineStyle(
        I_Borders* iface,
        XlLineStyle linestyle)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;
    int i;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    for (i=1;i<=12;i++) {
        if ((i==5)||(i==6)) continue;
        I_Borders_get_Item(iface, i, &border_tmp);
        I_Border_put_LineStyle((I_Border*)border_tmp, linestyle);
        IDispatch_Release(border_tmp);
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Weight(
        I_Borders* iface,
        XlBorderWeight *pweight)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_get_Weight((I_Border*)border_tmp, pweight);
    IDispatch_Release(border_tmp);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Weight(
        I_Borders* iface,
        XlBorderWeight weight)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;
    int i;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    for (i=1;i<=12;i++) {
        if ((i==5)||(i==6)) continue;
        I_Borders_get_Item(iface, i, &border_tmp);
        I_Border_put_Weight((I_Border*)border_tmp, weight);
        IDispatch_Release(border_tmp);
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get__Default(
        I_Borders* iface,
        XlBordersIndex key,
        IDispatch **ppObject)
{
    BordersImpl *This = (BordersImpl*)iface;
    IUnknown *punk = NULL;
    IDispatch *pborder;
    HRESULT hres;
    TRACE_IN;
    TRACE("key=%08x\n",key);

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }
    /*Создаем объект Border*/
    *ppObject = NULL;

    hres = _I_BorderConstructor((LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Border_QueryInterface(punk, &IID_I_Border, (void**) &pborder);
    if (pborder == NULL) {
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Border_Initialize((I_Border*)pborder, iface, key);

    if (FAILED(hres)) {
        IDispatch_Release(pborder);
        return hres;
    }

    *ppObject = pborder;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Item(
        I_Borders* iface,
        XlBordersIndex key,
        IDispatch **ppObject)
{
    TRACE("  ----> get__Default");
    return MSO_TO_OO_I_Borders_get__Default(iface, key, ppObject);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Value(
        I_Borders* iface,
        XlLineStyle *plinestyle)
{
    TRACE(" ----> get_LineStyle");
    return MSO_TO_OO_I_Borders_get_LineStyle(iface, plinestyle);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Value(
        I_Borders* iface,
        XlLineStyle linestyle)
{
    TRACE(" ----> put_LineStyle");
    return MSO_TO_OO_I_Borders_put_LineStyle(iface, linestyle);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Count(
        I_Borders* iface,
        long *pretval)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetEnumerator(
        I_Borders* iface,
        IDispatch **pdretval)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Borders_GetTypeInfoCount(
        I_Borders* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetTypeInfo(
        I_Borders* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_borders(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetIDsOfNames(
        I_Borders* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_borders(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_Invoke(
        I_Borders* iface,
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
    hres = get_typeinfo_borders(&typeinfo);
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

const I_BordersVtbl MSO_TO_OO_I_Borders_Vtbl =
{
    MSO_TO_OO_I_Borders_QueryInterface,
    MSO_TO_OO_I_Borders_AddRef,
    MSO_TO_OO_I_Borders_Release,
    MSO_TO_OO_I_Borders_GetTypeInfoCount,
    MSO_TO_OO_I_Borders_GetTypeInfo,
    MSO_TO_OO_I_Borders_GetIDsOfNames,
    MSO_TO_OO_I_Borders_Invoke,
    MSO_TO_OO_I_Borders_get_Application,
    MSO_TO_OO_I_Borders_get_Creator,
    MSO_TO_OO_I_Borders_get_Parent,
    MSO_TO_OO_I_Borders_get_Color,
    MSO_TO_OO_I_Borders_put_Color,
    MSO_TO_OO_I_Borders_get_ColorIndex,
    MSO_TO_OO_I_Borders_put_ColorIndex,
    MSO_TO_OO_I_Borders_get_Count,
    MSO_TO_OO_I_Borders_get_Item,
    MSO_TO_OO_I_Borders_get_LineStyle,
    MSO_TO_OO_I_Borders_put_LineStyle,
    MSO_TO_OO_I_Borders_GetEnumerator,
    MSO_TO_OO_I_Borders_get_Value,
    MSO_TO_OO_I_Borders_put_Value,
    MSO_TO_OO_I_Borders_get_Weight,
    MSO_TO_OO_I_Borders_put_Weight,
    MSO_TO_OO_I_Borders_get__Default
};

extern HRESULT _I_BordersConstructor(LPVOID *ppObj)
{
    BordersImpl *borders;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    borders = HeapAlloc(GetProcessHeap(), 0, sizeof(*borders));
    if (!borders)
    {
        return E_OUTOFMEMORY;
    }

    borders->_bordersVtbl = &MSO_TO_OO_I_Borders_Vtbl;
    borders->ref = 0;
    borders->prange = NULL;

    *ppObj = &borders->_bordersVtbl;
    TRACE_OUT;
    return S_OK;
}
