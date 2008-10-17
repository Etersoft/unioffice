/*
 * IShape interface functions
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

ITypeInfo *ti_shape = NULL;

HRESULT get_typeinfo_shape(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if(ti_shape) {
        *typeinfo = ti_shape;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Shape, &ti_shape);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_shape;
    return hres;
}


/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Shape_AddRef(
        I_Shape* iface)
{
    ShapeImpl *This = (ShapeImpl*)iface;
    ULONG ref;

    TRACE("REF=%i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Shape_QueryInterface(
        I_Shape* iface,
        REFIID riid,
        void **ppvObject)
{
    ShapeImpl *This = (ShapeImpl*)iface;

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Shape)) {
        *ppvObject = &This->_shapeVtbl;
        MSO_TO_OO_I_Shape_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Shape_Release(
        I_Shape* iface)
{
    ShapeImpl *This = (ShapeImpl*)iface;
    ULONG ref;

    TRACE("REF=%i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pOOShape != NULL) {
            IDispatch_Release(This->pOOShape);
            This->pOOShape = NULL;
        }
        if (This->pShapes != NULL) {
            I_Workbook_Release(This->pShapes);
            This->pShapes = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Shape methods ***/



/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Shape_GetTypeInfoCount(
        I_Shape* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Shape_GetTypeInfo(
        I_Shape* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_shape(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Shape_GetIDsOfNames(
        I_Shape* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_shape(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Shape_Invoke(
        I_Shape* iface,
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

    hres = get_typeinfo_shape(&typeinfo);
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

const I_ShapeVtbl MSO_TO_OO_I_ShapeVtbl =
{
    MSO_TO_OO_I_Shape_QueryInterface,
    MSO_TO_OO_I_Shape_AddRef,
    MSO_TO_OO_I_Shape_Release,
    MSO_TO_OO_I_Shape_GetTypeInfoCount,
    MSO_TO_OO_I_Shape_GetTypeInfo,
    MSO_TO_OO_I_Shape_GetIDsOfNames,
    MSO_TO_OO_I_Shape_Invoke
};

extern HRESULT _I_ShapeConstructor(LPVOID *ppObj)
{
    ShapeImpl *shape;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    shape = HeapAlloc(GetProcessHeap(), 0, sizeof(*shape));
    if (!shape)
    {
        return E_OUTOFMEMORY;
    }

    shape->_shapeVtbl = &MSO_TO_OO_I_ShapeVtbl;
    shape->ref = 0;
    shape->pOOShape = NULL;
    shape->pShapes = NULL;

    *ppObj = &shape->_shapeVtbl;
    TRACE_OUT;
    return S_OK;
}
