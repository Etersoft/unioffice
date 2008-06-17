/*
 * IShapes interface functions
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


/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Shapes_AddRef(
        I_Shapes* iface)
{
    ShapesImpl *This = (ShapesImpl*)iface;
    ULONG ref;

    TRACE("REF=%i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Shapes_QueryInterface(
        I_Shapes* iface,
        REFIID riid,
        void **ppvObject)
{
    ShapesImpl *This = (ShapesImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Shapes)) {
        *ppvObject = &This->_shapesVtbl;
        MSO_TO_OO_I_Shapes_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Shapes_Release(
        I_Shapes* iface)
{
    ShapesImpl *This = (ShapesImpl*)iface;
    ULONG ref;

    TRACE("REF=%i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pOOPage != NULL) {
            IDispatch_Release(This->pOOPage);
            This->pOOPage = NULL;
        }
        if (This->pwsheet != NULL) {
            I_Workbook_Release(This->pwsheet);
            This->pwsheet = NULL;
        }
        if (This->pApplication != NULL) {
            IDispatch_Release(This->pApplication);
            This->pApplication = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Shapes methods ***/





/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Shapes_GetTypeInfoCount(
        I_Shapes* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Shapes_GetTypeInfo(
        I_Shapes* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Shapes_GetIDsOfNames(
        I_Shapes* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L" %s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Shapes_Invoke(
        I_Shapes* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    TRACE("\n");
    return E_NOTIMPL;
}

const I_ShapesVtbl MSO_TO_OO_I_ShapesVtbl =
{
    MSO_TO_OO_I_Shapes_QueryInterface,
    MSO_TO_OO_I_Shapes_AddRef,
    MSO_TO_OO_I_Shapes_Release,
    MSO_TO_OO_I_Shapes_GetTypeInfoCount,
    MSO_TO_OO_I_Shapes_GetTypeInfo,
    MSO_TO_OO_I_Shapes_GetIDsOfNames,
    MSO_TO_OO_I_Shapes_Invoke
};

ShapesImpl MSO_TO_OO_Shapes =
{
    &MSO_TO_OO_I_ShapesVtbl,
    0
};


extern HRESULT _I_ShapesConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    ShapesImpl *shapes;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

    shapes = HeapAlloc(GetProcessHeap(), 0, sizeof(*shapes));
    if (!shapes)
    {
        return E_OUTOFMEMORY;
    }

    shapes->_shapesVtbl = &MSO_TO_OO_I_ShapesVtbl;
    shapes->ref = 0;
    shapes->pOOPage = NULL;
    shapes->pwsheet = NULL;
    shapes->pApplication = NULL;

    *ppObj = &shapes->_shapesVtbl;

    return S_OK;
}
