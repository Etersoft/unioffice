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

    TRACE("\n");

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
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Shape_GetTypeInfo(
        I_Shape* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Shape_GetIDsOfNames(
        I_Shape* iface,
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
    TRACE("%i not supported\n");
    return E_NOTIMPL;
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

ShapeImpl MSO_TO_OO_Shape =
{
    &MSO_TO_OO_I_ShapeVtbl,
    0,
    NULL,
    NULL
};


extern HRESULT _I_ShapeConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    ShapeImpl *shape;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

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

    return S_OK;
}
