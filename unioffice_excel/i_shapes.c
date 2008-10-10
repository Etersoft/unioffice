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

ITypeInfo *ti_shapes = NULL;

HRESULT get_typeinfo_shapes(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if(ti_shapes) {
        *typeinfo = ti_shapes;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Shapes, &ti_shapes);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_shapes;
    return hres;
}

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

static HRESULT WINAPI MSO_TO_OO_I_Shapes_AddLine(
        I_Shapes* iface,
        float beginX,
        float beginY,
        float endX,
        float endY,
        IDispatch **ppValue)
{
    ShapesImpl *This = (ShapesImpl*)iface;
    HRESULT hres;
    IUnknown *pObj;
    TRACE_IN;
    TRACE("%f;%f;%f;%f\n",beginX, beginY, endX, endY);

    hres = _I_ShapeConstructor((void**)&pObj);
    if (FAILED(hres)) {
        TRACE(" ERROR when call constructor IShape\n");
        return E_FAIL;
    }

    hres = I_Shape_QueryInterface(pObj, &IID_I_Shape, (void**)ppValue);
    if (FAILED(hres)) {
        TRACE(" ERROR when call IShape->QueryInterface\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Shape_Line_Initialize((I_Shape*)*ppValue, iface, beginX*10, beginY*10, endX*10, endY*10);
    if (FAILED(hres)) {
        TRACE(" ERROR when call Shape_Line initialize\n");
        return E_FAIL;
    }
    TRACE_OUT;
    return S_OK;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Shapes_GetTypeInfoCount(
        I_Shapes* iface,
        UINT *pctinfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Shapes_GetTypeInfo(
        I_Shapes* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE_NOTIMPL;
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
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_shapes(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
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
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_shapes(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
    if (FAILED(hres)) {
        TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
    }
    TRACE_OUT;
    return hres;
}

const I_ShapesVtbl MSO_TO_OO_I_ShapesVtbl =
{
    MSO_TO_OO_I_Shapes_QueryInterface,
    MSO_TO_OO_I_Shapes_AddRef,
    MSO_TO_OO_I_Shapes_Release,
    MSO_TO_OO_I_Shapes_GetTypeInfoCount,
    MSO_TO_OO_I_Shapes_GetTypeInfo,
    MSO_TO_OO_I_Shapes_GetIDsOfNames,
    MSO_TO_OO_I_Shapes_Invoke,
    MSO_TO_OO_I_Shapes_AddLine
};

extern HRESULT _I_ShapesConstructor(LPVOID *ppObj)
{
    ShapesImpl *shapes;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

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
    TRACE_OUT;
    return S_OK;
}
