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

static WCHAR const str_addline[] = {
    'A','d','d','L','i','n','e',0};

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
    if (!lstrcmpiW(*rgszNames, str_addline)) {
        *rgDispId = 1;
        return S_OK;
    }
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
    HRESULT hres;
    VARIANT par1,par2,par3,par4;
    IDispatch *dret;

    TRACE("\n");

    VariantInit(&par1);
    VariantInit(&par2);
    VariantInit(&par3);
    VariantInit(&par4);

    if (iface == NULL) return E_POINTER;

    switch(dispIdMember) 
    {
    case 1://AddLine
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE("NOTIMPL\n");
            return E_NOTIMPL;
        } else {
            if (pDispParams->cArgs!=4) {
                TRACE("ERROR parameters\n");
                return E_FAIL;
            }
            hres = MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par1);
            if (FAILED(hres)) {
                TRACE("ERROR when CorrectArg par1 \n");
                return E_FAIL;
            }
            hres = VariantChangeTypeEx(&par1, &par1, 0, 0, VT_R4);
            if (FAILED(hres)) {
                TRACE("ERROR when VariantChangeTypeEx par1 \n");
                return E_FAIL;
            }
            hres = MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par2);
            if (FAILED(hres)) {
                TRACE("ERROR when CorrectArg par1 \n");
                return E_FAIL;
            }
            hres = VariantChangeTypeEx(&par2, &par2, 0, 0, VT_R4);
            if (FAILED(hres)) {
                TRACE("ERROR when VariantChangeTypeEx par1 \n");
                return E_FAIL;
            }
            hres = MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par3);
            if (FAILED(hres)) {
                TRACE("ERROR when CorrectArg par1 \n");
                return E_FAIL;
            }
            hres = VariantChangeTypeEx(&par3, &par3, 0, 0, VT_R4);
            if (FAILED(hres)) {
                TRACE("ERROR when VariantChangeTypeEx par1 \n");
                return E_FAIL;
            }
            hres = MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par4);
            if (FAILED(hres)) {
                TRACE("ERROR when CorrectArg par1 \n");
                return E_FAIL;
            }
            hres = VariantChangeTypeEx(&par4, &par4, 0, 0, VT_R4);
            if (FAILED(hres)) {
                TRACE("ERROR when VariantChangeTypeEx par1 \n");
                return E_FAIL;
            }

            hres = MSO_TO_OO_I_Shapes_AddLine(iface, V_R4(&par1), V_R4(&par2), V_R4(&par3), V_R4(&par4), &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
                return hres;
            }
            IDispatch_Release(dret);
            return E_FAIL;
        }
    }

    TRACE("%i not supported\n");
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
