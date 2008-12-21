/*
 * IPageSetup interface functions
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

ITypeInfo *ti_pagesetup = NULL;

HRESULT get_typeinfo_pagesetup(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_pagesetup) {
        *typeinfo = ti_pagesetup;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        ERR("LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_PageSetup, &ti_pagesetup);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_pagesetup;
    return hres;
}

#define PAGESETUP_THIS(iface) DEFINE_THIS(PageSetupImpl, pagesetup, iface)

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_PageSetup_AddRef(
        I_PageSetup* iface)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (!This) {
        ERR("Object is NULL \n");       
        return E_POINTER;
    }

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_QueryInterface(
        I_PageSetup* iface,
        REFIID riid,
        void **ppvObject)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);

    if (!This) {
        ERR("Object is NULL \n");       
        return E_POINTER;
    }
    
    if (!ppvObject) {
        ERR("Object is NULL \n");       
        return E_POINTER;
    }
    
    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_PageSetup)) {
        *ppvObject = PAGESETUP_PAGESETUP(This);
        I_PageSetup_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_PageSetup_Release(
        I_PageSetup* iface)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (!This) {
        ERR("Object is NULL \n");       
        return E_POINTER;
    }

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pWorksheet) {
            I_Worksheet_Release(This->pWorksheet);
            This->pWorksheet = NULL;
        }
        if (This->pOOSheet) {
            IDispatch_Release(This->pOOSheet);
            This->pOOSheet = NULL;
        }
        if (This->pOODocument) {
            IDispatch_Release(This->pOODocument);
            This->pOODocument = NULL;
        }       
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
        DELETE_OBJECT;
    }
    return ref;
}

/*** IPageSetup methods ***/

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");       
        return E_POINTER;
    }
    
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"LeftMargin",0);
    if (FAILED(hres)) {
        ERR("LeftMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        ERR("(1)VariantChangeType \n");
        return hres;
    }

    *value = V_R8(&vres)/1000*28;
    VarR8Round(*value, 0, value);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_LeftMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }

    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"LeftMargin",1, param1);
    if (FAILED(hres)) {
        ERR("LeftMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"RightMargin",0);
    if (FAILED(hres)) {
        ERR("RightMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    *value = V_R8(&vres)/1000*28;
    VarR8Round(*value, 0, value);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_RightMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"RightMargin",1, param1);
    if (FAILED(hres)) {
        ERR(" RightMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_TopMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"TopMargin",0);
    if (FAILED(hres)) {
        ERR("TopMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    *value = V_R8(&vres)/1000*28;
    VarR8Round(*value, 0, value);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_TopMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"TopMargin",1, param1);
    if (FAILED(hres)) {
        ERR("TopMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_BottomMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR(" Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"BottomMargin",0);
    if (FAILED(hres)) {
        ERR("BottomMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    *value = V_R8(&vres)/1000*28;
    VarR8Round(*value, 0, value);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_BottomMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"BottomMargin",1, param1);
    if (FAILED(hres)) {
        ERR("BottomMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Orientation(
        I_PageSetup* iface,
        long *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        TRACE("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"IsLandscape",0);
    if (FAILED(hres)) {
        ERR("Orientation \n");
        return hres;
    }
    switch (V_BOOL(&vres)){
    case VARIANT_TRUE:
        *value = xlLandscape;
        break;
    case VARIANT_FALSE:
        *value = xlPortrait;
        break;
    }


    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Orientation(
        I_PageSetup* iface,
        long value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    switch (value) {
    case xlLandscape:
        VariantClear(&param1);
        V_VT(&param1) = VT_BOOL;
        V_BOOL(&param1) = VARIANT_TRUE;
        break;
    case xlPortrait:
    default :
        VariantClear(&param1);
        V_VT(&param1) = VT_BOOL;
        V_BOOL(&param1) = VARIANT_FALSE;
        break;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"IsLandscape",1, param1);
    if (FAILED(hres)) {
        ERR("IsLandscape \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Zoom(
        I_PageSetup* iface,
        VARIANT *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"PageScale",0);
    if (FAILED(hres)) {
        ERR("PageScale \n");
        return hres;
    }

    *value = vres;

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Zoom(
        I_PageSetup* iface,
        VARIANT value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(value, &value);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"PageScale",1, value);
    if (FAILED(hres)) {
        ERR("PageScale \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FitToPagesTall(
        I_PageSetup* iface,
        VARIANT *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesY",0);
    if (FAILED(hres)) {
        ERR("ScaleToPagesY \n");
        return hres;
    }

    *value = vres;

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FitToPagesTall(
        I_PageSetup* iface,
        VARIANT value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(value, &value);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&value, &value, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR("VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesY",1, value);
    if (FAILED(hres)) {
        ERR(" ScaleToPagesY \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FitToPagesWide(
        I_PageSetup* iface,
        VARIANT *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesX",0);
    if (FAILED(hres)) {
        ERR(" ScaleToPagesX \n");
        return hres;
    }

    *value = vres;

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FitToPagesWide(
        I_PageSetup* iface,
        VARIANT value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(value, &value);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesX",1, value);
    if (FAILED(hres)) {
        ERR(" ScaleToPagesX \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_HeaderMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR(" StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR(" getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"HeaderHeight",0);
    if (FAILED(hres)) {
        ERR(" HeaderHeight \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    *value = V_R8(&vres)/1000*28;
    VarR8Round(*value, 0, value);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_HeaderMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR(" PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR(" StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR(" getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"HeaderHeight",1, param1);
    if (FAILED(hres)) {
        ERR(" HeaderHeight \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FooterMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR(" PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR(" StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR(" getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR(" getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"FooterHeight",0);
    if (FAILED(hres)) {
        ERR(" FooterHeight \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    *value = V_R8(&vres)/1000*28;
    VarR8Round(*value, 0, value);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FooterMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR(" PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR(" StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR(" getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR(" getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR(" (1) VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"FooterHeight",1, param1);
    if (FAILED(hres)) {
        ERR(" FooterHeight \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterHorizontally(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"CenterHorizontally",0);
    if (FAILED(hres)) {
        ERR(" CenterHorizontally \n");
        return hres;
    }

    *value = V_BOOL(&vres);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterHorizontally(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR("PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR(" StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR(" getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_BOOL;
    V_BOOL(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"CenterHorizontally",1, param1);
    if (FAILED(hres)) {
        ERR(" CenterHorizontally \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterVertically(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR(" PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR(" getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"CenterVertically",0);
    if (FAILED(hres)) {
        ERR(" CenterVertically \n");
        return hres;
    }

    *value = V_BOOL(&vres);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterVertically(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    TRACE_IN;

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, This->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        ERR(" PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, This->pOODocument, L"StyleFamilies",0);
    if (FAILED(hres)) {
        ERR("StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        ERR("getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        ERR("getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_BOOL;
    V_BOOL(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"CenterVertically",1, param1);
    if (FAILED(hres)) {
        ERR("CenterVertically \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);
    TRACE_OUT;
    return S_OK;
}


static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintTitleRows(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_IN;
    TRACE_NOTIMPL;
    *value = SysAllocString(L"");
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintTitleRows(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Application(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_IN;
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    HRESULT hres;
    
    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    
    hres = I_Worksheet_get_Application(This->pWorksheet, value);
    
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Creator(
        I_PageSetup* iface,
        VARIANT *result)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Parent(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_IN;
    PageSetupImpl *This = PAGESETUP_THIS(iface);
    
    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    if (!value) {
        ERR("Object2 is NULL \n");
        return E_POINTER;
    }

    *value = (IDispatch*)(This->pWorksheet);
    I_Worksheet_AddRef(This->pWorksheet);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_BlackAndWhite(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_BlackAndWhite(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}


static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterFooter(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterFooter(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterHeader(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterHeader(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_ChartSize(
        I_PageSetup* iface,
        XlObjectSize *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_ChartSize(
        I_PageSetup* iface,
        XlObjectSize value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Draft(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Draft(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FirstPageNumber(
        I_PageSetup* iface,
        long *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FirstPageNumber(
        I_PageSetup* iface,
        long value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftFooter(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_LeftFooter(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftHeader(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_LeftHeader(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Order(
        I_PageSetup* iface,
        XlOrder *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Order(
        I_PageSetup* iface,
        XlOrder value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PaperSize(
        I_PageSetup* iface,
        XlPaperSize *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PaperSize(
        I_PageSetup* iface,
        XlPaperSize value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintArea(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintArea(
        I_PageSetup* iface,
        BSTR value)
{
    PageSetupImpl *This = PAGESETUP_THIS(iface); 
    HRESULT hres;
    VARIANT celladdress, vret, cell1, vNull, param1;
    I_Range* range;
    int i, count;
    I_Worksheet* wrksh;    
    SAFEARRAY FAR* pPropVals;
    long ix = 0;
    
    VariantInit(&celladdress);
    VariantInit(&vret);
    VariantInit(&cell1); 
    VariantInit(&param1);       
    VariantInit(&vNull);   
    V_VT(&vNull) = VT_NULL; 
          
    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;           
    } 
      
    hres = I_PageSetup_get_Parent(iface, (IDispatch**)&wrksh);
    if (FAILED(hres)) {
        ERR("PageSetup_get_Parent\n");
        return E_FAIL;
    }     
       
    V_VT(&cell1) = VT_BSTR;
    V_BSTR(&cell1) = SysAllocString(value);  
      
    hres = I_Worksheet_get_Range(wrksh, cell1, vNull, (IDispatch**)&range);  
    if (FAILED(hres)) {
        ERR("Worksheet_get_Range\n");
        return E_FAIL;
    }      
      
    RangeImpl* rangeimpl = (RangeImpl*)range;  
      
    hres = AutoWrap(DISPATCH_METHOD, &celladdress, rangeimpl->pOORange, L"getRangeAddress", 0);
    if (FAILED(hres)) {
        ERR("getRangeAddress\n");
        if (wrksh) {
            I_Worksheet_Release(wrksh);
            wrksh = NULL;           
        }
        return E_FAIL;
    }
    
    pPropVals = SafeArrayCreateVector( VT_VARIANT, 0, 1 );

    hres = SafeArrayPutElement( pPropVals, &ix, &celladdress );

    V_VT(&param1) = VT_ARRAY | VT_VARIANT;
    V_ARRAY(&param1) = pPropVals;
    
    hres = AutoWrap(DISPATCH_METHOD, &vret, This->pOOSheet, L"setPrintAreas", 1, param1);
    if (FAILED(hres)) {
        ERR("setPrintAreas \n");
    }   

    VariantClear(&celladdress);
    VariantClear(&vret);
    VariantClear(&param1);
    VariantClear(&cell1);
    if (wrksh) {
        I_Worksheet_Release(wrksh);
        wrksh = NULL;           
    }
    if (range) {
        I_Range_Release(range);
        range = NULL;           
    }    
        
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintGridlines(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintGridlines(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintHeadings(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintHeadings(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintNotes(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintNotes(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintQuality(
        I_PageSetup* iface,
        VARIANT index,
        VARIANT *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintQuality(
        I_PageSetup* iface,
        VARIANT index,
        VARIANT *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintTitleColumns(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintTitleColumns(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightFooter(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_RightFooter(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightHeader(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_RightHeader(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintComments(
        I_PageSetup* iface,
        XlPrintLocation *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintComments(
        I_PageSetup* iface,
        XlPrintLocation value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintErrors(
        I_PageSetup* iface,
        XlPrintErrors *value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintErrors(
        I_PageSetup* iface,
        XlPrintErrors value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterHeaderPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterFooterPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftHeaderPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftFooterPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightHeaderPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightFooterPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_PageSetup_GetTypeInfoCount(
        I_PageSetup* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_GetTypeInfo(
        I_PageSetup* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_pagesetup(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        ERR("GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_GetIDsOfNames(
        I_PageSetup* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_pagesetup(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WERR(L" name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_Invoke(
        I_PageSetup* iface,
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
    hres = get_typeinfo_pagesetup(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
    if (FAILED(hres)) {
        ERR("wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
    }
    TRACE_OUT;
    return hres;
}

#undef PAGESETUP_THIS

const I_PageSetupVtbl MSO_TO_OO_I_PageSetupVtbl =
{
    MSO_TO_OO_I_PageSetup_QueryInterface,
    MSO_TO_OO_I_PageSetup_AddRef,
    MSO_TO_OO_I_PageSetup_Release,
    MSO_TO_OO_I_PageSetup_GetTypeInfoCount,
    MSO_TO_OO_I_PageSetup_GetTypeInfo,
    MSO_TO_OO_I_PageSetup_GetIDsOfNames,
    MSO_TO_OO_I_PageSetup_Invoke,
    MSO_TO_OO_I_PageSetup_get_Application,
    MSO_TO_OO_I_PageSetup_get_Creator,
    MSO_TO_OO_I_PageSetup_get_Parent,
    MSO_TO_OO_I_PageSetup_get_BlackAndWhite,
    MSO_TO_OO_I_PageSetup_put_BlackAndWhite,
    MSO_TO_OO_I_PageSetup_get_BottomMargin,
    MSO_TO_OO_I_PageSetup_put_BottomMargin,
    MSO_TO_OO_I_PageSetup_get_CenterFooter,
    MSO_TO_OO_I_PageSetup_put_CenterFooter,
    MSO_TO_OO_I_PageSetup_get_CenterHeader,
    MSO_TO_OO_I_PageSetup_put_CenterHeader,
    MSO_TO_OO_I_PageSetup_get_CenterHorizontally,
    MSO_TO_OO_I_PageSetup_put_CenterHorizontally,
    MSO_TO_OO_I_PageSetup_get_CenterVertically,
    MSO_TO_OO_I_PageSetup_put_CenterVertically,
    MSO_TO_OO_I_PageSetup_get_ChartSize,
    MSO_TO_OO_I_PageSetup_put_ChartSize,
    MSO_TO_OO_I_PageSetup_get_Draft,
    MSO_TO_OO_I_PageSetup_put_Draft,
    MSO_TO_OO_I_PageSetup_get_FirstPageNumber,
    MSO_TO_OO_I_PageSetup_put_FirstPageNumber,
    MSO_TO_OO_I_PageSetup_get_FitToPagesTall,
    MSO_TO_OO_I_PageSetup_put_FitToPagesTall,
    MSO_TO_OO_I_PageSetup_get_FitToPagesWide,
    MSO_TO_OO_I_PageSetup_put_FitToPagesWide,
    MSO_TO_OO_I_PageSetup_get_FooterMargin,
    MSO_TO_OO_I_PageSetup_put_FooterMargin,
    MSO_TO_OO_I_PageSetup_get_HeaderMargin,
    MSO_TO_OO_I_PageSetup_put_HeaderMargin,
    MSO_TO_OO_I_PageSetup_get_LeftFooter,
    MSO_TO_OO_I_PageSetup_put_LeftFooter,
    MSO_TO_OO_I_PageSetup_get_LeftHeader,
    MSO_TO_OO_I_PageSetup_put_LeftHeader,
    MSO_TO_OO_I_PageSetup_get_LeftMargin,
    MSO_TO_OO_I_PageSetup_put_LeftMargin,
    MSO_TO_OO_I_PageSetup_get_Order,
    MSO_TO_OO_I_PageSetup_put_Order,
    MSO_TO_OO_I_PageSetup_get_Orientation,
    MSO_TO_OO_I_PageSetup_put_Orientation,
    MSO_TO_OO_I_PageSetup_get_PaperSize,
    MSO_TO_OO_I_PageSetup_put_PaperSize,
    MSO_TO_OO_I_PageSetup_get_PrintArea,
    MSO_TO_OO_I_PageSetup_put_PrintArea,
    MSO_TO_OO_I_PageSetup_get_PrintGridlines,
    MSO_TO_OO_I_PageSetup_put_PrintGridlines,
    MSO_TO_OO_I_PageSetup_get_PrintHeadings,
    MSO_TO_OO_I_PageSetup_put_PrintHeadings,
    MSO_TO_OO_I_PageSetup_get_PrintNotes,
    MSO_TO_OO_I_PageSetup_put_PrintNotes,
    MSO_TO_OO_I_PageSetup_get_PrintQuality,
    MSO_TO_OO_I_PageSetup_put_PrintQuality,
    MSO_TO_OO_I_PageSetup_get_PrintTitleColumns,
    MSO_TO_OO_I_PageSetup_put_PrintTitleColumns,
    MSO_TO_OO_I_PageSetup_get_PrintTitleRows,
    MSO_TO_OO_I_PageSetup_put_PrintTitleRows,
    MSO_TO_OO_I_PageSetup_get_RightFooter,
    MSO_TO_OO_I_PageSetup_put_RightFooter,
    MSO_TO_OO_I_PageSetup_get_RightHeader,
    MSO_TO_OO_I_PageSetup_put_RightHeader,
    MSO_TO_OO_I_PageSetup_get_RightMargin,
    MSO_TO_OO_I_PageSetup_put_RightMargin,
    MSO_TO_OO_I_PageSetup_get_TopMargin,
    MSO_TO_OO_I_PageSetup_put_TopMargin,
    MSO_TO_OO_I_PageSetup_get_Zoom,
    MSO_TO_OO_I_PageSetup_put_Zoom,
    MSO_TO_OO_I_PageSetup_get_PrintComments,
    MSO_TO_OO_I_PageSetup_put_PrintComments,
    MSO_TO_OO_I_PageSetup_get_PrintErrors,
    MSO_TO_OO_I_PageSetup_put_PrintErrors,
    MSO_TO_OO_I_PageSetup_get_CenterHeaderPicture,
    MSO_TO_OO_I_PageSetup_get_CenterFooterPicture,
    MSO_TO_OO_I_PageSetup_get_LeftHeaderPicture,
    MSO_TO_OO_I_PageSetup_get_LeftFooterPicture,
    MSO_TO_OO_I_PageSetup_get_RightHeaderPicture,
    MSO_TO_OO_I_PageSetup_get_RightFooterPicture
};

extern HRESULT _I_PageSetupConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    PageSetupImpl *pagesetup;
    TRACE_IN;
    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

    pagesetup = HeapAlloc(GetProcessHeap(), 0, sizeof(*pagesetup));
    if (!pagesetup)
    {
        return E_OUTOFMEMORY;
    }

    pagesetup->ppagesetupVtbl = &MSO_TO_OO_I_PageSetupVtbl;
    pagesetup->ref = 0;
    pagesetup->pWorksheet = NULL;
    pagesetup->pOOSheet = NULL;   
    pagesetup->pOODocument = NULL;
    
    *ppObj = PAGESETUP_PAGESETUP(pagesetup);
    
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}
