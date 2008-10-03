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
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_PageSetup, &ti_pagesetup);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_pagesetup;
    return hres;
}

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_PageSetup_AddRef(
        I_PageSetup* iface)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

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
    PageSetupImpl *This = (PageSetupImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_PageSetup)) {
        *ppvObject = &This->_pagesetupVtbl;
        MSO_TO_OO_I_PageSetup_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_PageSetup_Release(
        I_PageSetup* iface)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pwsheet != NULL) {
            IDispatch_Release(This->pwsheet);
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

/*** IPageSetup methods ***/

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"LeftMargin",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  LeftMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_LeftMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }

    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"LeftMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when LeftMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"RightMargin",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  RightMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_RightMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"RightMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when  RightMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_TopMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"TopMargin",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  TopMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_TopMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"TopMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when  TopMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_BottomMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"BottomMargin",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  BottomMargin \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_BottomMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"BottomMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when  BottomMargin \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Orientation(
        I_PageSetup* iface,
        long *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"IsLandscape",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  Orientation \n");
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Orientation(
        I_PageSetup* iface,
        long value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
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
        TRACE("ERROR when  IsLandscape \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Zoom(
        I_PageSetup* iface,
        VARIANT *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"PageScale",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  PageScale \n");
        return hres;
    }

    *value = vres;

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Zoom(
        I_PageSetup* iface,
        VARIANT value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"PageScale",1, value);
    if (FAILED(hres)) {
        TRACE("ERROR when  PageScale \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FitToPagesTall(
        I_PageSetup* iface,
        VARIANT *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesY",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  ScaleToPagesY \n");
        return hres;
    }

    *value = vres;

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FitToPagesTall(
        I_PageSetup* iface,
        VARIANT value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&value, &value, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesY",1, value);
    if (FAILED(hres)) {
        TRACE("ERROR when  ScaleToPagesY \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FitToPagesWide(
        I_PageSetup* iface,
        VARIANT *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesX",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  ScaleToPagesX \n");
        return hres;
    }

    *value = vres;

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FitToPagesWide(
        I_PageSetup* iface,
        VARIANT value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesX",1, value);
    if (FAILED(hres)) {
        TRACE("ERROR when  ScaleToPagesX \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_HeaderMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"HeaderHeight",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  HeaderHeight \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_HeaderMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"HeaderHeight",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when  HeaderHeight \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FooterMargin(
        I_PageSetup* iface,
        double *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"FooterHeight",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  FooterHeight \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vres, &vres, 0, 0, VT_R8);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FooterMargin(
        I_PageSetup* iface,
        double value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R8;
    value = value/28*1000; /*т.к. OpenOffice измеряет в 1/100мм, а MS Office в точках 1/28 см.*/
    V_R8(&param1) = value;
    hres = VariantChangeTypeEx(&param1, &param1, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE(" (1) ERROR when VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"FooterHeight",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when  FooterHeight \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterHorizontally(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"CenterHorizontally",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  CenterHorizontally \n");
        return hres;
    }

    *value = V_BOOL(&vres);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterHorizontally(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_BOOL;
    V_BOOL(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"CenterHorizontally",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when CenterHorizontally \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterVertically(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"CenterVertically",0);
    if (FAILED(hres)) {
        TRACE("ERROR when  CenterVertically \n");
        return hres;
    }

    *value = V_BOOL(&vres);

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterVertically(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    HRESULT hres;
    VARIANT name_of_style, vstyles, vpagestyles, param1, vstyle, vres;
    WorksheetImpl *wsh = (WorksheetImpl *)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VariantInit(&name_of_style);
    VariantInit(&vstyles);
    VariantInit(&vpagestyles);
    VariantInit(&param1);
    VariantInit(&vstyle);
    VariantInit(&vres);
    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_BOOL;
    V_BOOL(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"CenterVertically",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when CenterVertically \n");
        return hres;
    }

    VariantClear(&name_of_style);
    VariantClear(&vstyles);
    VariantClear(&vpagestyles);
    VariantClear(&param1);
    VariantClear(&vstyle);
    VariantClear(&vres);

    return S_OK;
}


static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintTitleRows(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("NEED to implement but return S_OK and empty string \n");
    *value = SysAllocString(L"");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintTitleRows(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("NEED to implement but return S_OK\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Application(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Creator(
        I_PageSetup* iface,
        VARIANT *result)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Parent(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_BlackAndWhite(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_BlackAndWhite(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE("\n");
    return E_NOTIMPL;
}


static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterFooter(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterFooter(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterHeader(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_CenterHeader(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_ChartSize(
        I_PageSetup* iface,
        XlObjectSize *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_ChartSize(
        I_PageSetup* iface,
        XlObjectSize value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Draft(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Draft(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_FirstPageNumber(
        I_PageSetup* iface,
        long *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_FirstPageNumber(
        I_PageSetup* iface,
        long value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftFooter(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_LeftFooter(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftHeader(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_LeftHeader(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_Order(
        I_PageSetup* iface,
        XlOrder *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_Order(
        I_PageSetup* iface,
        XlOrder value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PaperSize(
        I_PageSetup* iface,
        XlPaperSize *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PaperSize(
        I_PageSetup* iface,
        XlPaperSize value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintArea(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintArea(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintGridlines(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintGridlines(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintHeadings(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintHeadings(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintNotes(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintNotes(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintQuality(
        I_PageSetup* iface,
        VARIANT index,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintQuality(
        I_PageSetup* iface,
        VARIANT index,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintTitleColumns(
        I_PageSetup* iface,
        VARIANT_BOOL *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintTitleColumns(
        I_PageSetup* iface,
        VARIANT_BOOL value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightFooter(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_RightFooter(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightHeader(
        I_PageSetup* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_RightHeader(
        I_PageSetup* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintComments(
        I_PageSetup* iface,
        XlPrintLocation *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintComments(
        I_PageSetup* iface,
        XlPrintLocation value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_PrintErrors(
        I_PageSetup* iface,
        XlPrintErrors *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_put_PrintErrors(
        I_PageSetup* iface,
        XlPrintErrors value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterHeaderPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_CenterFooterPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftHeaderPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_LeftFooterPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightHeaderPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_get_RightFooterPicture(
        I_PageSetup* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_PageSetup_GetTypeInfoCount(
        I_PageSetup* iface,
        UINT *pctinfo)
{
    TRACE("NEED to implement \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_GetTypeInfo(
        I_PageSetup* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
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

    hres = get_typeinfo_pagesetup(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }

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
    HRESULT hres;
    double dval;
    long lval;
    VARIANT vval,vtmp;
    VARIANT_BOOL vbval;
    BSTR bstr_val;

    VariantInit(&vval);
    VariantInit(&vtmp);

    TRACE(" \n");

    switch(dispIdMember)
    {
    case dispid_pagesetup_leftmargin://leftmargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (1) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE(" (1) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
            VarR8Round(dval, 0, &dval);
            hres = MSO_TO_OO_I_PageSetup_put_LeftMargin(iface, dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_LeftMargin(iface, &dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_R8;
                V_R8(pVarResult) = dval;
            }
            return hres;
        }
    case dispid_pagesetup_rightmargin://rightmargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (2) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE(" (2) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
            VarR8Round(dval, 0, &dval);
            hres = MSO_TO_OO_I_PageSetup_put_RightMargin(iface, dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_RightMargin(iface, &dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_R8;
                V_R8(pVarResult) = dval;
            }
            return hres;
        }
    case dispid_pagesetup_topmargin://topmargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (3) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE("(3) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
            VarR8Round(dval, 0, &dval);
            hres = MSO_TO_OO_I_PageSetup_put_TopMargin(iface, dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_TopMargin(iface, &dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_R8;
                V_R8(pVarResult) = dval;
            }
            return hres;
        }
    case dispid_pagesetup_bottommargin://bottommargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (4) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE(" (4) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
            VarR8Round(dval, 0, &dval);
            hres = MSO_TO_OO_I_PageSetup_put_BottomMargin(iface, dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_BottomMargin(iface, &dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_R8;
                V_R8(pVarResult) = dval;
            }
            return hres;
        }
    case dispid_pagesetup_orientation://orientation
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (5) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (5) ERROR when VariantChangeType \n");
                return hres;
            }
            lval = V_I4(&vtmp);
            hres = MSO_TO_OO_I_PageSetup_put_Orientation(iface, lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_Orientation(iface, &lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lval;
            }
            return hres;
        }
    case dispid_pagesetup_zoom://zoom
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (6) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = MSO_TO_OO_I_PageSetup_put_Zoom(iface, vtmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_Zoom(iface, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return hres;
        }
    case dispid_pagesetup_fittopagestall://FitToPagesTall
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (7) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = MSO_TO_OO_I_PageSetup_put_FitToPagesTall(iface, vtmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_FitToPagesTall(iface, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return hres;
        }
    case dispid_pagesetup_fittopageswide://FitToPagesWide
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (8) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = MSO_TO_OO_I_PageSetup_put_FitToPagesWide(iface, vtmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_FitToPagesWide(iface, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return hres;
        }
    case dispid_pagesetup_headermargin://HeaderMargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (9) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE(" (9) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
            VarR8Round(dval, 0, &dval);
            hres = MSO_TO_OO_I_PageSetup_put_HeaderMargin(iface, dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_HeaderMargin(iface, &dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_R8;
                V_R8(pVarResult) = dval;
            }
            return hres;
        }
    case dispid_pagesetup_footermargin://FooterMargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (10) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE(" (10) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
            VarR8Round(dval, 0, &dval);
            hres = MSO_TO_OO_I_PageSetup_put_FooterMargin(iface, dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_FooterMargin(iface, &dval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_R8;
                V_R8(pVarResult) = dval;
            }
            return hres;
        }
    case dispid_pagesetup_centerhorizontall://CenterHorizontally
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (11) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE(" (11) ERROR when VariantChangeType \n");
                return hres;
            }
            vbval = V_BOOL(&vtmp);
            hres = MSO_TO_OO_I_PageSetup_put_CenterHorizontally(iface, vbval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_CenterHorizontally(iface, &vbval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = vbval;
            }
            return hres;
        }
    case dispid_pagesetup_centervertically://CenterVertically
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (12) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
    HRESULT (STDMETHODCALLTYPE *get_Application)(
        I_PageSetup* This,
        IDispatch **value);

    HRESULT (STDMETHODCALLTYPE *get_Creator)(
        I_PageSetup* This,
        VARIANT *result);

    HRESULT (STDMETHODCALLTYPE *get_Parent)(
        I_PageSetup* This,
        IDispatch **value);

    HRESULT (STDMETHODCALLTYPE *get_BlackAndWhite)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_BlackAndWhite)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_BottomMargin)(
        I_PageSetup* This,
        double *value);

    HRESULT (STDMETHODCALLTYPE *put_BottomMargin)(
        I_PageSetup* This,
        double value);

    HRESULT (STDMETHODCALLTYPE *get_CenterFooter)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_CenterFooter)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_CenterHeader)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_CenterHeader)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_CenterHorizontally)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_CenterHorizontally)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_CenterVertically)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_CenterVertically)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_ChartSize)(
        I_PageSetup* This,
        XlObjectSize *value);

    HRESULT (STDMETHODCALLTYPE *put_ChartSize)(
        I_PageSetup* This,
        XlObjectSize value);

    HRESULT (STDMETHODCALLTYPE *get_Draft)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_Draft)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_FirstPageNumber)(
        I_PageSetup* This,
        long *value);

    HRESULT (STDMETHODCALLTYPE *put_FirstPageNumber)(
        I_PageSetup* This,
        long value);

    HRESULT (STDMETHODCALLTYPE *get_FitToPagesTall)(
        I_PageSetup* This,
        VARIANT *value);

    HRESULT (STDMETHODCALLTYPE *put_FitToPagesTall)(
        I_PageSetup* This,
        VARIANT value);

    HRESULT (STDMETHODCALLTYPE *get_FitToPagesWide)(
        I_PageSetup* This,
        VARIANT *value);

    HRESULT (STDMETHODCALLTYPE *put_FitToPagesWide)(
        I_PageSetup* This,
        VARIANT value);

    HRESULT (STDMETHODCALLTYPE *get_FooterMargin)(
        I_PageSetup* This,
        double *value);

    HRESULT (STDMETHODCALLTYPE *put_FooterMargin)(
        I_PageSetup* This,
        double value);

    HRESULT (STDMETHODCALLTYPE *get_HeaderMargin)(
        I_PageSetup* This,
        double *value);

    HRESULT (STDMETHODCALLTYPE *put_HeaderMargin)(
        I_PageSetup* This,
        double value);

    HRESULT (STDMETHODCALLTYPE *get_LeftFooter)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_LeftFooter)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_LeftHeader)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_LeftHeader)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_LeftMargin)(
        I_PageSetup* This,
        double *value);

    HRESULT (STDMETHODCALLTYPE *put_LeftMargin)(
        I_PageSetup* This,
        double value);

    HRESULT (STDMETHODCALLTYPE *get_Order)(
        I_PageSetup* This,
        XlOrder *value);

    HRESULT (STDMETHODCALLTYPE *put_Order)(
        I_PageSetup* This,
        XlOrder value);

    HRESULT (STDMETHODCALLTYPE *get_Orientation)(
        I_PageSetup* This,
        long *value);

    HRESULT (STDMETHODCALLTYPE *put_Orientation)(
        I_PageSetup* This,
        long value);

    HRESULT (STDMETHODCALLTYPE *get_PaperSize)(
        I_PageSetup* This,
        XlPaperSize *value);

    HRESULT (STDMETHODCALLTYPE *put_PaperSize)(
        I_PageSetup* This,
        XlPaperSize value);

    HRESULT (STDMETHODCALLTYPE *get_PrintArea)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintArea)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_PrintGridlines)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintGridlines)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_PrintHeadings)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintHeadings)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_PrintNotes)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintNotes)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_PrintQuality)(
        I_PageSetup* This,
        VARIANT index,
        VARIANT *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintQuality)(
        I_PageSetup* This,
        VARIANT index,
        VARIANT *value);

    HRESULT (STDMETHODCALLTYPE *get_PrintTitleColumns)(
        I_PageSetup* This,
        VARIANT_BOOL *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintTitleColumns)(
        I_PageSetup* This,
        VARIANT_BOOL value);

    HRESULT (STDMETHODCALLTYPE *get_PrintTitleRows)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintTitleRows)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_RightFooter)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_RightFooter)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_RightHeader)(
        I_PageSetup* This,
        BSTR *value);

    HRESULT (STDMETHODCALLTYPE *put_RightHeader)(
        I_PageSetup* This,
        BSTR value);

    HRESULT (STDMETHODCALLTYPE *get_RightMargin)(
        I_PageSetup* This,
        double *value);

    HRESULT (STDMETHODCALLTYPE *put_RightMargin)(
        I_PageSetup* This,
        double value);

    HRESULT (STDMETHODCALLTYPE *get_TopMargin)(
        I_PageSetup* This,
        double *value);

    HRESULT (STDMETHODCALLTYPE *put_TopMargin)(
        I_PageSetup* This,
        double value);

    HRESULT (STDMETHODCALLTYPE *get_Zoom)(
        I_PageSetup* This,
        VARIANT *value);

    HRESULT (STDMETHODCALLTYPE *put_Zoom)(
        I_PageSetup* This,
        VARIANT value);

    HRESULT (STDMETHODCALLTYPE *get_PrintComments)(
        I_PageSetup* This,
        XlPrintLocation *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintComments)(
        I_PageSetup* This,
        XlPrintLocation value);

    HRESULT (STDMETHODCALLTYPE *get_PrintErrors)(
        I_PageSetup* This,
        XlPrintErrors *value);

    HRESULT (STDMETHODCALLTYPE *put_PrintErrors)(
        I_PageSetup* This,
        XlPrintErrors value);

    HRESULT (STDMETHODCALLTYPE *get_CenterHeaderPicture)(
        I_PageSetup* This,
        IDispatch **value);

    HRESULT (STDMETHODCALLTYPE *get_CenterFooterPicture)(
        I_PageSetup* This,
        IDispatch **value);

    HRESULT (STDMETHODCALLTYPE *get_LeftHeaderPicture)(
        I_PageSetup* This,
        IDispatch **value);

    HRESULT (STDMETHODCALLTYPE *get_LeftFooterPicture)(
        I_PageSetup* This,
        IDispatch **value);

    HRESULT (STDMETHODCALLTYPE *get_RightHeaderPicture)(
        I_PageSetup* This,
        IDispatch **value);

    HRESULT (STDMETHODCALLTYPE *get_RightFooterPicture)(
        I_PageSetup* This,
        IDispatch **value);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE(" (12) ERROR when VariantChangeType \n");
                return hres;
            }
            vbval = V_BOOL(&vtmp);
            hres = MSO_TO_OO_I_PageSetup_put_CenterVertically(iface, vbval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_CenterVertically(iface, &vbval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = vbval;
            }
            return hres;
        }
    case dispid_pagesetup_printtitlerows://PrintTitleRows
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (13) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BSTR);
            if (FAILED(hres)) {
                TRACE(" (12) ERROR when VariantChangeType \n");
                return hres;
            }
            hres = MSO_TO_OO_I_PageSetup_put_PrintTitleRows(iface, V_BSTR(&vtmp));
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_PageSetup_get_PrintTitleRows(iface, &bstr_val);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BSTR;
                V_BSTR(pVarResult) = bstr_val;
            }
            return hres;
        }
    }
    TRACE(" dispIdMember = %i NOT REALIZE\n",dispIdMember);
    return E_NOTIMPL;
}


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

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

    pagesetup = HeapAlloc(GetProcessHeap(), 0, sizeof(*pagesetup));
    if (!pagesetup)
    {
        return E_OUTOFMEMORY;
    }

    pagesetup->_pagesetupVtbl = &MSO_TO_OO_I_PageSetupVtbl;
    pagesetup->ref = 0;
    pagesetup->pwsheet = NULL;
    pagesetup->pApplication = NULL;

    *ppObj = &pagesetup->_pagesetupVtbl;

    return S_OK;
}
