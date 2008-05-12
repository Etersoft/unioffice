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

static WCHAR const str_leftmargin[] = {
    'L','e','f','t','M','a','r','g','i','n',0};
static WCHAR const str_rightmargin[] = {
    'R','i','g','h','t','M','a','r','g','i','n',0};
static WCHAR const str_topmargin[] = {
    'T','o','p','M','a','r','g','i','n',0};
static WCHAR const str_bottommargin[] = {
    'B','o','t','t','o','m','M','a','r','g','i','n',0};
static WCHAR const str_orientation[] = {
    'O','r','i','e','n','t','a','t','i','o','n',0};
static WCHAR const str_zoom[] = {
    'Z','o','o','m',0};
static WCHAR const str_fittopagestall[] = {
    'F','i','t','T','o','P','a','g','e','s','T','a','l','l',0};
static WCHAR const str_fittopageswide[] = {
    'F','i','t','T','o','P','a','g','e','s','W','i','d','e',0};
static WCHAR const str_headermargin[] = {
    'H','e','a','d','e','r','M','a','r','g','i','n',0};
static WCHAR const str_footermargin[] = {
    'F','o','o','t','e','r','M','a','r','g','i','n',0};
static WCHAR const str_centerhorizontally[] = {
    'C','e','n','t','e','r','H','o','r','i','z','o','n','t','a','l','l','y',0};
static WCHAR const str_centervertically[] = {
    'C','e','n','t','e','r','V','e','r','t','i','c','a','l','l','y',0};

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_PageSetup_AddRef(
        I_PageSetup* iface)
{
    PageSetupImpl *This = (PageSetupImpl*)iface;
    ULONG ref;

    TRACE("mso_to_oo.dll:i_pagesetup.c:AddRef REF = %i \n", This->ref);

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

    TRACE("mso_to_oo.dll:i_pagesetup.c:QueryInterface \n");

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

    TRACE("mso_to_oo.dll:i_pagesetup.c:Release REF = %i \n", This->ref);

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
    TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"LeftMargin",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (GET) ERROR when  LeftMargin \n");
        return hres;
    }

    *value = V_I4(&vres);

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
    TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"LeftMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:LeftMargin (PUT) ERROR when  LeftMargin \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"RightMargin",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (GET) ERROR when  RightMargin \n");
        return hres;
    }

    *value = V_I4(&vres);

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
    TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"RightMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:RightMargin (PUT) ERROR when  RightMargin \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"TopMargin",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (GET) ERROR when  TopMargin \n");
        return hres;
    }

    *value = V_I4(&vres);

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
    TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"TopMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:TopMargin (PUT) ERROR when  TopMargin \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"BottomMargin",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (GET) ERROR when  BottomMargin \n");
        return hres;
    }

    *value = V_I4(&vres);

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
    TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"BottomMargin",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:BottomMargin (PUT) ERROR when  BottomMargin \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"IsLandscape",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (GET) ERROR when  Orientation \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (PUT) ERROR when getByName2 \n");
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
        TRACE("mso_to_oo.dll:i_pagesetup.c:Orientation (PUT) ERROR when  IsLandscape \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"PageScale",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (GET) ERROR when  PageScale \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (PUT) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"PageScale",1, value);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:Zoom (PUT) ERROR when  PageScale \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesY",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (GET) ERROR when  ScaleToPagesY \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT) ERROR when getByName2 \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&value, &value, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT) ERROR when VariantChangeType \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesY",1, value);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesTall (PUT) ERROR when  ScaleToPagesY \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesX",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (GET) ERROR when  ScaleToPagesX \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (PUT) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"ScaleToPagesX",1, value);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FitToPagesWide (PUT) ERROR when  ScaleToPagesX \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"HeaderHeight",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (GET) ERROR when  HeaderHeight \n");
        return hres;
    }

    *value = V_I4(&vres);

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
    TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"HeaderHeight",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:HeaderMargin (PUT) ERROR when  HeaderHeight \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"FooterHeight",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (GET) ERROR when  FooterHeight \n");
        return hres;
    }

    *value = V_I4(&vres);

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
    TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"FooterHeight",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:FooterMargin (PUT) ERROR when  FooterHeight \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"CenterHorizontally",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (GET) ERROR when  CenterHorizontally \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_BOOL;
    V_BOOL(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"CenterHorizontally",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterHorizontally (PUT) ERROR when CenterHorizontally \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (GET)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (GET) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (GET) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (GET) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (GET) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (GET) ERROR when getByName2 \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vres, V_DISPATCH(&vstyle), L"CenterVertically",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (GET) ERROR when  CenterVertically \n");
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
    TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (PUT)\n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (PUT) ERROR Object is NULL");
        return E_FAIL;
    }
    /*С начала необходимо узнать название стиля используемого на странице*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &name_of_style, wsh->pOOSheet, L"PageStyle",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (PUT) ERROR when PageStyle \n");
        return hres;
    }
    WTRACE(L"name of Style - %s \n", V_BSTR(&name_of_style));

    /*Теперь получим этот стиль из списка всех стилей*/
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vstyles, wb->pDoc, L"StyleFamilies",0);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (PUT) ERROR when StyleFamilies \n");
        return hres;
    }
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"PageStyles");
    hres = AutoWrap(DISPATCH_METHOD, &vpagestyles, V_DISPATCH(&vstyles), L"getByName",1,param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (PUT) ERROR when getByName \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &vstyle, V_DISPATCH(&vpagestyles), L"getByName",1, name_of_style);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (PUT) ERROR when getByName2 \n");
        return hres;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_BOOL;
    V_BOOL(&param1) = value;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vres, V_DISPATCH(&vstyle), L"CenterVertically",1, param1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_pagesetup.c:CenterVertically (PUT) ERROR when CenterVertically \n");
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


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_PageSetup_GetTypeInfoCount(
        I_PageSetup* iface,
        UINT *pctinfo)
{
    TRACE("mso_to_oo.dll:i_pagesetup.c:GetTypeInfoCount \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_PageSetup_GetTypeInfo(
        I_PageSetup* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("mso_to_oo.dll:i_pagesetup.c:GetTypeInfo\n");
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
    if (!lstrcmpiW(*rgszNames, str_leftmargin)) {
        *rgDispId = 1;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_rightmargin)) {
        *rgDispId = 2;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_topmargin)) {
        *rgDispId = 3;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_bottommargin)) {
        *rgDispId = 4;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_orientation)) {
        *rgDispId = 5;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_zoom)) {
        *rgDispId = 6;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_fittopagestall)) {
        *rgDispId = 7;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_fittopageswide)) {
        *rgDispId = 8;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_headermargin)) {
        *rgDispId = 9;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_footermargin)) {
        *rgDispId = 10;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_centerhorizontally)) {
        *rgDispId = 11;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_centervertically)) {
        *rgDispId = 12;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L"mso_to_oo.dll:i_pagesetup.c:PageSetup - %s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
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

    VariantInit(&vval);
    VariantInit(&vtmp);

    TRACE("mso_to_oo.dll:i_range:Invoke \n");

    switch(dispIdMember)
    {
    case 1://leftmargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (1) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (1) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
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
    case 2://rightmargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (2) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (2) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
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
    case 3://topmargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (3) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (3) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
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
    case 4://bottommargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (4) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (4) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
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
    case 5://orientation
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (5) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (5) ERROR when VariantChangeType \n");
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
    case 6://zoom
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (6) ERROR Number of parameters \n");
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
    case 7://FitToPagesTall
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (7) ERROR Number of parameters \n");
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
    case 8://FitToPagesWide
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (8) ERROR Number of parameters \n");
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
    case 9://HeaderMargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (9) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (9) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
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
    case 10://FooterMargin
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (10) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_R8);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (10) ERROR when VariantChangeType \n");
                return hres;
            }
            dval = V_R8(&vtmp);
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
    case 11://CenterHorizontally
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (11) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (11) ERROR when VariantChangeType \n");
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
    case 12://CenterVertically
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("mso_to_oo.dll:i_range:Invoke (12) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE("mso_to_oo.dll:i_range:Invoke (12) ERROR when VariantChangeType \n");
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
    }
    TRACE("mso_to_oo.dll:i_pagesetup.c:Invoke dispIdMember = %i NOT REALIZE\n",dispIdMember);
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
    MSO_TO_OO_I_PageSetup_get_LeftMargin,
    MSO_TO_OO_I_PageSetup_put_LeftMargin,
    MSO_TO_OO_I_PageSetup_get_RightMargin,
    MSO_TO_OO_I_PageSetup_put_RightMargin,
    MSO_TO_OO_I_PageSetup_get_TopMargin,
    MSO_TO_OO_I_PageSetup_put_TopMargin,
    MSO_TO_OO_I_PageSetup_get_BottomMargin,
    MSO_TO_OO_I_PageSetup_put_BottomMargin,
    MSO_TO_OO_I_PageSetup_get_Orientation,
    MSO_TO_OO_I_PageSetup_put_Orientation,
    MSO_TO_OO_I_PageSetup_get_Zoom,
    MSO_TO_OO_I_PageSetup_put_Zoom,
    MSO_TO_OO_I_PageSetup_get_FitToPagesTall,
    MSO_TO_OO_I_PageSetup_put_FitToPagesTall,
    MSO_TO_OO_I_PageSetup_get_FitToPagesWide,
    MSO_TO_OO_I_PageSetup_put_FitToPagesWide,
    MSO_TO_OO_I_PageSetup_get_HeaderMargin,
    MSO_TO_OO_I_PageSetup_put_HeaderMargin,
    MSO_TO_OO_I_PageSetup_get_FooterMargin,
    MSO_TO_OO_I_PageSetup_put_FooterMargin,
    MSO_TO_OO_I_PageSetup_get_CenterHorizontally,
    MSO_TO_OO_I_PageSetup_put_CenterHorizontally,
    MSO_TO_OO_I_PageSetup_get_CenterVertically,
    MSO_TO_OO_I_PageSetup_put_CenterVertically
};

PageSetupImpl MSO_TO_OO_PageSetup =
{
    &MSO_TO_OO_I_PageSetupVtbl,
    0,
    NULL,
    NULL
};

extern HRESULT _I_PageSetupConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    PageSetupImpl *pagesetup;

    TRACE("mso_to_oo.dll:i_pagesetup:Constructor (%p,%p)\n", pUnkOuter, ppObj);

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
