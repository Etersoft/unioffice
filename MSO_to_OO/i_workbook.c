/*
 * IWorkbook interface functions
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

static WCHAR const str_sheets[] = {
    'S','h','e','e','t','s',0};
static WCHAR const str_worksheets[] = {
    'W','o','r','k','S','h','e','e','t','s',0};
static WCHAR const str_close[] = {
    'C','l','o','s','e',0};
static WCHAR const str_saveas[] = {
    'S','a','v','e','A','s',0};
static WCHAR const str_save[] = {
    'S','a','v','e',0};

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Workbook_AddRef(
        I_Workbook* iface)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    ULONG ref;

    TRACE("mso_to_oo.dll:i_workbook.c:AddRef REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}


static HRESULT WINAPI MSO_TO_OO_I_Workbook_QueryInterface(
        I_Workbook* iface,
        REFIID riid,
        void **ppvObject)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;

    TRACE("mso_to_oo.dll:i_workbook.c:QueryInterface \n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Workbook)) {
        *ppvObject = &This->_workbookVtbl;
        MSO_TO_OO_I_Workbook_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}


static ULONG WINAPI MSO_TO_OO_I_Workbook_Release(
        I_Workbook* iface)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    ULONG ref;

    TRACE("mso_to_oo.dll:i_workbook.c:Release REF = %i \n", This->ref);
    
    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pApplication != NULL) {
            I_ApplicationExcell_Release(This->pApplication);
            This->pApplication = NULL;
        }
        if (This->pDoc != NULL) {
            IDispatch_Release(This->pDoc);
            This->pDoc = NULL;
        }
        if (This->pSheets != NULL) {
            IDispatch_Release(This->pSheets);
            This->pSheets = NULL;
        }
        /*SysFreeString(This->filename);*/
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Workbook methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Sheets(
        I_Workbook* iface,
        IDispatch **ppSheets)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;

    TRACE("mso_to_oo.dll:i_workbook.c:Sheets (GET) \n");

    if (This->pSheets == NULL) 
        return E_FAIL;

    *ppSheets = This->pSheets;
    IDispatch_AddRef(This->pSheets);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_WorkSheets(
        I_Workbook* iface,
        IDispatch **ppSheets)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;

    TRACE("msi_to_oo.dll:i_workbook.c:WorkSheets (GET) \n");

    if (This->pSheets == NULL) 
        return E_FAIL;

    *ppSheets = This->pSheets;
    IDispatch_AddRef(This->pSheets);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Close(
        I_Workbook* iface,
        VARIANT SaveChanges,
        VARIANT Filename,
        VARIANT RouteWorkbook)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    BSTR filename;
    HRESULT hres;
    _ApplicationExcellImpl *app = (_ApplicationExcellImpl*)This->pApplication;
    WorkbooksImpl *wbs = (WorkbooksImpl*)app->pdWorkbooks;
    int i;
    IDispatch *pdtmp;
/*TODO*/
/*Игнорируем все параметры*/
    TRACE("mso_to_oo.dll:i_workbook.c:Close \n");
    filename = SysAllocString(L"");
    hres = MSO_TO_OO_CloseWorkbook(iface, filename);
    SysFreeString(filename);
    if (FAILED(hres)) {
        TRACE("\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Workbook_Release(iface);
    iface = NULL;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SaveAs(
        I_Workbook* iface,
        VARIANT Filename,
        VARIANT FileFormat,
        VARIANT Password,
        VARIANT WriteResPassword,
        VARIANT ReadOnlyRecommended,
        VARIANT CreateBackup,
        XlSaveAsAccessMode AccessMode,
        VARIANT ConflictResolution,
        VARIANT AddToMru,
        VARIANT TextCodepage,
        VARIANT TextVisualLayout,
        VARIANT Local)
{
/*Пока игнорируем все параметры кроме первого*/
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT res, p3, p1;
    HRESULT hres;
    long ix = 0;
    SAFEARRAY FAR* pPropVals;
    BSTR FilenameURL;

    TRACE("mso_to_oo.dll:i_workbook.c:SaveAs \n");

    if (This==NULL) {
        TRACE("mso_to_oo.dll:i_workbook.c:SaveAs ERROR objetct is NULL \n");
        return E_FAIL;
    }
    if (V_VT(&Filename)!=VT_BSTR) {
        TRACE("mso_to_oo.dll:i_workbook.c:SaveAs ERROR no filename \n");
        return E_FAIL;
    }

    /* Create PropertyValue with save-format-data */
    IDispatch *dpv;
    MSO_TO_OO_GetDispatchPropertyValue((I_ApplicationExcell*)(This->pApplication), &dpv);
    if (dpv == NULL)
        return E_FAIL;

    /* Set PropertyValue data */
    V_VT(&p1) = VT_BSTR;
    V_BSTR(&p1) = SysAllocString(L"FilterName");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Name", 1, p1);
    SysFreeString(V_BSTR(&p1));
    V_BSTR(&p1) = SysAllocString(L"MS Excel 97");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Value", 1, p1);
    SysFreeString(V_BSTR(&p1));
    /* Init params */
    pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 1 );

    hres = SafeArrayPutElement( pPropVals, &ix, dpv );

    VariantInit (&p3);
    V_VT(&p3) = VT_DISPATCH | VT_ARRAY;
    V_ARRAY(&p3) = pPropVals;
    MSO_TO_OO_MakeURLFromFilename(V_BSTR(&Filename), &FilenameURL);
    V_BSTR(&p1) = SysAllocString(FilenameURL);

    WTRACE(L"mso_to_oo.dll:i_workbook.c:SaveAs FILENAME = %s \n", V_BSTR(&Filename));
    TRACE("\n");
    WTRACE(L"mso_to_oo.dll:i_workbook.c:SaveAs FILENAMEURL = %s  \n", FilenameURL);
    TRACE("\n");
    int i=0;
    while (*(FilenameURL+i)!=0) {
        WTRACE(L"%c",*(FilenameURL+i));
        i++;
    }
    TRACE("\n");
    /* Call StoreToURL for save document to file */
    hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"StoreToURL", 2, p3, p1);
    if (FAILED(hres)) {
        TRACE("mso_to_oo.dll:i_workbook.c:SaveAs ERROR when StoreToURL \n");
        return hres;
    }
    VariantClear(&res);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Save(
        I_Workbook* iface)
{
    TRACE("mso_to_oo.dll:i_workbook.c:Save \n");
    return E_NOTIMPL;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetTypeInfoCount(
        I_Workbook* iface,
        UINT *pctinfo)
{
    TRACE("mso_to_oo.dll:i_workbook.c:GetTypeInfoCount \n");
    return E_NOTIMPL;
}


static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetTypeInfo(
        I_Workbook* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("mso_to_oo.dll:i_workbook.c:GetTypeInfo \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetIDsOfNames(
        I_Workbook* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    if (!lstrcmpiW(*rgszNames, str_sheets)) {
        *rgDispId = 1;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_worksheets)) {
        *rgDispId = 2;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_close)) {
        *rgDispId = 3;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_saveas)) {
        *rgDispId = 4;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_save)) {
        *rgDispId = 5;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L"mso_to_oo.dll:i_workbook.c:Workbook - %s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Invoke(
        I_Workbook* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    HRESULT hr;
    IDispatch *sheets;
    IDispatch *sheet;
    VARIANT vNull,par1,par2,par3,par4,par5,par6,par7,par8,par9,par10,par11,par12;

    VariantInit(&vNull);
    VariantInit(&par1);
    VariantInit(&par2);
    VariantInit(&par3);
    VariantInit(&par4);
    VariantInit(&par5);
    VariantInit(&par6);
    VariantInit(&par7);
    VariantInit(&par8);
    VariantInit(&par9);
    VariantInit(&par10);
    VariantInit(&par11);
    VariantInit(&par12);

    TRACE("mso_to_oo.dll:i_workbook.c:Invoke \n");

    if (This == NULL) return E_POINTER;

    switch (dispIdMember) 
    {
    case 1:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_Workbook_get_Sheets(iface,&sheets);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pDispParams->cArgs==1) {
                hr = I_Sheets_get__Default((I_Sheets*)sheets, pDispParams->rgvarg[0], &sheet);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    I_Sheets_Release((I_Sheets*)sheets);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)sheet;
                    I_Sheets_Release((I_Sheets*)sheets);
                }
            } else {
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)sheets;
                }
            }
            return hr;
        }
    case 2:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_Workbook_get_WorkSheets(iface,&sheets);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pDispParams->cArgs==1) {
                hr = I_Sheets_get__Default((I_Sheets*)sheets, pDispParams->rgvarg[0], &sheet);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    I_Sheets_Release((I_Sheets*)sheets);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)sheet;
                    I_Sheets_Release((I_Sheets*)sheets);
                }
            } else {
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)sheets;
                }
            }
            return hr;
        }
    case 3:
        switch(pDispParams->cArgs) {
        case 0:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (3) 0 parameter \n");
            hr = MSO_TO_OO_I_Workbook_Close(iface, vNull, vNull, vNull);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            return hr;
        case 1:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (3) 1 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par1))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_Close(iface, par1, vNull, vNull);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            return hr;
        case 2:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (3) 2 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par2))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_Close(iface, par1, par2, vNull);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            return hr;
        case 3:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (3) 3 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par3))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_Close(iface, par1, par2, par3);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            return hr;
        default:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (3) ERROR Parameters");
            return E_FAIL;
        }
    case 4:
        switch (pDispParams->cArgs) {
        case 0:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 0 parameter \n");
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, vNull, vNull, vNull, vNull, vNull, vNull, 0, vNull, vNull, vNull, vNull, vNull);
            break;
        case 1:
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par1))) return E_FAIL;
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 1 parameter \n");
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, vNull, vNull, vNull, vNull, vNull, 0, vNull, vNull, vNull, vNull, vNull);
            break;
        case 2:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 2 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par2))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, vNull, vNull, vNull, vNull, 0, vNull, vNull, vNull, vNull, vNull);
            break;
        case 3:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 3 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par3))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, vNull, vNull, vNull, 0, vNull, vNull, vNull, vNull, vNull);
            break;
        case 4:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 4 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par4))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, vNull, vNull, 0, vNull, vNull, vNull, vNull, vNull);
            break;
        case 5:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 5 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par5))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, vNull, 0, vNull, vNull, vNull, vNull, vNull);
            break;
        case 6:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 6 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[5], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par5))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par6))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, par6, 0, vNull, vNull, vNull, vNull, vNull);
            break;
        case 7:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 7 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[6], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[5], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par5))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par6))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par7))) return E_FAIL;
            VariantChangeTypeEx(&par7, &par7, 0, 0, VT_I4);
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, par6, V_I4(&par7), vNull, vNull, vNull, vNull, vNull);
            break;
        case 8:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 8 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[7], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[6], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[5], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par5))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par6))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par7))) return E_FAIL;
            VariantChangeTypeEx(&par7, &par7, 0, 0, VT_I4);
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par8))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, par6, V_I4(&par7), par8, vNull, vNull, vNull, vNull);
            break;
        case 9:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 9 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[8], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[7], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[6], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[5], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par5))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par6))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par7))) return E_FAIL;
            VariantChangeTypeEx(&par7, &par7, 0, 0, VT_I4);
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par8))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par9))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, par6, V_I4(&par7), par8, par9, vNull, vNull, vNull);
            break;
        case 10:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 10 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[9], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[8], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[7], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[6], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[5], &par5))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par6))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par7))) return E_FAIL;
            VariantChangeTypeEx(&par7, &par7, 0, 0, VT_I4);
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par8))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par9))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par10))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, par6, V_I4(&par7), par8, par9, par10, vNull, vNull);
            break;
        case 11:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 11 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[10], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[9], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[8], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[7], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[6], &par5))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[5], &par6))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par7))) return E_FAIL;
            VariantChangeTypeEx(&par7, &par7, 0, 0, VT_I4);
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par8))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par9))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par10))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par11))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, par6, V_I4(&par7), par8, par9, par10, par11, vNull);
            break;
        case 12:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) 12 parameter \n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[11], &par1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[10], &par2))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[9], &par3))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[8], &par4))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[7], &par5))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[6], &par6))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[5], &par7))) return E_FAIL;
            VariantChangeTypeEx(&par7, &par7, 0, 0, VT_I4);
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[4], &par8))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[3], &par9))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &par10))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &par11))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &par12))) return E_FAIL;
            hr = MSO_TO_OO_I_Workbook_SaveAs(iface, par1, par2, par3, par4, par5, par6, V_I4(&par7), par8, par9, par10, par11, par12);
            break;
        default:
            TRACE("mso_to_oo.dll:i_workbook.c:Invoke (4) ERROR Parameters");
            hr = E_FAIL;
            break;
        }
        return hr;
    case 5:
        hr = MSO_TO_OO_I_Workbook_Save(iface);
        return hr;
    }

    return E_NOTIMPL;
}


const I_WorkbookVtbl MSO_TO_OO_I_WorkbookVtbl =
{
    MSO_TO_OO_I_Workbook_QueryInterface,
    MSO_TO_OO_I_Workbook_AddRef,
    MSO_TO_OO_I_Workbook_Release,
    MSO_TO_OO_I_Workbook_GetTypeInfoCount,
    MSO_TO_OO_I_Workbook_GetTypeInfo,
    MSO_TO_OO_I_Workbook_GetIDsOfNames,
    MSO_TO_OO_I_Workbook_Invoke,
    MSO_TO_OO_I_Workbook_get_Sheets,
    MSO_TO_OO_I_Workbook_get_WorkSheets,
    MSO_TO_OO_I_Workbook_Close,
    MSO_TO_OO_I_Workbook_SaveAs,
    MSO_TO_OO_I_Workbook_Save
};

WorkbookImpl MSO_TO_OO_Workbook =
{
    &MSO_TO_OO_I_WorkbookVtbl,
    0,
    NULL,
    NULL,
    NULL
};

extern HRESULT _I_WorkbookConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    WorkbookImpl *workbook;

    TRACE("mso_to_oo.dll:i_workbook.c:Constructor (%p,%p)\n", pUnkOuter, ppObj);

    workbook = HeapAlloc(GetProcessHeap(), 0, sizeof(*workbook));
    if (!workbook)
    {
        return E_OUTOFMEMORY;
    }

    workbook->_workbookVtbl = &MSO_TO_OO_I_WorkbookVtbl;
    workbook->ref = 0;
    workbook->pApplication = NULL;
    workbook->pDoc = NULL;
    workbook->pSheets = NULL;

    *ppObj = &workbook->_workbookVtbl;

    return S_OK;
}
