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
static WCHAR const str_protect[] = {
    'P','r','o','t','e','c','t',0};
static WCHAR const str_unprotect[] = {
    'U','n','p','r','o','t','e','c','t',0};
static WCHAR const str_name[] = {
    'N','a','m','e',0};

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Workbook_AddRef(
        I_Workbook* iface)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

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

    TRACE("\n");

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

    TRACE("REF = %i \n", This->ref);
    
    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pApplication != NULL) {
            I_ApplicationExcel_Release(This->pApplication);
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

    TRACE("\n");

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

    TRACE("\n");

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
    _ApplicationExcelImpl *app = (_ApplicationExcelImpl*)This->pApplication;
    WorkbooksImpl *wbs = (WorkbooksImpl*)app->pdWorkbooks;
    int i;
    IDispatch *pdtmp;
/*TODO*/
/*Игнорируем все параметры*/
    TRACE("\n");
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

    TRACE(" \n");

    if (This==NULL) {
        TRACE("ERROR objetct is NULL \n");
        return E_FAIL;
    }
    if (V_VT(&Filename)!=VT_BSTR) {
        TRACE("ERROR no filename \n");
        return E_FAIL;
    }

    /* Create PropertyValue with save-format-data */
    IDispatch *dpv;
    MSO_TO_OO_GetDispatchPropertyValue((I_ApplicationExcel*)(This->pApplication), &dpv);
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

    WTRACE(L"FILENAME = %s \n", V_BSTR(&Filename));
    TRACE("\n");
    WTRACE(L"SaveAs FILENAMEURL = %s  \n", FilenameURL);
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
        TRACE("ERROR when StoreToURL \n");
        return hres;
    }
    VariantClear(&res);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Save(
        I_Workbook* iface)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Protect(
        I_Workbook* iface,
        VARIANT Password,
        VARIANT Structure,
        VARIANT Windows)
{
    /*TODO Think about other parameters*/
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT param, res;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    VariantInit(&param);
    if ((V_VT(&Password)==VT_EMPTY)||(V_VT(&Password)==VT_NULL)) {
        V_VT(&param) = VT_BSTR;
        V_BSTR(&param) = SysAllocString(L"");
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"protect", 1, param);
        if (FAILED(hres)) {
            TRACE("ERROR when protect\n");
            return hres;
        }
    } else {
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"protect", 1, Password);
        if (FAILED(hres)) {
            TRACE("ERROR when protect\n");
            return hres;
        }
    }

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Unprotect(
        I_Workbook* iface,
        VARIANT Password)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT param, res;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    VariantInit(&param);
    if ((V_VT(&Password)==VT_EMPTY)||(V_VT(&Password)==VT_NULL)) {
        V_VT(&param) = VT_BSTR;
        V_BSTR(&param) = SysAllocString(L"");
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"unprotect", 1, param);
        if (FAILED(hres)) {
            TRACE("ERROR when unprotect\n");
            return hres;
        }
    } else {
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"unprotect", 1, Password);
        if (FAILED(hres)) {
            TRACE("ERROR when unprotect\n");
            return hres;
        }
    }

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Name(
        I_Workbook* iface,
        BSTR *retval)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT vRes, ooframe, oocontr;
    HRESULT hres;

    TRACE("\n");

    VariantInit(&vRes);
    VariantInit(&oocontr);
    VariantInit(&ooframe);

    hres = AutoWrap(DISPATCH_METHOD, &oocontr,This->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &ooframe, V_DISPATCH(&oocontr), L"getFrame",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getFrame \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, V_DISPATCH(&ooframe), L"Title",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Title \n");
        return hres;
    }

    *retval = SysAllocString(V_BSTR(&vRes));

    VariantClear(&vRes);
    VariantClear(&oocontr);
    VariantClear(&ooframe);

    return S_OK;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetTypeInfoCount(
        I_Workbook* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}


static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetTypeInfo(
        I_Workbook* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
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
    if (!lstrcmpiW(*rgszNames, str_protect)) {
        *rgDispId = 6;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_unprotect)) {
        *rgDispId = 7;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_name)) {
        *rgDispId = 8;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L"%s NOT REALIZE\n",*rgszNames);
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
    VARIANT vmas[12];
    int i;
    VARIANT vtmp;
    BSTR bret;

    for (i=0;i<12;i++) {
         VariantInit(&vmas[i]);
         V_VT(&vmas[i])=VT_EMPTY;
    }

    TRACE("\n");

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
        if (pDispParams->cArgs>3) {
            TRACE(" (3) ERROR Parameters");
            return E_FAIL;
        }
        /*необходимо перевернуть параметры*/
        for (i=0;i<pDispParams->cArgs;i++) {
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-i-1], &vmas[i]))) return E_FAIL;
        }
        hr = MSO_TO_OO_I_Workbook_Close(iface, vmas[0], vmas[1], vmas[2]);
        if (FAILED(hr)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hr;
        }
        return hr;
    case 4:
        if (pDispParams->cArgs>12) {
            TRACE(" (4) ERROR Parameters");
            return E_FAIL;
        }
        /*необходимо перевернуть параметры*/
        for (i=0;i<pDispParams->cArgs;i++) {
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-i-1], &vmas[i]))) return E_FAIL;
        }
        VariantChangeTypeEx(&vmas[6], &vmas[6], 0, 0, VT_I4);
        hr = MSO_TO_OO_I_Workbook_SaveAs(iface, vmas[0], vmas[1], vmas[2], vmas[3], vmas[4], vmas[5], V_I4(&vmas[6]), vmas[7], vmas[8], vmas[9], vmas[10], vmas[11]);
        if (FAILED(hr)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hr;
        }
        return hr;
    case 5:
        hr = MSO_TO_OO_I_Workbook_Save(iface);
        return hr;
    case 6://Protect
        for (i=0;i<12;i++) {
            VariantInit(&vmas[i]);
            V_VT(&vmas[i])=VT_EMPTY;
        }
        /*необходимо перевернуть параметры*/
        for (i=0;i<pDispParams->cArgs;i++) {
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-i-1], &vmas[i]))) return E_FAIL;
        }
        return MSO_TO_OO_I_Workbook_Protect(iface, vmas[0], vmas[1], vmas[2]);
    case 7://UnProtect
        switch (pDispParams->cArgs) {
        case 0:
            VariantClear(&vtmp);
            V_VT(&vtmp) = VT_EMPTY;
            break;
        case 1:
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp))) return E_FAIL;
            break;
        default:
            TRACE("ERROR parameters \n");
            return E_INVALIDARG;
        }
        return MSO_TO_OO_I_Workbook_Unprotect(iface,vtmp);
    case 8:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_Workbook_get_Name(iface, &bret);
            if (FAILED(hr)) {
                TRACE("Error get Name \n");
                return hr;
            }
            if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_BSTR;
                    V_BSTR(pVarResult)=bret;
            }
            return S_OK;
        }
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
    MSO_TO_OO_I_Workbook_Save,
    MSO_TO_OO_I_Workbook_Protect,
    MSO_TO_OO_I_Workbook_Unprotect,
    MSO_TO_OO_I_Workbook_get_Name
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

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

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
