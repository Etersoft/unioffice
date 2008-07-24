/*
 * IWorkbooks interface functions
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
#include <oleauto.h>

static WCHAR const str_add[] = {
    'A','d','d',0};
static WCHAR const str__open[] = {
    '_','o','p','e','n',0};
static WCHAR const str_close[] = {
    'C','l','o','s','e',0};
static WCHAR const str_count[] = {
    'C','o','u','n','t',0};
static WCHAR const str_application[] = {
    'A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR const str_parent[] = {
    'P','a','r','e','n','t',0};
static WCHAR const str_open[] = {
    'O','p','e','n',0};
static WCHAR const str_opentext[] = {
    'O','p','e','n','T','e','x','t',0};
static WCHAR const str__opentext[] = {
    '_','O','p','e','n','T','e','x','t',0};
static WCHAR const str_openxml[] = {
    'O','p','e','n','X','M','L',0};
static WCHAR const str__openxml[] = {
    '_','O','p','e','n','X','M','L',0};
static WCHAR const str_opendatabase[] = {
    'O','p','e','n','D','a','t','a','b','a','s','e',0};
static WCHAR const str_cancheckout[] = {
    'C','a','n','C','h','e','c','k','O','u','t',0};
static WCHAR const str_checkout[] = {
    'C','h','e','c','k','O','u','t',0};
static WCHAR const str_creator[] = {
    'C','r','e','a','t','o','r',0};
static WCHAR const str__default[] = {
    '_','D','e','f','a','u','l','t',0};
static WCHAR const str_item[] = {
    'I','t','e','m',0};
static WCHAR const str___opentext[] = {
    '_','_','O','p','e','n','T','e','x','t',0};

static WCHAR const str_pusto[]= {0};

/*** IUnknown methods ***/

static ULONG WINAPI MSO_TO_OO_I_Workbooks_AddRef(
        I_Workbooks* iface)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;
    ULONG ref;

    TRACE("REF=%i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);

    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}


static HRESULT WINAPI MSO_TO_OO_I_Workbooks_QueryInterface(
        I_Workbooks* iface,
        REFIID riid,
        void **ppvObject)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Workbooks)) {
        *ppvObject = &This->_workbooksVtbl;
        MSO_TO_OO_I_Workbooks_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}


static ULONG WINAPI MSO_TO_OO_I_Workbooks_Release(
        I_Workbooks* iface)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;
    ULONG ref;
    int i;

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);

    TRACE("REF = %i \n", This->ref);

    if (ref == 0) {
        if (This->pApplication!=NULL) {
            IDispatch_Release(This->pApplication);
            This->pApplication==NULL;
        }
        for (i=0;i<This->count_workbooks;i++)
            if (This->pworkbook[i]!=NULL) IDispatch_Release(This->pworkbook[i]);
        if (This->count_workbooks>0) HeapFree(GetProcessHeap(),HEAP_ZERO_MEMORY,This->pworkbook);
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Workbooks methods ***/

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_Add(
        I_Workbooks* iface,
        VARIANT varTemplate,
        IDispatch **ppWorkbook)
{
/*TODO подумать как добавить поддержку шаблонов*/

    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    IUnknown *punk = NULL;
    HRESULT hres;

    TRACE(" \n");

    if (This == NULL) return E_POINTER;

    hres = _I_WorkbookConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    if (This->count_workbooks==0){
        This->count_workbooks += 1;
        This->pworkbook = HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, sizeof(WorkbookImpl*));
        if (FAILED(hres)) return E_OUTOFMEMORY;
        This->current_workbook = 0;
    } else {
        This->count_workbooks += 1;
        This->pworkbook = HeapReAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, This->pworkbook, This->count_workbooks * sizeof(WorkbookImpl*));
        if (FAILED(hres)) return E_OUTOFMEMORY;
        This->current_workbook = This->count_workbooks - 1;
    }

    hres = I_Workbook_QueryInterface(punk, &IID_I_Workbook, (void**) &(This->pworkbook[This->current_workbook]));
/*    I_Workbook_Release(punk);*/
    if (FAILED(hres)) {
        TRACE("ERROR when QueryInterface\n");
        return E_FAIL;
    }
    if (!ppWorkbook) {
        /*подумать над правильностью такого решения*/
        ppWorkbook = HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, sizeof(WorkbookImpl*));
        TRACE(" AllocMemory");
        *ppWorkbook = This->pworkbook[This->current_workbook];
    } else {
        *ppWorkbook = This->pworkbook[This->current_workbook];
    }

    if ((V_VT(&varTemplate)==VT_EMPTY)||
        (V_VT(&varTemplate)==VT_NULL) ||
        (lstrlenW(V_BSTR(&varTemplate))==0)) {
        hres = MSO_TO_OO_I_Workbook_Initialize( This->pworkbook[This->current_workbook], This->pApplication);
    } else {
       /* Необходимо преобразовать путь+имя в нужную форму
       от C:\test test.xls
       к file:///c:/test%20test.xls */
       BSTR Filename;
       MSO_TO_OO_MakeURLFromFilename(V_BSTR(&varTemplate),&Filename);
       /*преобразовали*/
       WTRACE(L"FILENAME ------>  %s \n",Filename);
       hres = MSO_TO_OO_I_Workbook_Initialize2( This->pworkbook[This->current_workbook], This->pApplication, Filename, VARIANT_TRUE);
       SysFreeString(Filename);
    }
    if (FAILED(hres)) {
        *ppWorkbook = NULL;
        TRACE("ERROR when Workbook_Initialize");
        return hres;
    }
    I_Workbook_AddRef(*ppWorkbook);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks__Open(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT UpdateLinks,
        VARIANT ReadOnly,
        VARIANT Format,
        VARIANT Password,
        VARIANT WriteResPassword,
        VARIANT IgnoreReadOnlyRecommended,
        VARIANT Origin,
        VARIANT Delimiter,
        VARIANT Editable,
        VARIANT Notify,
        VARIANT Converter,
        VARIANT AddToMru,
        long Lcid,
        IDispatch **ppWorkbook)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_Close(
        I_Workbooks* iface,
        long lcid)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;
    int i;
    BSTR filename;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    filename = SysAllocString(L"");

    for (i=0;i<This->count_workbooks;i++) {
/*        WorkbookImpl *wb = (WorkbookImpl*)(This->pworkbook[i]);
        filename = SysAllocString(wb->filename);*/
        if (This->pworkbook[i]!=NULL) {
            MSO_TO_OO_CloseWorkbook(This->pworkbook[i], filename);
            I_Workbook_Release(This->pworkbook[i]);
            This->pworkbook[i]=NULL;
        }
/*        SysFreeString(filename);*/
    }
    SysFreeString(filename);
    HeapFree(GetProcessHeap(),HEAP_ZERO_MEMORY,This->pworkbook);
    This->count_workbooks = 0;
    This->current_workbook = -1;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Count(
        I_Workbooks* iface,
        int *count)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    TRACE(" \n");

    if (This == NULL) return E_POINTER;

    *count = This->count_workbooks;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Application(
        I_Workbooks* iface,
        IDispatch **value)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *value = This->pApplication;
    I_ApplicationExcel_AddRef(This->pApplication);

    if (value==NULL)
        return E_POINTER;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Parent(
        I_Workbooks* iface,
        IDispatch **value)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *value = This->pApplication;
    I_ApplicationExcel_AddRef(This->pApplication);

    if (value==NULL)
        return E_POINTER;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_Open(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT UpdateLinks,
        VARIANT ReadOnly,
        VARIANT Format,
        VARIANT Password,
        VARIANT WriteResPassword,
        VARIANT IgnoreReadOnlyRecommended,
        VARIANT Origin,
        VARIANT Delimiter,
        VARIANT Editable,
        VARIANT Notify,
        VARIANT Converter,
        VARIANT AddToMru,
        VARIANT Local,
        VARIANT CorruptLoad,
        long Lcid,
        IDispatch **ppWorkbook)
{
/*TODO подумать как добавить поддержку остальных параметров
в данный момент используется только имя файла*/

    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    IUnknown *punk = NULL;
    HRESULT hres;

    TRACE(" \n");

    if (This == NULL) return E_POINTER;

    hres = _I_WorkbookConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    if (This->count_workbooks==0){
        This->count_workbooks += 1;
        This->pworkbook = HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, sizeof(WorkbookImpl*));
        if (FAILED(hres)) return E_OUTOFMEMORY;
        This->current_workbook = 0;
    } else {
        This->count_workbooks += 1;
        This->pworkbook = HeapReAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, This->pworkbook, This->count_workbooks * sizeof(WorkbookImpl*));
        if (FAILED(hres)) return E_OUTOFMEMORY;
        This->current_workbook = This->count_workbooks - 1;
    }

    hres = I_Workbook_QueryInterface(punk, &IID_I_Workbook, (void**) &(This->pworkbook[This->current_workbook]));
/*    I_Workbook_Release(punk);*/
    if (FAILED(hres)) return E_FAIL;
    *ppWorkbook = This->pworkbook[This->current_workbook];

    /* Необходимо преобразовать путь+имя в нужную форму
    от C:\test test.xls
    к file:///c:/test%20test.xls */
    BSTR filenametmp;
    MSO_TO_OO_MakeURLFromFilename(Filename, &filenametmp);
    /*преобразовали*/
    WTRACE(L"FILENAME ------>  %s \n",filenametmp);
    VARIANT_BOOL astemp;
    astemp = VARIANT_FALSE;

    hres = MSO_TO_OO_I_Workbook_Initialize2(This->pworkbook[This->current_workbook], This->pApplication, filenametmp,astemp);
    if (FAILED(hres)) {
        *ppWorkbook = NULL;
        return hres;
    }
    SysFreeString(filenametmp);
    I_Workbook_AddRef(*ppWorkbook);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_OpenText(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT Origin,
        VARIANT StartRow,
        VARIANT DataType,
        VARIANT TextQualifier,
        VARIANT ConsecutiveDelimiter,
        VARIANT Tab,
        VARIANT Semicolon,
        VARIANT Comma,
        VARIANT Space,
        VARIANT Other,
        VARIANT OtherChar,
        VARIANT FieldInfo,
        VARIANT TextVisualLayout,
        VARIANT DecimalSeparator,
        VARIANT ThousandsSeparator,
        VARIANT TrailingMinusNumbers,
        VARIANT Local,
        long Lcid,
        IDispatch **ppWorkbook)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks__OpenText(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT Origin,
        VARIANT StartRow,
        VARIANT DataType,
        VARIANT TextQualifier,
        VARIANT ConsecutiveDelimiter,
        VARIANT Tab,
        VARIANT Semicolon,
        VARIANT Comma,
        VARIANT Space,
        VARIANT Other,
        VARIANT OtherChar,
        VARIANT FieldInfo,
        VARIANT TextVisualLayout,
        VARIANT DecimalSeparator,
        VARIANT ThousandsSeparator,
        long Lcid,
        IDispatch **ppWorkbook)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_OpenXML(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT Stylesheets,
        VARIANT LoadOption,
        long Lcid,
        IDispatch **ppWorkbook)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks__OpenXML(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT Stylesheets,
        long Lcid,
        IDispatch **ppWorkbook)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_OpenDatabase(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT CommandText,
        VARIANT CommandType,
        VARIANT BackgroundQuery,
        VARIANT ImportDataAs,
        long Lcid,
        IDispatch **ppWorkboo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_CanCheckOut(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT_BOOL *result)
{
    TRACE("\n");
    /*всегда возвращаем - нельзя получить*/
    *result = VARIANT_FALSE;
    
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_put_CanCheckOut(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT_BOOL result)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_CheckOut(
        I_Workbooks* iface,
        BSTR Filename)
{
    TRACE("\n");
    /*Возвращаем S_OK на всякий случай*/
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Creator(
        I_Workbooks* iface,
        VARIANT *result)
{
    TRACE("\n");
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get__Default(
        I_Workbooks* iface,
        VARIANT varIndex,
        IDispatch **ppSheet)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Item(
        I_Workbooks* iface,
        VARIANT index,
        IDispatch **result)
{
    /*Должны обрабатывать и имена и числовые значения*/
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks___OpenText(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT Origin,
        VARIANT StartRow,
        VARIANT DataType,
        VARIANT TextQualifier,
        VARIANT ConsecutiveDelimiter,
        VARIANT Tab,
        VARIANT Semicolon,
        VARIANT Comma,
        VARIANT Space,
        VARIANT Other,
        VARIANT OtherChar,
        VARIANT FieldInfo,
        VARIANT TextVisualLayout,
        long Lcid,
        IDispatch **ppWorkbook)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get__NewEnum(
        I_Workbooks* iface,
        IUnknown **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbooks_GetTypeInfoCount(
        I_Workbooks* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_GetTypeInfo(
        I_Workbooks* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_GetIDsOfNames(
        I_Workbooks* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    if (!lstrcmpiW(*rgszNames, str_add)) {
        *rgDispId = dispid_workbooks_Add;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str__open)) {
        *rgDispId = dispid_workbooks__Open;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_close)) {
        *rgDispId = dispid_workbooks_Close;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_count)) {
        *rgDispId = dispid_workbooks_Count;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = dispid_workbooks_application;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_parent)) {
        *rgDispId = dispid_workbooks_parent;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_open)) {
        *rgDispId = dispid_workbooks_Open;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_opentext)) {
        *rgDispId = dispid_workbooks_OpenText;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str__opentext)) {
        *rgDispId = dispid_workbooks__OpenText;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_openxml)) {
        *rgDispId = dispid_workbooks_OpenXML;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str__openxml)) {
        *rgDispId = dispid_workbooks__OpenXML;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_opendatabase)) {
        *rgDispId = dispid_workbooks_OpenDatabase;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_cancheckout)) {
        *rgDispId = dispid_workbooks_CanCheckOut;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_checkout)) {
        *rgDispId = dispid_workbooks_CheckOut;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_creator)) {
        *rgDispId = dispid_workbooks_creator;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str__default)) {
        *rgDispId = dispid_workbooks__Default;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_item)) {
        *rgDispId = dispid_workbooks_Item;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str___opentext)) {
        *rgDispId = dispid_workbooks___OpenText;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L"%s NOT REALIZE\n", *rgszNames);
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_Invoke(
        I_Workbooks* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;
    HRESULT hr;
    int iresult;
    IDispatch* iapp;
    VARIANT_BOOL vbool;
    VARIANT vresult;
    VARIANT astemp;
    VARIANT vnull;
    long l;

    TRACE("\n");

    VariantInit(&vnull);
    VariantInit(&astemp);
    VariantInit(&vresult);

    if (This == NULL) return E_POINTER;

    switch (dispIdMember) {
        case dispid_workbooks_Add:
            if (pDispParams->cArgs==1) {
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &astemp);
                if (V_VT(&astemp)!=VT_BSTR) {
                    VariantClear(&astemp);
                    V_VT(&astemp) = VT_EMPTY;
                } else 
                    if (lstrlenW(V_BSTR(&astemp))==0) {
                        VariantClear(&astemp);
                        V_VT(&astemp) = VT_EMPTY;
                    }
            } else {
                V_VT(&astemp) = VT_EMPTY;
            }
            /*MSO_TO_OO_I_Workbooks_Add */
            hr = MSO_TO_OO_I_Workbooks_Add(iface,astemp,&iapp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (iapp==NULL) TRACE("ERROR iapp is NULL \n");
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)iapp;
            }
            return S_OK;
        case dispid_workbooks__Open:
            /*Зависит от кол-ва посланных параметров*/
            if (pDispParams->cArgs==0) return E_FAIL;
            /*Используем только имя файла*/
            VariantInit(&vnull);
            hr = MSO_TO_OO_I_Workbooks_Open(iface, V_BSTR(&(pDispParams->rgvarg[pDispParams->cArgs-1])), vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, l, &iapp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)iapp;
            }
            return S_OK;
        case dispid_workbooks_Close:
            MSO_TO_OO_I_Workbooks_Close(iface, l);
            return S_OK;
        case dispid_workbooks_Count:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                return E_NOTIMPL;
            } else {
                iresult = -1;
                hr = MSO_TO_OO_I_Workbooks_get_Count(iface, &iresult);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription = SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_I2;
                    V_I2(pVarResult) = iresult;
                }
                return hr;
            }
        case dispid_workbooks_application:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                return E_NOTIMPL;
            } else {
                hr = MSO_TO_OO_I_Workbooks_get_Application(iface,&iapp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)iapp;
                }
                return S_OK;
            }
        case dispid_workbooks_parent:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                return E_NOTIMPL;
            } else {
                hr = MSO_TO_OO_I_Workbooks_get_Parent(iface,&iapp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)iapp;
                }
                return S_OK;
            }
        case dispid_workbooks_Open:
            /*Зависит от кол-ва посланных параметров*/
            if (pDispParams->cArgs==0) return E_FAIL;
            /*Используем только имя файла*/
            VariantInit(&vnull);
            hr = MSO_TO_OO_I_Workbooks_Open(iface, V_BSTR(&(pDispParams->rgvarg[pDispParams->cArgs-1])), vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull,l, &iapp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)iapp;
            }
            return S_OK;
        case dispid_workbooks_OpenText:
            /*MSO_TO_OO_I_Workbooks_OpenText*/
            TRACE("Stub: METHOD OpenText \n");
            return S_OK;
        case dispid_workbooks__OpenText:
            /*MSO_TO_OO_I_Workbooks__OpenText*/
            TRACE("Stub: METHOD _OpenText \n");
            return S_OK;
        case dispid_workbooks_OpenXML:
            /*MSO_TO_OO_I_Workbooks_OpenXML*/
            TRACE("Stub: METHOD OpenXML \n");
            return S_OK;
        case dispid_workbooks__OpenXML:
            /*MSO_TO_OO_I_Workbooks__OpenXML*/
            TRACE("Stub: METHOD _OpenXML \n");
            return S_OK;
        case dispid_workbooks_OpenDatabase:
            /*MSO_TO_OO_I_Workbooks_OpenDatabase*/
            TRACE("Stub: METHOD OpenDatabase \n");
            return S_OK;
        case dispid_workbooks_CanCheckOut:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                if (pDispParams->cArgs!=2) return E_FAIL;
                hr = MSO_TO_OO_I_Workbooks_put_CanCheckOut(iface, V_BSTR(&(pDispParams->rgvarg[1])), V_BOOL(&(pDispParams->rgvarg[0])));
                return hr;
            } else {
                if (pDispParams->cArgs!=1) return E_FAIL;
                hr = MSO_TO_OO_I_Workbooks_get_CanCheckOut(iface, V_BSTR(&(pDispParams->rgvarg[0])), &vbool);
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_BOOL;
                    V_BOOL(pVarResult) = vbool;
                }
                return hr;
            }
        case dispid_workbooks_CheckOut:
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = MSO_TO_OO_I_Workbooks_CheckOut(iface, V_BSTR(&(pDispParams->rgvarg[0])));
            return hr;
        case dispid_workbooks_creator:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                /*это свойство только для чтения*/
                return E_NOTIMPL;
            } else {
               hr = MSO_TO_OO_I_Workbooks_get_Creator(iface, &vresult);
               *pVarResult = vresult;
               return hr;
            }
        case dispid_workbooks__Default:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                return E_NOTIMPL;
            } else {
                /*MSO_TO_OO_I_Workbooks_get__Default*/
                TRACE("Stub: PPROPERTY _Default \n");
                return S_OK;
            }
        case dispid_workbooks_Item:
            if (wFlags==DISPATCH_PROPERTYPUT) {
                return S_OK;
            } else {
                /*MSO_TO_OO_I_Workbooks_get_Item*/
                TRACE("Stub: PROPERTY Item \n");
                return S_OK;
            }
        case dispid_workbooks___OpenText:
            /*MSO_TO_OO_I_Workbooks___OpenText*/
            TRACE("Stub: METHOD __OpenText \n");
            return S_OK;
    }
    
    return E_NOTIMPL;
}

const I_WorkbooksVtbl MSO_TO_OO_I_WorkbooksVtbl =
{
    MSO_TO_OO_I_Workbooks_QueryInterface,
    MSO_TO_OO_I_Workbooks_AddRef,
    MSO_TO_OO_I_Workbooks_Release,
    MSO_TO_OO_I_Workbooks_GetTypeInfoCount,
    MSO_TO_OO_I_Workbooks_GetTypeInfo,
    MSO_TO_OO_I_Workbooks_GetIDsOfNames,
    MSO_TO_OO_I_Workbooks_Invoke,
    MSO_TO_OO_I_Workbooks_get_Application,
    MSO_TO_OO_I_Workbooks_get_Creator,
    MSO_TO_OO_I_Workbooks_get_Parent,
    MSO_TO_OO_I_Workbooks_Add,
    MSO_TO_OO_I_Workbooks_Close,
    MSO_TO_OO_I_Workbooks_get_Count,
    MSO_TO_OO_I_Workbooks_get_Item,
    MSO_TO_OO_I_Workbooks_get__NewEnum,
    MSO_TO_OO_I_Workbooks__Open,
    MSO_TO_OO_I_Workbooks___OpenText,
    MSO_TO_OO_I_Workbooks_get__Default,
    MSO_TO_OO_I_Workbooks__OpenText,
    MSO_TO_OO_I_Workbooks_Open,
    MSO_TO_OO_I_Workbooks_OpenText,
    MSO_TO_OO_I_Workbooks_OpenDatabase,
    MSO_TO_OO_I_Workbooks_CheckOut,
    MSO_TO_OO_I_Workbooks_get_CanCheckOut,
    MSO_TO_OO_I_Workbooks_put_CanCheckOut,
    MSO_TO_OO_I_Workbooks__OpenXML,
    MSO_TO_OO_I_Workbooks_OpenXML
};

extern HRESULT _I_WorkbooksConstructor(LPVOID *ppObj)
{
    WorkbooksImpl *workbooks;

    TRACE("(%p)\n", ppObj);

    workbooks = HeapAlloc(GetProcessHeap(), 0, sizeof(*workbooks));
    if (!workbooks) {
        return E_OUTOFMEMORY;
    }

    workbooks->_workbooksVtbl = &MSO_TO_OO_I_WorkbooksVtbl;
    workbooks->ref = 0;
    workbooks->pApplication = NULL;
    workbooks->count_workbooks = 0;
    workbooks->pworkbook = NULL;
    workbooks->current_workbook = -1;

    *ppObj = &workbooks->_workbooksVtbl;

    return S_OK;
}

