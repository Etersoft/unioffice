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

ITypeInfo *ti_workbook = NULL;

HRESULT get_typeinfo_workbook(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_workbook) {
        *typeinfo = ti_workbook;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Workbook, &ti_workbook);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_workbook;
    return hres;
}

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

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Workbook)) {
        *ppvObject = &This->pworkbookVtbl;
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
        TRACE("(%p) (%p) (%p) (%p)\n", iface, This, This->pDoc, This->pSheets);
        if (This->pworkbooks != NULL) {
            I_Workbooks_Release(This->pworkbooks);
            This->pworkbooks = NULL;
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
        DELETE_OBJECT;
    }
    return ref;
}

/*** I_Workbook methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Sheets(
        I_Workbook* iface,
        IDispatch **ppSheets)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    TRACE_IN;

    if (This->pSheets == NULL) 
        return E_FAIL;

    *ppSheets = This->pSheets;
    IDispatch_AddRef(This->pSheets);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_WorkSheets(
        I_Workbook* iface,
        IDispatch **ppSheets)
{
    TRACE(" ----> get_Sheets \n");
    return MSO_TO_OO_I_Workbook_get_Sheets(iface, ppSheets);
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Close(
        I_Workbook* iface,
        VARIANT SaveChanges,
        VARIANT Filename,
        VARIANT RouteWorkbook,
        LCID lcid)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    WorkbooksImpl *This_wbks = (WorkbooksImpl*)(This->pworkbooks);
    
    BSTR filename;
    HRESULT hres;
    int i;
    IDispatch *pdtmp;

    TRACE_IN;

    filename = SysAllocString(L"");
    hres = MSO_TO_OO_CloseWorkbook(iface, filename);
    SysFreeString(filename);
    if (FAILED(hres)) {
        TRACE("\n");
        return E_FAIL;
    }

    /*Find this workbook in array? and delete it*/
    int find_index=-1;
    for (i=0; i<This_wbks->count_workbooks; i++) {
        if (This_wbks->pworkbook[i] == iface) {
            find_index = i;
        }
    }
    
    if (find_index < 0) ERR("Workbook NOT FIND \n");

    hres = MSO_TO_OO_I_Workbook_Release(iface);
    iface = NULL;

    for (i = find_index; i < This_wbks->count_workbooks - 1; i++) {
        This_wbks->pworkbook[i] = This_wbks->pworkbook[i + 1];
    }
    This_wbks->count_workbooks = This_wbks->count_workbooks - 1;
    

    TRACE_OUT;
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
        VARIANT Local,
        LCID    lcid)
{
/*Пока игнорируем все параметры кроме первого*/
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT res, p3, p1;
    HRESULT hres;
    long ix = 0;
    SAFEARRAY FAR* pPropVals;
    BSTR FilenameURL;
    WorkbooksImpl *wbks = (WorkbooksImpl*)(This->pworkbooks);
    TRACE_IN;

    MSO_TO_OO_CorrectArg(Filename, &Filename);
    MSO_TO_OO_CorrectArg(FileFormat, &FileFormat);
    MSO_TO_OO_CorrectArg(Password, &Password);
    MSO_TO_OO_CorrectArg(WriteResPassword, &WriteResPassword);
    MSO_TO_OO_CorrectArg(ReadOnlyRecommended, &ReadOnlyRecommended);
    MSO_TO_OO_CorrectArg(CreateBackup, &CreateBackup);
    MSO_TO_OO_CorrectArg(ConflictResolution, &ConflictResolution);
    MSO_TO_OO_CorrectArg(AddToMru, &AddToMru);
    MSO_TO_OO_CorrectArg(TextCodepage, &TextCodepage);
    MSO_TO_OO_CorrectArg(TextVisualLayout, &TextVisualLayout);
    MSO_TO_OO_CorrectArg(Local, &Local);

    if (This==NULL) {
        TRACE("ERROR object is NULL \n");
        return E_FAIL;
    }
    if (V_VT(&Filename)!=VT_BSTR) {
        TRACE("ERROR no filename \n");
        return E_FAIL;
    }

    /* Create PropertyValue with save-format-data */
    IDispatch *dpv;
    MSO_TO_OO_GetDispatchPropertyValue((_Application*)(wbks->pApplication), &dpv);
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
    int i=0;
    while (*(FilenameURL+i)!=0) {
        WTRACE(L"%c \n",*(FilenameURL+i));
        i++;
    }
    TRACE("\n");
    /* Call StoreToURL for save document to file */
    hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"StoreAsURL", 2, p3, p1);
    if (FAILED(hres)) {
        TRACE("ERROR when StoreAsURL \n");
        return hres;
    }
    VariantClear(&res);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Save(
        I_Workbook* iface,
        LCID lcid)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT res;
    HRESULT hres;
    TRACE_IN;

    if (This==NULL) {
        TRACE("ERROR objetct is NULL \n");
        return E_FAIL;
    }

    /* Call Store for save document to file */
    hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"Store",0);
    if (FAILED(hres)) {
        TRACE("ERROR when StoreAsURL \n");
        return hres;
    }
    VariantClear(&res);

    TRACE_OUT;
    return S_OK;
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
    TRACE_IN;

    MSO_TO_OO_CorrectArg(Password, &Password);
    MSO_TO_OO_CorrectArg(Structure, &Structure);
    MSO_TO_OO_CorrectArg(Windows, &Windows);

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    VariantInit(&param);
    if (Is_Variant_Null(Password)) {
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
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Unprotect(
        I_Workbook* iface,
        VARIANT Password,
        LCID lcid)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT param, res;
    HRESULT hres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(Password, &Password);

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    VariantInit(&param);
    if (Is_Variant_Null(Password)) {
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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Name(
        I_Workbook* iface,
        BSTR *retval)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT vRes, ooframe, oocontr;
    HRESULT hres;
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Names(
        I_Workbook* iface,
        IDispatch **retval)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;

    HRESULT hres;
    IUnknown *punk = NULL;
    IDispatch *pNames;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *retval = NULL;

    hres = _NamesConstructor((LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = Names_QueryInterface(punk, &IID_Names, (void**) &pNames);
    if (pNames == NULL) {
        return E_FAIL;
    }

    hres = MSO_TO_OO_Names_Initialize((Names*)pNames, iface);

    if (FAILED(hres)) {
        IDispatch_Release(pNames);
        TRACE("Error when MSO_TO_OO_Names_Initialize \n");
        return hres;
    }

    *retval = pNames;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Application(
        I_Workbook* iface,
        IDispatch **RHS)
{   
    WorkbookImpl *This = (WorkbookImpl*)iface;
    TRACE_IN;
    
    *RHS = NULL;
    
    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    TRACE_OUT;
    return I_Workbooks_get_Application((I_Workbooks*)(This->pworkbooks), RHS);
    
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Creator(
        I_Workbook* iface,
        XlCreator *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Parent(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_AcceptLabelsInFormulas(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_AcceptLabelsInFormulas(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Activate(
        I_Workbook* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ActiveChart(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ActiveSheet(
        I_Workbook* iface,
        IDispatch **RHS)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    I_Sheets *pSheets;
    HRESULT hres;
    TRACE_IN;

    hres = I_Workbook_get_Sheets(iface, (IDispatch**) &pSheets);
    if (FAILED(hres)) {
        *RHS = NULL;
        return E_FAIL;
    }

    hres = MSO_TO_OO_GetActiveSheet(pSheets, (I_Worksheet**)RHS);
    if (FAILED(hres)) {
        I_Sheets_Release(pSheets);
        *RHS = NULL;
        return hres;
    }
    I_Sheets_Release(pSheets);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Author(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Author(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_AutoUpdateFrequency(
        I_Workbook* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_AutoUpdateFrequency(
        I_Workbook* iface,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_AutoUpdateSaveChanges(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_AutoUpdateSaveChanges(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ChangeHistoryDuration(
        I_Workbook* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_ChangeHistoryDuration(
        I_Workbook* iface,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_BuiltinDocumentProperties(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ChangeFileAccess(
        I_Workbook* iface,
        XlFileAccess Mode,
        VARIANT WritePassword,
        VARIANT Notify,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ChangeLink(
        I_Workbook* iface,
        BSTR Name,
        BSTR NewName,
        XlLinkType Type,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Charts(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_CodeName(
        I_Workbook* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get__CodeName(
        I_Workbook* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put__CodeName(
        I_Workbook* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Colors(
        I_Workbook* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Colors(
        I_Workbook* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_CommandBars(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Comments(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Comments(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ConflictResolution(
        I_Workbook* iface,
        XlSaveConflictResolution *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_ConflictResolution(
        I_Workbook* iface,
        XlSaveConflictResolution RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Container(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_CreateBackup(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_CustomDocumentProperties(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Date1904(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Date1904(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_DeleteNumberFormat(
        I_Workbook* iface,
        BSTR NumberFormat,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_DialogSheets(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_DisplayDrawingObjects(
        I_Workbook* iface,
        LCID lcid,
        XlDisplayDrawingObjects *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_DisplayDrawingObjects(
        I_Workbook* iface,
        LCID lcid,
        XlDisplayDrawingObjects RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ExclusiveAccess(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_FileFormat(
        I_Workbook* iface,
        LCID lcid,
        XlFileFormat *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ForwardMailer(
        I_Workbook* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_FullName(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_HasMailer(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_HasMailer(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_HasPassword(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_HasRoutingSlip(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_HasRoutingSlip(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_IsAddin(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_IsAddin(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Keywords(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Keywords(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_LinkInfo(
        I_Workbook* iface,
        BSTR Name,
        XlLinkInfo LinkInfo,
        VARIANT Type,
        VARIANT EditionRef,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_LinkSources(
        I_Workbook* iface,
        VARIANT Type,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Mailer(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_MergeWorkbook(
        I_Workbook* iface,
        VARIANT Filename)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Modules(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_MultiUserEditing(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_NewWindow(
        I_Workbook* iface,
        LCID lcid,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_OnSave(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_OnSave(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_OnSheetActivate(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_OnSheetActivate(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_OnSheetDeactivate(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_OnSheetDeactivate(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_OpenLinks(
        I_Workbook* iface,
        BSTR Name,
        VARIANT ReadOnly,
        VARIANT Type,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Path(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PersonalViewListSettings(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_PersonalViewListSettings(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PersonalViewPrintSettings(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_PersonalViewPrintSettings(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_PivotCaches(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Post(
        I_Workbook* iface,
        VARIANT DestName,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PrecisionAsDisplayed(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_PrecisionAsDisplayed(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook__PrintOut(
        I_Workbook* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_PrintPreview(
        I_Workbook* iface,
        VARIANT EnableChanges,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook__Protect(
        I_Workbook* iface,
        VARIANT Password,
        VARIANT Structure,
        VARIANT Windows)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ProtectSharing(
        I_Workbook* iface,
        VARIANT Filename,
        VARIANT Password,
        VARIANT WriteResPassword,
        VARIANT ReadOnlyRecommended,
        VARIANT CreateBackup,
        VARIANT SharingPassword)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ProtectStructure(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ProtectWindows(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ReadOnly(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get__ReadOnlyRecommended(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_RefreshAll(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Reply(
        I_Workbook* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ReplyAll(
        I_Workbook* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_RemoveUser(
        I_Workbook* iface,
        long Index)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_RevisionNumber(
        I_Workbook* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Route(
        I_Workbook* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Routed(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_RoutingSlip(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_RunAutoMacros(
        I_Workbook* iface,
        XlRunAutoMacro Which,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook__SaveAs(
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
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SaveCopyAs(
        I_Workbook* iface,
        VARIANT Filename,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Saved(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Saved(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_SaveLinkValues(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_SaveLinkValues(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SendMail(
        I_Workbook* iface,
        VARIANT Recipients,
        VARIANT Subject,
        VARIANT ReturnReceipt,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SendMailer(
        I_Workbook* iface,
        VARIANT FileFormat,
        XlPriority Priority,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SetLinkOnData(
        I_Workbook* iface,
        BSTR Name,
        VARIANT Procedure,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ShowConflictHistory(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_ShowConflictHistory(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Styles(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Subject(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Subject(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Title(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Title(
        I_Workbook* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_UnprotectSharing(
        I_Workbook* iface,
        VARIANT SharingPassword)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_UpdateFromFile(
        I_Workbook* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_UpdateLink(
        I_Workbook* iface,
        VARIANT Name,
        VARIANT Type,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_UpdateRemoteReferences(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_UpdateRemoteReferences(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_UserControl(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_UserControl(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_UserStatus(
        I_Workbook* iface,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_CustomViews(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Windows(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_WriteReserved(
        I_Workbook* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_WriteReservedBy(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Excel4IntlMacroSheets(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Excel4MacroSheets(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_TemplateRemoveExtData(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_TemplateRemoveExtData(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_HighlightChangesOptions(
        I_Workbook* iface,
        VARIANT When,
        VARIANT Who,
        VARIANT Where)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_HighlightChangesOnScreen(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_HighlightChangesOnScreen(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_KeepChangeHistory(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_KeepChangeHistory(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ListChangesOnNewSheet(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_ListChangesOnNewSheet(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_PurgeChangeHistoryNow(
        I_Workbook* iface,
        long Days,
        VARIANT SharingPassword)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_AcceptAllChanges(
        I_Workbook* iface,
        VARIANT When,
        VARIANT Who,
        VARIANT Where)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_RejectAllChanges(
        I_Workbook* iface,
        VARIANT When,
        VARIANT Who,
        VARIANT Where)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_PivotTableWizard(
        I_Workbook* iface,
        VARIANT SourceType,
        VARIANT SourceData,
        VARIANT TableDestination,
        VARIANT TableName,
        VARIANT RowGrand,
        VARIANT ColumnGrand,
        VARIANT SaveData,
        VARIANT HasAutoFormat,
        VARIANT AutoPage,
        VARIANT Reserved,
        VARIANT BackgroundQuery,
        VARIANT OptimizeCache,
        VARIANT PageFieldOrder,
        VARIANT PageFieldWrapCount,
        VARIANT ReadData,
        VARIANT Connection,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ResetColors(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_VBProject(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_FollowHyperlink(
        I_Workbook* iface,
        BSTR Address,
        VARIANT SubAddress,
        VARIANT NewWindow,
        VARIANT AddHistory,
        VARIANT ExtraInfo,
        VARIANT Method,
        VARIANT HeaderInfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_AddToFavorites(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_IsInplace(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_PrintOut(
        I_Workbook* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        VARIANT PrToFileName,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_WebPagePreview(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PublishObjects(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_WebOptions(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ReloadAs(
        I_Workbook* iface,
        MsoEncoding Encoding)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_HTMLProject(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_EnvelopeVisible(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_EnvelopeVisible(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_CalculationVersion(
        I_Workbook* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Dummy17(
        I_Workbook* iface,
        long calcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_sblt(
        I_Workbook* iface,
        BSTR s)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_VBASigned(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ShowPivotTableFieldList(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_ShowPivotTableFieldList(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_UpdateLinks(
        I_Workbook* iface,
        XlUpdateLinks *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_UpdateLinks(
        I_Workbook* iface,
        XlUpdateLinks RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_BreakLink(
        I_Workbook* iface,
        BSTR Name,
        XlLinkType Type)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_Dummy16(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_EnableAutoRecover(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_EnableAutoRecover(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_RemovePersonalInformation(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_RemovePersonalInformation(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_FullNameURLEncoded(
        I_Workbook* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_CheckIn(
        I_Workbook* iface,
        VARIANT SaveChanges,
        VARIANT Comments,
        VARIANT MakePublic)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_CanCheckIn(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SendForReview(
        I_Workbook* iface,
        VARIANT Recipients,
        VARIANT Subject,
        VARIANT ShowMessage,
        VARIANT IncludeAttachment)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ReplyWithChanges(
        I_Workbook* iface,
        VARIANT ShowMessage)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_EndReview(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Password(
        I_Workbook* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_Password(
        I_Workbook* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_WritePassword(
        I_Workbook* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_WritePassword(
        I_Workbook* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PasswordEncryptionProvider(
        I_Workbook* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PasswordEncryptionAlgorithm(
        I_Workbook* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PasswordEncryptionKeyLength(
        I_Workbook* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SetPasswordEncryptionOptions(
        I_Workbook* iface,
        VARIANT PasswordEncryptionProvider,
        VARIANT PasswordEncryptionAlgorithm,
        VARIANT PasswordEncryptionKeyLength,
        VARIANT PasswordEncryptionFileProperties)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_PasswordEncryptionFileProperties(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_ReadOnlyRecommended(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_ReadOnlyRecommended(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_SmartTagOptions(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_RecheckSmartTags(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Permission(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_SharedWorkspace(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_Sync(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SendFaxOverInternet(
        I_Workbook* iface,
        VARIANT Recipients,
        VARIANT Subject,
        VARIANT ShowMessage)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_XmlNamespaces(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_XmlMaps(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_XmlImport(
        I_Workbook* iface,
        BSTR Url,
        IDispatch **ImportMap,
        VARIANT Overwrite,
        VARIANT Destination,
        XlXmlImportResult *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_SmartDocument(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_DocumentLibraryVersions(
        I_Workbook* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_InactiveListBorderVisible(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_InactiveListBorderVisible(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_get_DisplayInkComments(
        I_Workbook* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_put_DisplayInkComments(
        I_Workbook* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_XmlImportXml(
        I_Workbook* iface,
        BSTR Data,
        IDispatch **ImportMap,
        VARIANT Overwrite,
        VARIANT Destination,
        XlXmlImportResult *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_SaveAsXMLData(
        I_Workbook* iface,
        BSTR Filename,
        IDispatch *Map)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_ToggleFormsDesign(
        I_Workbook* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetTypeInfoCount(
        I_Workbook* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}


static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetTypeInfo(
        I_Workbook* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_workbook(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbook_GetIDsOfNames(
        I_Workbook* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_workbook(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
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
    HRESULT hres;
    IDispatch *drets;
    IDispatch *dret;
    VARIANT vmas[12];
    VARIANT vNull;
    int i;
    VARIANT vtmp;
    BSTR bret;
    ITypeInfo *typeinfo;

    for (i=0;i<12;i++) {
         VariantInit(&vmas[i]);
         V_VT(&vmas[i])=VT_EMPTY;
    }

    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;
    TRACE("\n");

    if (This == NULL) return E_POINTER;

    switch (dispIdMember) 
    {
    case dispid_workbook_close:
        if (pDispParams->cArgs>3) {
            TRACE(" (3) ERROR Parameters");
            return E_FAIL;
        }
        /*необходимо перевернуть параметры*/
        for (i=0;i<pDispParams->cArgs;i++) {
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-i-1], &vmas[i]))) return E_FAIL;
        }
        hres = MSO_TO_OO_I_Workbook_Close(iface, vmas[0], vmas[1], vmas[2], 0);
        if (FAILED(hres)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hres;
        }
        return hres;
    case dispid_workbook_saveas:
        if (pDispParams->cArgs>12) {
            TRACE(" (4) ERROR Parameters");
            return E_FAIL;
        }
        /*необходимо перевернуть параметры*/
        for (i=0;i<pDispParams->cArgs;i++) {
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-i-1], &vmas[i]))) return E_FAIL;
        }
        VariantChangeTypeEx(&vmas[6], &vmas[6], 0, 0, VT_I4);
        hres = MSO_TO_OO_I_Workbook_SaveAs(iface, vmas[0], vmas[1], vmas[2], vmas[3], vmas[4], vmas[5], V_I4(&vmas[6]), vmas[7], vmas[8], vmas[9], vmas[10], vmas[11], 0);
        if (FAILED(hres)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hres;
        }
        return hres;
    case dispid_workbook_save:
        return MSO_TO_OO_I_Workbook_Save(iface, 0);
    case dispid_workbook_unprotect://UnProtect
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
        return MSO_TO_OO_I_Workbook_Unprotect(iface,vtmp, 0);
    default:
        hres = get_typeinfo_workbook(&typeinfo);
        if(FAILED(hres))
            return hres;

        hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
        if (FAILED(hres)) {
            TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
        }

        return hres;
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
    MSO_TO_OO_I_Workbook_get_Application,
    MSO_TO_OO_I_Workbook_get_Creator,
    MSO_TO_OO_I_Workbook_get_Parent,
    MSO_TO_OO_I_Workbook_get_AcceptLabelsInFormulas,
    MSO_TO_OO_I_Workbook_put_AcceptLabelsInFormulas,
    MSO_TO_OO_I_Workbook_Activate,
    MSO_TO_OO_I_Workbook_get_ActiveChart,
    MSO_TO_OO_I_Workbook_get_ActiveSheet,
    MSO_TO_OO_I_Workbook_get_Author,
    MSO_TO_OO_I_Workbook_put_Author,
    MSO_TO_OO_I_Workbook_get_AutoUpdateFrequency,
    MSO_TO_OO_I_Workbook_put_AutoUpdateFrequency,
    MSO_TO_OO_I_Workbook_get_AutoUpdateSaveChanges,
    MSO_TO_OO_I_Workbook_put_AutoUpdateSaveChanges,
    MSO_TO_OO_I_Workbook_get_ChangeHistoryDuration,
    MSO_TO_OO_I_Workbook_put_ChangeHistoryDuration,
    MSO_TO_OO_I_Workbook_get_BuiltinDocumentProperties,
    MSO_TO_OO_I_Workbook_ChangeFileAccess,
    MSO_TO_OO_I_Workbook_ChangeLink,
    MSO_TO_OO_I_Workbook_get_Charts,
    MSO_TO_OO_I_Workbook_Close,
    MSO_TO_OO_I_Workbook_get_CodeName,
    MSO_TO_OO_I_Workbook_get__CodeName,
    MSO_TO_OO_I_Workbook_put__CodeName,
    MSO_TO_OO_I_Workbook_get_Colors,
    MSO_TO_OO_I_Workbook_put_Colors,
    MSO_TO_OO_I_Workbook_get_CommandBars,
    MSO_TO_OO_I_Workbook_get_Comments,
    MSO_TO_OO_I_Workbook_put_Comments,
    MSO_TO_OO_I_Workbook_get_ConflictResolution,
    MSO_TO_OO_I_Workbook_put_ConflictResolution,
    MSO_TO_OO_I_Workbook_get_Container,
    MSO_TO_OO_I_Workbook_get_CreateBackup,
    MSO_TO_OO_I_Workbook_get_CustomDocumentProperties,
    MSO_TO_OO_I_Workbook_get_Date1904,
    MSO_TO_OO_I_Workbook_put_Date1904,
    MSO_TO_OO_I_Workbook_DeleteNumberFormat,
    MSO_TO_OO_I_Workbook_get_DialogSheets,
    MSO_TO_OO_I_Workbook_get_DisplayDrawingObjects,
    MSO_TO_OO_I_Workbook_put_DisplayDrawingObjects,
    MSO_TO_OO_I_Workbook_ExclusiveAccess,
    MSO_TO_OO_I_Workbook_get_FileFormat,
    MSO_TO_OO_I_Workbook_ForwardMailer,
    MSO_TO_OO_I_Workbook_get_FullName,
    MSO_TO_OO_I_Workbook_get_HasMailer,
    MSO_TO_OO_I_Workbook_put_HasMailer,
    MSO_TO_OO_I_Workbook_get_HasPassword,
    MSO_TO_OO_I_Workbook_get_HasRoutingSlip,
    MSO_TO_OO_I_Workbook_put_HasRoutingSlip,
    MSO_TO_OO_I_Workbook_get_IsAddin,
    MSO_TO_OO_I_Workbook_put_IsAddin,
    MSO_TO_OO_I_Workbook_get_Keywords,
    MSO_TO_OO_I_Workbook_put_Keywords,
    MSO_TO_OO_I_Workbook_LinkInfo,
    MSO_TO_OO_I_Workbook_LinkSources,
    MSO_TO_OO_I_Workbook_get_Mailer,
    MSO_TO_OO_I_Workbook_MergeWorkbook,
    MSO_TO_OO_I_Workbook_get_Modules,
    MSO_TO_OO_I_Workbook_get_MultiUserEditing,
    MSO_TO_OO_I_Workbook_get_Name,
    MSO_TO_OO_I_Workbook_get_Names,
    MSO_TO_OO_I_Workbook_NewWindow,
    MSO_TO_OO_I_Workbook_get_OnSave,
    MSO_TO_OO_I_Workbook_put_OnSave,
    MSO_TO_OO_I_Workbook_get_OnSheetActivate,
    MSO_TO_OO_I_Workbook_put_OnSheetActivate,
    MSO_TO_OO_I_Workbook_get_OnSheetDeactivate,
    MSO_TO_OO_I_Workbook_put_OnSheetDeactivate,
    MSO_TO_OO_I_Workbook_OpenLinks,
    MSO_TO_OO_I_Workbook_get_Path,
    MSO_TO_OO_I_Workbook_get_PersonalViewListSettings,
    MSO_TO_OO_I_Workbook_put_PersonalViewListSettings,
    MSO_TO_OO_I_Workbook_get_PersonalViewPrintSettings,
    MSO_TO_OO_I_Workbook_put_PersonalViewPrintSettings,
    MSO_TO_OO_I_Workbook_PivotCaches,
    MSO_TO_OO_I_Workbook_Post,
    MSO_TO_OO_I_Workbook_get_PrecisionAsDisplayed,
    MSO_TO_OO_I_Workbook_put_PrecisionAsDisplayed,
    MSO_TO_OO_I_Workbook__PrintOut,
    MSO_TO_OO_I_Workbook_PrintPreview,
    MSO_TO_OO_I_Workbook__Protect,
    MSO_TO_OO_I_Workbook_ProtectSharing,
    MSO_TO_OO_I_Workbook_get_ProtectStructure,
    MSO_TO_OO_I_Workbook_get_ProtectWindows,
    MSO_TO_OO_I_Workbook_get_ReadOnly,
    MSO_TO_OO_I_Workbook_get__ReadOnlyRecommended,
    MSO_TO_OO_I_Workbook_RefreshAll,
    MSO_TO_OO_I_Workbook_Reply,
    MSO_TO_OO_I_Workbook_ReplyAll,
    MSO_TO_OO_I_Workbook_RemoveUser,
    MSO_TO_OO_I_Workbook_get_RevisionNumber,
    MSO_TO_OO_I_Workbook_Route,
    MSO_TO_OO_I_Workbook_get_Routed,
    MSO_TO_OO_I_Workbook_get_RoutingSlip,
    MSO_TO_OO_I_Workbook_RunAutoMacros,
    MSO_TO_OO_I_Workbook_Save,
    MSO_TO_OO_I_Workbook__SaveAs,
    MSO_TO_OO_I_Workbook_SaveCopyAs,
    MSO_TO_OO_I_Workbook_get_Saved,
    MSO_TO_OO_I_Workbook_put_Saved,
    MSO_TO_OO_I_Workbook_get_SaveLinkValues,
    MSO_TO_OO_I_Workbook_put_SaveLinkValues,
    MSO_TO_OO_I_Workbook_SendMail,
    MSO_TO_OO_I_Workbook_SendMailer,
    MSO_TO_OO_I_Workbook_SetLinkOnData,
    MSO_TO_OO_I_Workbook_get_Sheets,
    MSO_TO_OO_I_Workbook_get_ShowConflictHistory,
    MSO_TO_OO_I_Workbook_put_ShowConflictHistory,
    MSO_TO_OO_I_Workbook_get_Styles,
    MSO_TO_OO_I_Workbook_get_Subject,
    MSO_TO_OO_I_Workbook_put_Subject,
    MSO_TO_OO_I_Workbook_get_Title,
    MSO_TO_OO_I_Workbook_put_Title,
    MSO_TO_OO_I_Workbook_Unprotect,
    MSO_TO_OO_I_Workbook_UnprotectSharing,
    MSO_TO_OO_I_Workbook_UpdateFromFile,
    MSO_TO_OO_I_Workbook_UpdateLink,
    MSO_TO_OO_I_Workbook_get_UpdateRemoteReferences,
    MSO_TO_OO_I_Workbook_put_UpdateRemoteReferences,
    MSO_TO_OO_I_Workbook_get_UserControl,
    MSO_TO_OO_I_Workbook_put_UserControl,
    MSO_TO_OO_I_Workbook_get_UserStatus,
    MSO_TO_OO_I_Workbook_get_CustomViews,
    MSO_TO_OO_I_Workbook_get_Windows,
    MSO_TO_OO_I_Workbook_get_WorkSheets,
    MSO_TO_OO_I_Workbook_get_WriteReserved,
    MSO_TO_OO_I_Workbook_get_WriteReservedBy,
    MSO_TO_OO_I_Workbook_get_Excel4IntlMacroSheets,
    MSO_TO_OO_I_Workbook_get_Excel4MacroSheets,
    MSO_TO_OO_I_Workbook_get_TemplateRemoveExtData,
    MSO_TO_OO_I_Workbook_put_TemplateRemoveExtData,
    MSO_TO_OO_I_Workbook_HighlightChangesOptions,
    MSO_TO_OO_I_Workbook_get_HighlightChangesOnScreen,
    MSO_TO_OO_I_Workbook_put_HighlightChangesOnScreen,
    MSO_TO_OO_I_Workbook_get_KeepChangeHistory,
    MSO_TO_OO_I_Workbook_put_KeepChangeHistory,
    MSO_TO_OO_I_Workbook_get_ListChangesOnNewSheet,
    MSO_TO_OO_I_Workbook_put_ListChangesOnNewSheet,
    MSO_TO_OO_I_Workbook_PurgeChangeHistoryNow,
    MSO_TO_OO_I_Workbook_AcceptAllChanges,
    MSO_TO_OO_I_Workbook_RejectAllChanges,
    MSO_TO_OO_I_Workbook_PivotTableWizard,
    MSO_TO_OO_I_Workbook_ResetColors,
    MSO_TO_OO_I_Workbook_get_VBProject,
    MSO_TO_OO_I_Workbook_FollowHyperlink,
    MSO_TO_OO_I_Workbook_AddToFavorites,
    MSO_TO_OO_I_Workbook_get_IsInplace,
    MSO_TO_OO_I_Workbook_PrintOut,
    MSO_TO_OO_I_Workbook_WebPagePreview,
    MSO_TO_OO_I_Workbook_get_PublishObjects,
    MSO_TO_OO_I_Workbook_get_WebOptions,
    MSO_TO_OO_I_Workbook_ReloadAs,
    MSO_TO_OO_I_Workbook_get_HTMLProject,
    MSO_TO_OO_I_Workbook_get_EnvelopeVisible,
    MSO_TO_OO_I_Workbook_put_EnvelopeVisible,
    MSO_TO_OO_I_Workbook_get_CalculationVersion,
    MSO_TO_OO_I_Workbook_Dummy17,
    MSO_TO_OO_I_Workbook_sblt,
    MSO_TO_OO_I_Workbook_get_VBASigned,
    MSO_TO_OO_I_Workbook_get_ShowPivotTableFieldList,
    MSO_TO_OO_I_Workbook_put_ShowPivotTableFieldList,
    MSO_TO_OO_I_Workbook_get_UpdateLinks,
    MSO_TO_OO_I_Workbook_put_UpdateLinks,
    MSO_TO_OO_I_Workbook_BreakLink,
    MSO_TO_OO_I_Workbook_Dummy16,
    MSO_TO_OO_I_Workbook_SaveAs,
    MSO_TO_OO_I_Workbook_get_EnableAutoRecover,
    MSO_TO_OO_I_Workbook_put_EnableAutoRecover,
    MSO_TO_OO_I_Workbook_get_RemovePersonalInformation,
    MSO_TO_OO_I_Workbook_put_RemovePersonalInformation,
    MSO_TO_OO_I_Workbook_get_FullNameURLEncoded,
    MSO_TO_OO_I_Workbook_CheckIn,
    MSO_TO_OO_I_Workbook_CanCheckIn,
    MSO_TO_OO_I_Workbook_SendForReview,
    MSO_TO_OO_I_Workbook_ReplyWithChanges,
    MSO_TO_OO_I_Workbook_EndReview,
    MSO_TO_OO_I_Workbook_get_Password,
    MSO_TO_OO_I_Workbook_put_Password,
    MSO_TO_OO_I_Workbook_get_WritePassword,
    MSO_TO_OO_I_Workbook_put_WritePassword,
    MSO_TO_OO_I_Workbook_get_PasswordEncryptionProvider,
    MSO_TO_OO_I_Workbook_get_PasswordEncryptionAlgorithm,
    MSO_TO_OO_I_Workbook_get_PasswordEncryptionKeyLength,
    MSO_TO_OO_I_Workbook_SetPasswordEncryptionOptions,
    MSO_TO_OO_I_Workbook_get_PasswordEncryptionFileProperties,
    MSO_TO_OO_I_Workbook_get_ReadOnlyRecommended,
    MSO_TO_OO_I_Workbook_put_ReadOnlyRecommended,
    MSO_TO_OO_I_Workbook_Protect,
    MSO_TO_OO_I_Workbook_get_SmartTagOptions,
    MSO_TO_OO_I_Workbook_RecheckSmartTags,
    MSO_TO_OO_I_Workbook_get_Permission,
    MSO_TO_OO_I_Workbook_get_SharedWorkspace,
    MSO_TO_OO_I_Workbook_get_Sync,
    MSO_TO_OO_I_Workbook_SendFaxOverInternet,
    MSO_TO_OO_I_Workbook_get_XmlNamespaces,
    MSO_TO_OO_I_Workbook_get_XmlMaps,
    MSO_TO_OO_I_Workbook_XmlImport,
    MSO_TO_OO_I_Workbook_get_SmartDocument,
    MSO_TO_OO_I_Workbook_get_DocumentLibraryVersions,
    MSO_TO_OO_I_Workbook_get_InactiveListBorderVisible,
    MSO_TO_OO_I_Workbook_put_InactiveListBorderVisible,
    MSO_TO_OO_I_Workbook_get_DisplayInkComments,
    MSO_TO_OO_I_Workbook_put_DisplayInkComments,
    MSO_TO_OO_I_Workbook_XmlImportXml,
    MSO_TO_OO_I_Workbook_SaveAsXMLData,
    MSO_TO_OO_I_Workbook_ToggleFormsDesign
};

extern HRESULT _I_WorkbookConstructor(LPVOID *ppObj)
{
    WorkbookImpl *workbook;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    workbook = HeapAlloc(GetProcessHeap(), 0, sizeof(*workbook));
    if (!workbook)
    {
        return E_OUTOFMEMORY;
    }

    workbook->pworkbookVtbl = &MSO_TO_OO_I_WorkbookVtbl;
    workbook->ref = 0;
    workbook->pworkbooks = NULL;
    workbook->pDoc = NULL;
    workbook->pSheets = NULL;

    *ppObj = &workbook->pworkbookVtbl;
    
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}
