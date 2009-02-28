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

ITypeInfo *ti_workbooks = NULL;

HRESULT get_typeinfo_workbooks(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_workbooks) {
        *typeinfo = ti_workbooks;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Workbooks, &ti_workbooks);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_workbooks;
    return hres;
}

static WCHAR const str_pusto[]= {0};

#define WORKBOOKS_THIS(iface) DEFINE_THIS(WorkbooksImpl, workbooks, iface)

/*** IUnknown methods ***/

static ULONG WINAPI MSO_TO_OO_I_Workbooks_AddRef(
        I_Workbooks* iface)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    ULONG ref;
    TRACE("REF=%i \n", This->ref);

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


static HRESULT WINAPI MSO_TO_OO_I_Workbooks_QueryInterface(
        I_Workbooks* iface,
        REFIID riid,
        void **ppvObject)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);

    if (!This) {
        ERR("object is null \n");
        return E_POINTER;
    }

    if (!ppvObject) {
        ERR("object ppvObject is null \n");
        return E_POINTER;
    }
    
    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Workbooks)) {
        *ppvObject = WORKBOOKS_WORKBOOKS(This);
        I_Workbooks_AddRef(iface);
        return S_OK;
    }

    if (IsEqualGUID(riid, &IID_IEnumVARIANT)) {
        *ppvObject = WORKBOOKS_ENUM(This);
        IUnknown_AddRef((IUnknown*)(*ppvObject));
        return S_OK;
    }

    return E_NOINTERFACE;
}


static ULONG WINAPI MSO_TO_OO_I_Workbooks_Release(
        I_Workbooks* iface)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    ULONG ref;
    int i;

    if (!This) {
        ERR("object is null \n");
        return E_POINTER;
    }

    TRACE("REF = %i \n", This->ref);

    ref = InterlockedDecrement(&This->ref);

    if (ref == 0) {
/*        if (This->pApplication!=NULL) {
            IDispatch_Release(This->pApplication);*/
            This->pApplication==NULL;
/*        }*/
        for (i=0;i<This->count_workbooks;i++)
            if (This->pworkbook[i]!=NULL) I_Workbook_Release(This->pworkbook[i]);
        if (This->count_workbooks>0) HeapFree(GetProcessHeap(),HEAP_ZERO_MEMORY,This->pworkbook);
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
        DELETE_OBJECT;
    }
    return ref;
}

/*** I_Workbooks methods ***/

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_Add(
        I_Workbooks* iface,
        VARIANT varTemplate,
        IDispatch **ppWorkbook)
{
/*TODO :use of templates*/

    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    IUnknown *punk = NULL;
    HRESULT hres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(varTemplate, &varTemplate);

    if (!This) {
        ERR("object is null \n");
        return E_POINTER;
    }

    hres = _I_WorkbookConstructor((LPVOID*) &punk);
    if (FAILED(hres)) {
        ERR("Constructor Failed \n");                  
        return E_NOINTERFACE;
    }

    if (This->capasity_workbooks == 0){
        This->count_workbooks = 1;       
        This->capasity_workbooks = This->count_workbooks;
        This->pworkbook = HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, sizeof(WorkbookImpl*));
        if (!(This->pworkbook)) {
            ERR("Out of memory \n");
            return E_OUTOFMEMORY;
        }
        This->current_workbook = 0;
    } else {
        This->count_workbooks += 1;
        if (This->count_workbooks > This->capasity_workbooks) {
            This->capasity_workbooks = This->count_workbooks;
            This->pworkbook = HeapReAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, This->pworkbook, This->capasity_workbooks * sizeof(WorkbookImpl*));
        }
        if (!(This->pworkbook)) {
            ERR("Out of memory \n");                    
            return E_OUTOFMEMORY;
        }
        This->current_workbook = This->count_workbooks - 1;
    }

    hres = I_Workbook_QueryInterface(punk, &IID_I_Workbook, (void**) &(This->pworkbook[This->current_workbook]));
/*    I_Workbook_Release(punk);*/
    if (FAILED(hres)) {
        ERR(" QueryInterface\n");
        return E_FAIL;
    }
    if (!ppWorkbook) {
        /*Alloc memory to object*/
        ppWorkbook = HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, sizeof(WorkbookImpl*));
        TRACE(" AllocMemory \n");
        *ppWorkbook = (IDispatch*)(This->pworkbook[This->current_workbook]);
    } else {
        *ppWorkbook = (IDispatch*)(This->pworkbook[This->current_workbook]);
    }

    if ((Is_Variant_Null(varTemplate)) || (lstrlenW(V_BSTR(&varTemplate)) == 0)) {
        hres = MSO_TO_OO_I_Workbook_Initialize( This->pworkbook[This->current_workbook], iface);
    } else {
       /* Prepare name of file to right view
       From:    C:\test test.xls
       To:      file:///c:/test%20test.xls */
       BSTR Filename;
       MSO_TO_OO_MakeURLFromFilename(V_BSTR(&varTemplate),&Filename);
       /*Initialize object*/
       WTRACE(L"FILENAME ------>  %s \n",Filename);
       hres = MSO_TO_OO_I_Workbook_Initialize2( This->pworkbook[This->current_workbook], iface, Filename, VARIANT_TRUE);
       SysFreeString(Filename);
    }
    if (FAILED(hres)) {
        *ppWorkbook = NULL;
        ERR("Workbook_Initialize");
        return hres;
    }
    I_Workbook_AddRef(*ppWorkbook);
    TRACE_OUT;
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_Close(
        I_Workbooks* iface,
        LCID lcid)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    int i;
    BSTR filename;
    TRACE_IN;

    if (!This) {
        ERR("object is null \n");
        return E_POINTER;
    }

    filename = SysAllocString(L"");

    for (i=0;i<This->count_workbooks;i++) {
/*        WorkbookImpl *wb = (WorkbookImpl*)(This->pworkbook[i]);
        filename = SysAllocString(wb->filename);*/
        if (This->pworkbook[i]!=NULL) {
            MSO_TO_OO_CloseWorkbook(This->pworkbook[i], filename);
            I_Workbook_Release(This->pworkbook[i]);
            This->pworkbook[i] = NULL;
        }
/*        SysFreeString(filename);*/
    }
    SysFreeString(filename);
    HeapFree(GetProcessHeap(),HEAP_ZERO_MEMORY,This->pworkbook);
    This->count_workbooks = 0;
    This->current_workbook = -1;
    This->pworkbook = NULL;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Count(
        I_Workbooks* iface,
        int *count)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    TRACE_IN;

    if (!This) {
        ERR("object is null \n");
        return E_POINTER;
    }

    *count = This->count_workbooks;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Application(
        I_Workbooks* iface,
        IDispatch **value)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    TRACE_IN;

    if (!This) {
        ERR("object is null \n");
        return E_POINTER;
    }

    *value = (IDispatch*) (This->pApplication);
    I_ApplicationExcel_AddRef( This->pApplication );

    if (!value) {
        ERR("value == NULL \n");
        return E_POINTER;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Parent(
        I_Workbooks* iface,
        IDispatch **value)
{
    TRACE_IN;
    return MSO_TO_OO_I_Workbooks_get_Application(iface, value);
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
/*TODO : use other parameters*/

    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    IUnknown *punk = NULL;
    HRESULT hres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(UpdateLinks, &UpdateLinks);
    MSO_TO_OO_CorrectArg(ReadOnly, &ReadOnly);
    MSO_TO_OO_CorrectArg(Format, &Format);
    MSO_TO_OO_CorrectArg(Password, &Password);
    MSO_TO_OO_CorrectArg(WriteResPassword, &WriteResPassword);
    MSO_TO_OO_CorrectArg(IgnoreReadOnlyRecommended, &IgnoreReadOnlyRecommended);
    MSO_TO_OO_CorrectArg(Origin, &Origin);
    MSO_TO_OO_CorrectArg(Delimiter, &Delimiter);
    MSO_TO_OO_CorrectArg(Editable, &Editable);
    MSO_TO_OO_CorrectArg(Notify, &Notify);
    MSO_TO_OO_CorrectArg(Converter, &Converter);
    MSO_TO_OO_CorrectArg(AddToMru, &AddToMru);
    MSO_TO_OO_CorrectArg(Local, &Local);
    MSO_TO_OO_CorrectArg(CorruptLoad, &CorruptLoad);

    if (!This) {
        ERR("object is null \n");
        return E_POINTER;
    }

    hres = _I_WorkbookConstructor((LPVOID*) &punk);
    if (FAILED(hres)) {
        ERR("constructor failed \n ");
        return E_NOINTERFACE;
    }
    
    if (This->capasity_workbooks == 0){
        This->count_workbooks = 1;
        This->capasity_workbooks = This->count_workbooks;
        This->pworkbook = HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, sizeof(WorkbookImpl*));
        if (!(This->pworkbook)) {
            ERR("Out of memory \n");
            return E_OUTOFMEMORY;
        }
        This->current_workbook = 0;
    } else {
        This->count_workbooks += 1;
        if (This->count_workbooks > This->capasity_workbooks) {
            This->capasity_workbooks = This->count_workbooks;
            This->pworkbook = HeapReAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, This->pworkbook, This->capasity_workbooks * sizeof(WorkbookImpl*));
        }
        if (!(This->pworkbook)) {
            ERR("Out of memory \n");
            return E_OUTOFMEMORY;
        }
        This->current_workbook = This->count_workbooks - 1;
    }

    hres = I_Workbook_QueryInterface(punk, &IID_I_Workbook, (void**) &(This->pworkbook[This->current_workbook]));
/*    I_Workbook_Release(punk);*/
    if (FAILED(hres)) return E_FAIL;
    *ppWorkbook = (IDispatch*)(This->pworkbook[This->current_workbook]);

    /* Prepare file name to right view 
    From:   C:\test test.xls
    To:     file:///c:/test%20test.xls */
    BSTR filenametmp;
    MSO_TO_OO_MakeURLFromFilename(Filename, &filenametmp);
    /*Initialize*/
    WTRACE(L"FILENAME ------>  %s \n",filenametmp);
    VARIANT_BOOL astemp;
    astemp = VARIANT_FALSE;

    hres = MSO_TO_OO_I_Workbook_Initialize2(This->pworkbook[This->current_workbook], iface, filenametmp, astemp);
    if (FAILED(hres)) {
        *ppWorkbook = NULL;
        ERR("Workbook_Initialize2 failed \n");
        return hres;
    }
    SysFreeString(filenametmp);
    I_Workbook_AddRef(*ppWorkbook);

    TRACE_OUT;
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
    TRACE_NOTIMPL;
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
    TRACE_NOTIMPL;
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks__OpenXML(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT Stylesheets,
        long Lcid,
        IDispatch **ppWorkbook)
{
    TRACE_NOTIMPL;
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_CanCheckOut(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT_BOOL *result)
{
    TRACE_IN;
    /*всегда возвращаем - нельзя получить*/
    *result = VARIANT_FALSE;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_put_CanCheckOut(
        I_Workbooks* iface,
        BSTR Filename,
        VARIANT_BOOL result)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_CheckOut(
        I_Workbooks* iface,
        BSTR Filename)
{
    TRACE_NOTIMPL;
    /*Возвращаем S_OK на всякий случай*/
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Creator(
        I_Workbooks* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get__Default(
        I_Workbooks* iface,
        VARIANT varIndex,
        IDispatch **ppSheet)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);    
    HRESULT hres;
    VARIANT i4_index;
    TRACE_IN;
    
    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;           
    }
    
    if (!ppSheet) {
         ERR("ppSheet is NULL \n");
         return E_POINTER;
    }
    
    MSO_TO_OO_CorrectArg(varIndex, &varIndex);
    VariantInit(&i4_index);
    
    if (V_VT(&varIndex) == VT_BSTR) {
        ERR(" BSTR parameters not supported \n");                
    } else {
        hres = VariantChangeTypeEx(&i4_index, &varIndex,0,0,VT_I4);
        if (FAILED(hres)) {
            ERR(" ERROR when VariantChangeTypeEx\n");
            return E_FAIL;
        }       
           
        if (V_I4(&i4_index) > This->count_workbooks) {
            ERR("Index i4_index = %i,  count = %i", V_I4(&i4_index), This->count_workbooks);  
            return E_FAIL;               
        } 
        
        *ppSheet = (IDispatch*)(This->pworkbook[V_I4(&i4_index) - 1]);
        I_Workbook_AddRef((I_Workbook*)(*ppSheet));  
        return S_OK;
    }
    
    ERR(" parameters \n");
    
    TRACE_OUT;
    return E_FAIL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get_Item(
        I_Workbooks* iface,
        VARIANT index,
        IDispatch **result)
{
    TRACE("-------> get__Default \n");
    TRACE_IN;
    TRACE_OUT;
    return I_Workbooks_get__Default(iface, index, result);
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_get__NewEnum(
        I_Workbooks* iface,
        IUnknown **RHS)
{
    TRACE_IN;
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);    
    *RHS = (IUnknown*)WORKBOOKS_ENUM(This);
    IUnknown_AddRef(*RHS);
    TRACE_OUT;
    return S_OK;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbooks_GetTypeInfoCount(
        I_Workbooks* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_GetTypeInfo(
        I_Workbooks* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_workbooks(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_GetIDsOfNames(
        I_Workbooks* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_workbooks(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }

    TRACE_OUT;
    return hres;
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
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    HRESULT hres;
    int iresult;
    IDispatch* iapp;
    VARIANT_BOOL vbool;
    VARIANT vresult;
    VARIANT astemp;
    VARIANT vnull;
    long l;
    ITypeInfo *typeinfo;

    TRACE("\n");

    VariantInit(&vnull);
    VariantInit(&astemp);
    VariantInit(&vresult);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;           
    }

    switch (dispIdMember) {
        case dispid_workbooks__Open:
            /*Зависит от кол-ва посланных параметров*/
            if (pDispParams->cArgs==0) return E_FAIL;
            /*Используем только имя файла*/
            VariantInit(&vnull);
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-1], &astemp);
            hres = I_Workbooks_Open(iface, V_BSTR(&astemp), vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, l, &iapp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)iapp;
            } else {
                IDispatch_Release(iapp);
            }
            return S_OK;
        case dispid_workbooks_Close:
            I_Workbooks_Close(iface, l);
            return S_OK;
        case dispid_workbooks_Open:
            /*Зависит от кол-ва посланных параметров*/
            if (pDispParams->cArgs==0) return E_FAIL;
            /*Используем только имя файла*/
            VariantInit(&vnull);
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-1], &astemp);
            hres = I_Workbooks_Open(iface, V_BSTR(&astemp), vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull, vnull,l, &iapp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)iapp;
            } else {
                IDispatch_Release(iapp);
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
        case dispid_workbooks___OpenText:
            /*MSO_TO_OO_I_Workbooks___OpenText*/
            TRACE("Stub: METHOD __OpenText \n");
            return S_OK;
        default:
            hres = get_typeinfo_workbooks(&typeinfo);
            if (FAILED(hres))
                return hres;

            hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
            if (FAILED(hres)) {
                ERR("wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
            }

            return hres;
    }

    return E_NOTIMPL;
}

#undef WORKBOOKS_THIS

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


/*IEnumVARIANT interface*/

#define ENUMVAR_THIS(iface) DEFINE_THIS(WorkbooksImpl, enumerator, iface);

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Workbooks_EnumVAR_AddRef(
        IEnumVARIANT* iface)
{
    WorkbooksImpl *This = ENUMVAR_THIS(iface);
    return I_Workbooks_AddRef(WORKBOOKS_WORKBOOKS(This));
}


static HRESULT WINAPI MSO_TO_OO_I_Workbooks_EnumVAR_QueryInterface(
        IEnumVARIANT* iface,
        REFIID riid,
        void **ppvObject)
{
    WorkbooksImpl *This = ENUMVAR_THIS(iface);
    return I_Workbooks_QueryInterface(WORKBOOKS_WORKBOOKS(This), riid, ppvObject);
}

static ULONG WINAPI MSO_TO_OO_I_Workbooks_EnumVAR_Release(
        IEnumVARIANT* iface)
{
    WorkbooksImpl *This = ENUMVAR_THIS(iface);
    return I_Workbooks_Release(WORKBOOKS_WORKBOOKS(This));
}

/*** IEnumVARIANT methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Workbooks_EnumVAR_Next(
        IEnumVARIANT* iface,
        ULONG celt,
        VARIANT *rgVar,
        ULONG *pCeltFetched)
{
    WorkbooksImpl *This = ENUMVAR_THIS(iface);
    HRESULT hres;
    ULONG l;
    long l1;
    int count;
    ULONG l2;
    IDispatch *dret;
    VARIANT varindex, vNull;

    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;

    if (This->enum_position<0)
        return S_FALSE;

    if (pCeltFetched != NULL)
       *pCeltFetched = 0;

    if (rgVar == NULL)
       return E_INVALIDARG;

    VariantInit(&varindex);
    /*Init Array*/
    for (l=0; l<celt; l++)
       VariantInit(&rgVar[l]);

    I_Workbooks_get_Count(WORKBOOKS_WORKBOOKS(This), &count);
    V_VT(&varindex) = VT_I4;

    for (l1=This->enum_position, l2=0; l1<count && l2<celt; l1++, l2++) {
      V_I4(&varindex) = l1 + 1; //Because index of workbook start from 1
      hres = I_Workbooks_get_Item(WORKBOOKS_WORKBOOKS(This), varindex, &dret);
      V_VT(&rgVar[l2]) = VT_DISPATCH;
      V_DISPATCH(&rgVar[l2]) = dret;
      if (FAILED(hres))
         goto error;
    }

    if (pCeltFetched != NULL)
       *pCeltFetched = l2;

   This->enum_position = l1;

   return  (l2 < celt) ? S_FALSE : S_OK;

error:
   for (l=0; l<celt; l++)
      VariantClear(&rgVar[l]);
   VariantClear(&varindex);
   return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_EnumVAR_Skip(
        IEnumVARIANT* iface,
        ULONG celt)
{
    WorkbooksImpl *This = ENUMVAR_THIS(iface);
    int count;
    TRACE_IN;

    I_Workbooks_get_Count(WORKBOOKS_WORKBOOKS(This), &count);
    This->enum_position += celt;

    if (This->enum_position>=(count)) {
        This->enum_position = count - 1;
        TRACE_OUT;
        return S_FALSE;
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_EnumVAR_Reset(
        IEnumVARIANT* iface)
{
    WorkbooksImpl *This = ENUMVAR_THIS(iface);
    This->enum_position = 0;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Workbooks_EnumVAR_Clone(
        IEnumVARIANT* iface,
        IEnumVARIANT **ppEnum)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

#undef ENUMVAR_THIS

const IEnumVARIANTVtbl MSO_TO_OO_Workbooks_enumvarVtbl =
{
    MSO_TO_OO_I_Workbooks_EnumVAR_QueryInterface,
    MSO_TO_OO_I_Workbooks_EnumVAR_AddRef,
    MSO_TO_OO_I_Workbooks_EnumVAR_Release,
    MSO_TO_OO_I_Workbooks_EnumVAR_Next,
    MSO_TO_OO_I_Workbooks_EnumVAR_Skip,
    MSO_TO_OO_I_Workbooks_EnumVAR_Reset,
    MSO_TO_OO_I_Workbooks_EnumVAR_Clone
};

extern HRESULT _I_WorkbooksConstructor(LPVOID *ppObj)
{
    WorkbooksImpl *workbooks;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    workbooks = HeapAlloc(GetProcessHeap(), 0, sizeof(*workbooks));
    if (!workbooks) {
        return E_OUTOFMEMORY;
    }

    workbooks->pworkbooksVtbl = &MSO_TO_OO_I_WorkbooksVtbl;
    workbooks->penumeratorVtbl = &MSO_TO_OO_Workbooks_enumvarVtbl;
    workbooks->ref = 0;
    workbooks->pApplication = NULL;
    workbooks->count_workbooks = 0;
    workbooks->pworkbook = NULL;
    workbooks->current_workbook = -1;
    workbooks->capasity_workbooks = 0;
    workbooks->enum_position = 0;

    *ppObj = WORKBOOKS_WORKBOOKS(workbooks);
    
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}

