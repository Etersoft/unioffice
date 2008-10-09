/*
 * Names interface functions
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

ITypeInfo *ti_name = NULL;

HRESULT get_typeinfo_name(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};
    TRACE_IN;

    if (ti_name) {
        *typeinfo = ti_name;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_Name, &ti_name);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_name;
    TRACE_OUT;
    return hres;
}

/*Name interface*/
/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_Name_AddRef(
        Name* iface)
{
    NameImpl *This = (NameImpl*)iface;
    ULONG ref;
    TRACE_IN;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    TRACE_OUT;
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_Name_QueryInterface(
        Name* iface,
        REFIID riid,
        void **ppvObject)
{
    NameImpl *This = (NameImpl*)iface;
    TRACE_IN;

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_Name)) {
        *ppvObject = &This->nameVtbl;
        MSO_TO_OO_Name_AddRef(iface);
        TRACE_OUT;
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_Name_Release(
        Name* iface)
{
    NameImpl *This = (NameImpl*)iface;
    ULONG ref;
    TRACE_IN;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pApplication != NULL) {
            I_ApplicationExcel_Release(This->pApplication);
            This->pApplication = NULL;
        }
        if (This->pnames != NULL) {
            IDispatch_Release(This->pnames);
            This->pnames = NULL;
        }
        if (This->pOOName != NULL) {
            IDispatch_Release(This->pOOName);
            This->pOOName = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    TRACE_OUT;
    return ref;
}

/*** Name methods ***/
static HRESULT WINAPI MSO_TO_OO_Name_get_Application(
        Name* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_Creator(
        Name* iface,
        VARIANT *result)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_Parent(
        Name* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get__Default(
        Name* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_Index(
        Name* iface,
        int *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_Category(
        Name* iface,
        LCID lcid,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_Category(
        Name* iface,
        LCID lcid,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_CategoryLocal(
        Name* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_CategoryLocal(
        Name* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_Delete(
        Name* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_MacroType(
        Name* iface,
        XlXLMMacroType *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_MacroType(
        Name* iface,
        XlXLMMacroType value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_Name(
        Name* iface,
        LCID lcid,
        BSTR *value)
{
    NameImpl *This = (NameImpl*)iface;
    HRESULT hres;
    VARIANT vres;
    TRACE_IN;

    VariantInit(&vres);

    hres = AutoWrap(DISPATCH_METHOD, &vres, This->pOOName, L"getName", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when getName \n");
        return hres;
    }
    *value = SysAllocString(V_BSTR(&vres));
    VariantClear(&vres);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_Name(
        Name* iface,
        LCID lcid,
        BSTR value)
{
    NameImpl *This = (NameImpl*)iface;
    HRESULT hres;
    VARIANT vres, param1;
    TRACE_IN;

    VariantInit(&vres);
    VariantInit(&param1);

    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(value);

    hres = AutoWrap(DISPATCH_METHOD, &vres, This->pOOName, L"setName", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when setName \n");
    }
    VariantClear(&param1);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_RefersTo(
        Name* iface,
        LCID lcid,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_RefersTo(
        Name* iface,
        LCID lcid,
        VARIANT value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_ShortcutKey(
        Name* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_ShortcutKey(
        Name* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_Value(
        Name* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_Value(
        Name* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_Visible(
        Name* iface,
        VARIANT_BOOL *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_Visible(
        Name* iface,
        VARIANT_BOOL value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_NameLocal(
        Name* iface,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_NameLocal(
        Name* iface,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_RefersToLocal(
        Name* iface,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_RefersToLocal(
        Name* iface,
        VARIANT value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_RefersToR1C1(
        Name* iface,
        LCID lcid,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_RefersToR1C1(
        Name* iface,
        LCID lcid,
        VARIANT value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_RefersToR1C1Local(
        Name* iface,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_RefersToR1C1Local(
        Name* iface,
        VARIANT value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_RefersToRange(
        Name* iface,
        IDispatch **value)
{
    NameImpl *This = (NameImpl*)iface;
    NamesImpl *onames = (NamesImpl*)This->pnames;
    I_Sheets *shs;
    I_Worksheet *wsh;
    int i, count=0;
    VARIANT index,vNull, vname;
    BSTR tmpname;
    HRESULT hres;
    TRACE_IN;

    VariantInit(&index);
    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;
    VariantInit(&vname);

    hres = I_Workbook_get_Sheets((I_Workbook*)onames->pwb,(IDispatch**) &shs);
    if (FAILED(hres)) {
        TRACE("ERROR When get Sheets\n");
        I_Sheets_Release(shs);
        return E_FAIL;
    }

    I_Sheets_get_Count(shs, &count);
    if (FAILED(hres)) {
        TRACE("ERROR When get count\n");
        I_Sheets_Release(shs);
        return E_FAIL;
    }

    /*получаем имя объекта name*/
    hres = Name_get_Name(iface, 0, &tmpname);
    if (FAILED(hres)) {
        TRACE("ERROR When get name\n");
        I_Sheets_Release(shs);
        return E_FAIL;
    }
    V_VT(&vname) = VT_BSTR;
    V_BSTR(&vname) = SysAllocString(tmpname);
    SysFreeString(tmpname);


    V_VT(&index) = VT_I4;
    for (i=0;i<count;i++) {
         V_I4(&index) = i+1;
         hres = I_Sheets_get_Item(shs, index,(IDispatch**)&wsh);
         if (FAILED(hres)) {
             TRACE("ERROR When get Sheets\n");
             I_Sheets_Release(shs);
             VariantClear(&vname);
             return E_FAIL;
         }
         hres = I_Worksheet_get_Range(wsh, vname, vNull, value);
         I_Worksheet_Release(wsh);
         if (!FAILED(hres)) {
             I_Sheets_Release(shs);
             VariantClear(&vname);
             TRACE_OUT;
             return S_OK;
         }
    }

    I_Sheets_Release(shs);
    VariantClear(&vname);

    TRACE("NOT FIND\n");
    return E_FAIL;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_Name_GetTypeInfoCount(
        Name* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_GetTypeInfo(
        Name* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_GetIDsOfNames(
        Name* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_name(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_Name_Invoke(
        Name* iface,
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
    IDispatch *dret;
    BSTR tmpbstr;
    VARIANT vparam1;

    TRACE("\n");
    VariantInit(&vparam1);

    if (iface==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    switch(dispIdMember)
    {
    case dispid_name_referstorange:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE("NOT_IMPL \n");
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_Name_get_RefersToRange(iface, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case dispid_name_name:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs>1) {
                TRACE("ERROR parameter referstorange\n");
                return E_FAIL;
            }
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vparam1);
            return MSO_TO_OO_Name_put_Name(iface, 0, V_BSTR(&vparam1));
        } else {
            hres = MSO_TO_OO_Name_get_Name(iface, 0, &tmpbstr);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_BSTR;
                V_BSTR(pVarResult)=SysAllocString(tmpbstr);
            } else {
                IDispatch_Release(dret);
            }
            SysFreeString(tmpbstr);
            return S_OK;
        }
    }

    TRACE(" dispIdMember = %i NOT REALIZE\n",dispIdMember);
    return E_NOTIMPL;
}

const NameVtbl MSO_TO_OO_NameVtbl =
{
    MSO_TO_OO_Name_QueryInterface,
    MSO_TO_OO_Name_AddRef,
    MSO_TO_OO_Name_Release,
    MSO_TO_OO_Name_GetTypeInfoCount,
    MSO_TO_OO_Name_GetTypeInfo,
    MSO_TO_OO_Name_GetIDsOfNames,
    MSO_TO_OO_Name_Invoke,
    MSO_TO_OO_Name_get_Application,
    MSO_TO_OO_Name_get_Creator,
    MSO_TO_OO_Name_get_Parent,
    MSO_TO_OO_Name_get__Default,
    MSO_TO_OO_Name_get_Index,
    MSO_TO_OO_Name_get_Category,
    MSO_TO_OO_Name_put_Category,
    MSO_TO_OO_Name_get_CategoryLocal,
    MSO_TO_OO_Name_put_CategoryLocal,
    MSO_TO_OO_Name_Delete,
    MSO_TO_OO_Name_get_MacroType,
    MSO_TO_OO_Name_put_MacroType,
    MSO_TO_OO_Name_get_Name,
    MSO_TO_OO_Name_put_Name,
    MSO_TO_OO_Name_get_RefersTo,
    MSO_TO_OO_Name_put_RefersTo,
    MSO_TO_OO_Name_get_ShortcutKey,
    MSO_TO_OO_Name_put_ShortcutKey,
    MSO_TO_OO_Name_get_Value,
    MSO_TO_OO_Name_put_Value,
    MSO_TO_OO_Name_get_Visible,
    MSO_TO_OO_Name_put_Visible,
    MSO_TO_OO_Name_get_NameLocal,
    MSO_TO_OO_Name_put_NameLocal,
    MSO_TO_OO_Name_get_RefersToLocal,
    MSO_TO_OO_Name_put_RefersToLocal,
    MSO_TO_OO_Name_get_RefersToR1C1,
    MSO_TO_OO_Name_put_RefersToR1C1,
    MSO_TO_OO_Name_get_RefersToR1C1Local,
    MSO_TO_OO_Name_put_RefersToR1C1Local,
    MSO_TO_OO_Name_get_RefersToRange
};

extern HRESULT _NameConstructor(LPVOID *ppObj)
{
    NameImpl *name;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    name = HeapAlloc(GetProcessHeap(), 0, sizeof(*name));
    if (!name)
    {
        return E_OUTOFMEMORY;
    }

    name->nameVtbl = &MSO_TO_OO_NameVtbl;
    name->ref = 0;
    name->pApplication = NULL;
    name->pnames = NULL;
    name->pOOName = NULL;

    *ppObj = &name->nameVtbl;
    TRACE_OUT;
    return S_OK;
}

/*Names interface*/

ITypeInfo *ti_names = NULL;

HRESULT get_typeinfo_names(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};
    TRACE_IN;
    if (ti_names) {
        *typeinfo = ti_names;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_Names, &ti_names);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_names;
    TRACE_OUT;
    return hres;
}


/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_Names_AddRef(
        Names* iface)
{
    NamesImpl *This = (NamesImpl*)iface;
    ULONG ref;
    TRACE_IN;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    TRACE_OUT;
    return ref;
}


static HRESULT WINAPI MSO_TO_OO_Names_QueryInterface(
        Names* iface,
        REFIID riid,
        void **ppvObject)
{
    NamesImpl *This = (NamesImpl*)iface;
    TRACE_IN;

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_Names)) {
        *ppvObject = &This->namesVtbl;
        MSO_TO_OO_Names_AddRef(iface);
        TRACE_OUT;
        return S_OK;
    }

    return E_NOINTERFACE;
}


static ULONG WINAPI MSO_TO_OO_Names_Release(
        Names* iface)
{
    NamesImpl *This = (NamesImpl*)iface;
    ULONG ref;
    TRACE_IN;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pApplication != NULL) {
            I_ApplicationExcel_Release(This->pApplication);
            This->pApplication = NULL;
        }
        if (This->pwb != NULL) {
            IDispatch_Release(This->pwb);
            This->pwb = NULL;
        }
        if (This->pOONames != NULL) {
            IDispatch_Release(This->pOONames);
            This->pOONames = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    TRACE_OUT;
    return ref;
}

/*** Names methods ***/
static HRESULT WINAPI MSO_TO_OO_Names_get_Application(
        Names* iface,
        IDispatch **value)
{
    NamesImpl *This = (NamesImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (value==NULL)
        return E_POINTER;

    *value = This->pApplication;
    IDispatch_AddRef(This->pApplication);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names_get_Count(
        Names* iface,
        int *count)
{
    NamesImpl *This = (NamesImpl*)iface;
    VARIANT vret;
    HRESULT hres;
    TRACE_IN;

    VariantInit(&vret);

    hres = AutoWrap(DISPATCH_METHOD, &vret, This->pOONames, L"getCount", 0);
    if (FAILED(hres)) {
        TRACE("Error when getCount \n");
        return E_FAIL;
    }

    hres = VariantChangeTypeEx(&vret, &vret, 0, 0, VT_I2);
    if (FAILED(hres)) {
        TRACE("Error when VariantChangeTypeEx \n");
        return E_FAIL;
    }

    *count = V_I2(&vret);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names_get_Creator(
        Names* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names_get_Parent(
        Names* iface,
        IDispatch **value)
{
    NamesImpl *This = (NamesImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (value==NULL)
        return E_POINTER;

    *value = This->pwb;
    IDispatch_AddRef(This->pwb);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names__Default(
        Names* iface,
        VARIANT Index,
        VARIANT IndexLocal,
        VARIANT RefersTo,
        IDispatch **ppvalue)
{
    /*Используем пока только первый параметр*/
    NamesImpl *This = (NamesImpl*)iface;
    HRESULT hres;
    IUnknown *punk = NULL;
    IDispatch *pname;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *ppvalue = NULL;

    hres = _NameConstructor((LPVOID*) &punk);
    if (FAILED(hres)) {
        TRACE("ERROR when call constructor \n");
        return E_NOINTERFACE;
    }

    hres = Name_QueryInterface(punk, &IID_Name, (void**) &pname);
    if (pname == NULL) {
        return E_FAIL;
    }

    if (V_VT(&Index)==VT_BSTR) {
        hres = MSO_TO_OO_Name_Initialize_By_Name((Name*)pname, iface, Index);
        if (FAILED(hres)) {
            IDispatch_Release(pname);
            return hres;
        }
        *ppvalue = pname;
        return S_OK;
    } else {
        if (Is_Variant_Null(Index)) {
            TRACE("ERROR Empty param \n ");
            return E_FAIL;
        } else {
            /*доступ по индексу*/
            TRACE("NOT REALIZED \n");
            return E_FAIL;
        }
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names_Add(
        Names* iface,
        VARIANT Name,
        VARIANT RefersTo,
        VARIANT Visible,
        VARIANT MacroType,
        VARIANT ShortcutKey,
        VARIANT Category,
        VARIANT NameLocal,
        VARIANT RefersToLocal,
        VARIANT CategoryLocal,
        VARIANT RefersToR1C1,
        VARIANT RefersToR1C1Local,
        IDispatch **ppvalue)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Names_GetEnumerator(
        Names* iface,
        IDispatch **value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Names_Item(
        Names* iface,
        VARIANT Index,
        VARIANT IndexLocal,
        VARIANT RefersTo,
        IDispatch **ppvalue)
{
    TRACE("\n");
    return MSO_TO_OO_Names__Default(iface, Index, IndexLocal, RefersTo, ppvalue);
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_Names_GetTypeInfoCount(
        Names* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Names_GetTypeInfo(
        Names* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Names_GetIDsOfNames(
        Names* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_names(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

#define param_count 3
static HRESULT WINAPI MSO_TO_OO_Names_Invoke(
        Names* iface,
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

    hres = get_typeinfo_names(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
    if (FAILED(hres)) {
        TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
    }
    TRACE_OUT;
    return hres;
}

const NamesVtbl MSO_TO_OO_NamesVtbl =
{
    MSO_TO_OO_Names_QueryInterface,
    MSO_TO_OO_Names_AddRef,
    MSO_TO_OO_Names_Release,
    MSO_TO_OO_Names_GetTypeInfoCount,
    MSO_TO_OO_Names_GetTypeInfo,
    MSO_TO_OO_Names_GetIDsOfNames,
    MSO_TO_OO_Names_Invoke,
    MSO_TO_OO_Names_get_Application,
    MSO_TO_OO_Names_get_Creator,
    MSO_TO_OO_Names_get_Parent,
    MSO_TO_OO_Names_Add,
    MSO_TO_OO_Names_Item,
    MSO_TO_OO_Names__Default,
    MSO_TO_OO_Names_get_Count,
    MSO_TO_OO_Names_GetEnumerator
};

extern HRESULT _NamesConstructor(LPVOID *ppObj)
{
    NamesImpl *names;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    names = HeapAlloc(GetProcessHeap(), 0, sizeof(*names));
    if (!names)
    {
        return E_OUTOFMEMORY;
    }

    names->namesVtbl = &MSO_TO_OO_NamesVtbl;
    names->ref = 0;
    names->pApplication = NULL;
    names->pwb = NULL;
    names->pOONames = NULL;

    *ppObj = &names->namesVtbl;
    TRACE_IN;
    return S_OK;
}
