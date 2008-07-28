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

static WCHAR const str_application[] = {
    'A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR const str_parent[] = {
    'P','a','r','e','n','t',0};
static WCHAR const str_creator[] = {
    'C','r','e','a','t','o','r',0};
static WCHAR const str_add[] = {
    'A','d','d',0};
static WCHAR const str_item[] = {
    'I','t','e','m',0};
static WCHAR const str_count[] = {
    'C','o','u','n','t',0};
static WCHAR const str__default[] = {
    '_','D','e','f','a','u','l','t',0};
static WCHAR const str_getenumerator[] = {
    'G','e','t','E','n','u','m','e','r','a','t','o','r',0};


/*Name interface*/
/*** IUnknown methods ***/
static HRESULT WINAPI MSO_TO_OO_Name_QueryInterface(
        Name* iface,
        REFIID riid,
        void **ppvObject)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static ULONG WINAPI MSO_TO_OO_Name_AddRef(
        Name* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static ULONG WINAPI MSO_TO_OO_Name_Release(
        Name* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
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
        long lcid,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_Category(
        Name* iface,
        long lcid,
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
        long lcid,
        BSTR *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_Name(
        Name* iface,
        long lcid,
        BSTR value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_get_RefersTo(
        Name* iface,
        long lcid,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_RefersTo(
        Name* iface,
        long lcid,
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
        long lcid,
        VARIANT *value)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_Name_put_RefersToR1C1(
        Name* iface,
        long lcid,
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
    TRACE("\n");
    return E_NOTIMPL;
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
    TRACE("\n");
    return E_NOTIMPL;
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
    TRACE("\n");
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

    return S_OK;
}

/*Names interface*/
/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_Names_AddRef(
        Names* iface)
{
    NamesImpl *This = (NamesImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}


static HRESULT WINAPI MSO_TO_OO_Names_QueryInterface(
        Names* iface,
        REFIID riid,
        void **ppvObject)
{
    NamesImpl *This = (NamesImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_Names)) {
        *ppvObject = &This->namesVtbl;
        MSO_TO_OO_Names_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}


static ULONG WINAPI MSO_TO_OO_Names_Release(
        Names* iface)
{
    NamesImpl *This = (NamesImpl*)iface;
    ULONG ref;

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
    return ref;
}

/*** Names methods ***/
static HRESULT WINAPI MSO_TO_OO_Names_get_Application(
        Names* iface,
        IDispatch **value)
{
    NamesImpl *This = (NamesImpl*)iface;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if (value==NULL)
        return E_POINTER;

    *value = This->pApplication;
    IDispatch_AddRef(This->pApplication);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names_get_Count(
        Names* iface,
        int *count)
{
    NamesImpl *This = (NamesImpl*)iface;
    VARIANT vret;
    HRESULT hres;

    TRACE("\n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names_get_Creator(
        Names* iface,
        VARIANT *result)
{
    TRACE("\n");
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names_get_Parent(
        Names* iface,
        IDispatch **value)
{
    NamesImpl *This = (NamesImpl*)iface;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if (value==NULL)
        return E_POINTER;

    *value = This->pwb;
    IDispatch_AddRef(This->pwb);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_Names__Default(
        Names* iface,
        VARIANT Index,
        VARIANT IndexLocal,
        VARIANT RefersTo,
        IDispatch **ppvalue)
{
    TRACE("\n");
    return E_NOTIMPL;
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

}

static HRESULT WINAPI MSO_TO_OO_Names_Item(
        Names* iface,
        VARIANT Index,
        VARIANT IndexLocal,
        VARIANT RefersTo,
        IDispatch **ppvalue)
{
    TRACE("\n");
    return E_NOTIMPL;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_Names_GetTypeInfoCount(
        Names* iface,
        UINT *pctinfo)
{

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
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = dispid_names_application;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_creator)) {
        *rgDispId = dispid_names_creator;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_parent)) {
        *rgDispId = dispid_names_parent;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_add)) {
        *rgDispId = dispid_names_add;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_item)) {
        *rgDispId = dispid_names_item;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str__default)) {
        *rgDispId = dispid_names__default;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_count)) {
        *rgDispId = dispid_names_count;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_getenumerator)) {
        *rgDispId = dispid_names_getenumerator;
        return S_OK;
    }

    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L" %s NOT REALIZE \n",*rgszNames);
    return E_NOTIMPL;
}

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
    HRESULT hres;
    IDispatch *dret;
    VARIANT vresult;
    int iret;

    TRACE("\n");

    VariantInit(&vresult);

    if (iface == NULL) {
        TRACE("ERROR Object is NULL\n");
       return E_POINTER;
    }

    switch(dispIdMember)
    {
    case dispid_names_application:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_Names_get_Application(iface,&dret);
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
    case dispid_names_creator:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_Names_get_Creator(iface, &vresult);
            if (pVarResult!=NULL){
                *pVarResult = vresult;
            }
            return hres;
        }
    case dispid_names_parent:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_Names_get_Parent(iface,&dret);
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
    case dispid_names_add:
        return E_NOTIMPL;
    case dispid_names_item:
        return E_NOTIMPL;
    case dispid_names__default:
        return E_NOTIMPL;
    case dispid_names_count:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_Names_get_Count(iface, &iret);
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I2;
                V_I2(pVarResult) = iret;
            }
            return hres;
        }
    case dispid_names_getenumerator:
        return E_NOTIMPL;
    }
    TRACE("Unknown dispid =  \n",dispIdMember);
    return E_NOTIMPL;
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

    return S_OK;
}
