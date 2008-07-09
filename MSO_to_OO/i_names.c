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

}

static HRESULT WINAPI MSO_TO_OO_Names_get_Count(
        Names* iface,
        int *count)
{

}

static HRESULT WINAPI MSO_TO_OO_Names_get_Creator(
        Names* iface,
        VARIANT *result)
{

}

static HRESULT WINAPI MSO_TO_OO_Names_get_Parent(
        Names* iface,
        IDispatch **value)
{

}

static HRESULT WINAPI MSO_TO_OO_Names__Default(
        Names* iface,
        VARIANT Index,
        VARIANT IndexLocal,
        VARIANT RefersTo,
        IDispatch **ppvalue)
{

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

}

static HRESULT WINAPI MSO_TO_OO_Names_GetIDsOfNames(
        Names* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{

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
    MSO_TO_OO_Names_get_Count,
    MSO_TO_OO_Names_get_Creator,
    MSO_TO_OO_Names_get_Parent,
    MSO_TO_OO_Names__Default,
    MSO_TO_OO_Names_Add,
    MSO_TO_OO_Names_GetEnumerator,
    MSO_TO_OO_Names_Item
};

NamesImpl MSO_TO_OO_Names =
{
    &MSO_TO_OO_NamesVtbl,
    0,
    NULL,
    NULL
};

extern HRESULT _NamesConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    NamesImpl *names;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

    names = HeapAlloc(GetProcessHeap(), 0, sizeof(*names));
    if (!names)
    {
        return E_OUTOFMEMORY;
    }

    names->namesVtbl = &MSO_TO_OO_NamesVtbl;
    names->ref = 0;
    names->pApplication = NULL;
    names->pwb = NULL;

    *ppObj = &names->namesVtbl;

    return S_OK;
}
