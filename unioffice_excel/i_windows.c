/*
 * IRange interface functions
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
#include "special_functions.h"

/*IWindows interface*/
    /*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Windows_AddRef(
        I_Windows* iface)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_QueryInterface(
        I_Windows* iface,
        REFIID riid,
        void **ppvObject)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static ULONG WINAPI MSO_TO_OO_I_Windows_Release(
        I_Windows* iface)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

   /*** I_Windows methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Application(
        I_Windows* iface,
        IDispatch **value)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Creator(
        I_Windows* iface,
        XlCreator *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Parent(
        I_Windows* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_Arrange(
        I_Windows* iface,
        XlArrangeStyle ArrangeStyle,
        VARIANT ActiveWorkbook,
        VARIANT SyncHorizontal,
        VARIANT SyncVertical,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Count(
        I_Windows* iface,
        long *retval)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Item(
        I_Windows* iface,
        VARIANT Index,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get__NewEnum(
        I_Windows* iface,
        IUnknown **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get__Default(
        I_Windows* iface,
        VARIANT Index,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

    /*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Windows_GetTypeInfoCount(
        I_Windows* iface,
        UINT *pctinfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_GetTypeInfo(
        I_Windows* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_GetIDsOfNames(
        I_Windows* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_Invoke(
        I_Windows* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

const I_WindowsVtbl MSO_TO_OO_I_WindowsVtbl =
{
    MSO_TO_OO_I_Windows_QueryInterface,
    MSO_TO_OO_I_Windows_AddRef,
    MSO_TO_OO_I_Windows_Release,
    MSO_TO_OO_I_Windows_GetTypeInfoCount,
    MSO_TO_OO_I_Windows_GetTypeInfo,
    MSO_TO_OO_I_Windows_GetIDsOfNames,
    MSO_TO_OO_I_Windows_Invoke,
    MSO_TO_OO_I_Windows_get_Application,
    MSO_TO_OO_I_Windows_get_Creator,
    MSO_TO_OO_I_Windows_get_Parent,
    MSO_TO_OO_I_Windows_Arrange,
    MSO_TO_OO_I_Windows_get_Count,
    MSO_TO_OO_I_Windows_get_Item,
    MSO_TO_OO_I_Windows_get__NewEnum,
    MSO_TO_OO_I_Windows_get__Default
};

extern HRESULT _I_WindowsConstructor(LPVOID *ppObj)
{
    WindowsImpl *windows;

    TRACE("(%p)\n", ppObj);

    windows = HeapAlloc(GetProcessHeap(), 0, sizeof(*windows));
    if (!windows)
    {
        return E_OUTOFMEMORY;
    }

    windows->_windowsVtbl = &MSO_TO_OO_I_WindowsVtbl;
    windows->ref = 0;
    windows->pApplication = NULL;

    *ppObj = &windows->_windowsVtbl;

    return S_OK;
}

