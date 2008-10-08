/*
 * Register and unregister mso_to_oo.dll functions
 *
 * Copyright (C) 2002 John K. Hohm
 * Copyright (C) 2007 Roy Shea (Google)
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

static ULONG WINAPI MSO_TO_OO_IClassFactory_AddRef(LPCLASSFACTORY iface)
{
    ClassFactoryImpl *This = (ClassFactoryImpl *)iface;
    ULONG ref;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    TRACE_OUT;
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_IClassFactory_QueryInterface(
        LPCLASSFACTORY iface,
        REFIID riid,
        LPVOID *ppvObj)
{
    ClassFactoryImpl *This = (ClassFactoryImpl *)iface;
    TRACE_IN;

    if (This == NULL || ppvObj == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_IClassFactory)) {
        *ppvObj = (LPVOID)iface;
        MSO_TO_OO_IClassFactory_AddRef(iface);
        TRACE_OUT;
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_IClassFactory_Release(LPCLASSFACTORY iface)
{
    ClassFactoryImpl *This = (ClassFactoryImpl *)iface;
    ULONG ref;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        HeapFree(GetProcessHeap(), 0, This);
        InterlockedDecrement(&dll_ref);
    }
    TRACE_OUT;
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_IClassFactory_CreateInstance(
        LPCLASSFACTORY iface,
        LPUNKNOWN pUnkOuter,
        REFIID riid,
        LPVOID *ppvObj)
{
    ClassFactoryImpl *This = (ClassFactoryImpl *)iface;
    HRESULT res;
    IUnknown *punk = NULL;
    TRACE_IN;

    if (This == NULL || ppvObj == NULL) return E_POINTER;

    if (pUnkOuter != NULL) return CLASS_E_NOAGGREGATION;

    res = _ApplicationExcelConstructor((LPVOID*) &punk);
    if (FAILED(res)) {
        TRACE("ERROR when _ApplicationExcelConstructor \n");
        return res;
    }

    res = I_ApplicationExcel_QueryInterface(punk, riid, ppvObj);
    if (FAILED(res)) {
        TRACE("ERROR when ApplicationExcel_QueryInterface \n");
        return res;
    }
    TRACE_OUT;
    return res;
}

static HRESULT WINAPI MSO_TO_OO_IClassFactory_LockServer(
        LPCLASSFACTORY iface,
        BOOL fLock)
{
    TRACE_IN;

    if (fLock != FALSE) {
        IClassFactory_AddRef(iface);
    } else {
        IClassFactory_Release(iface);
    }
    TRACE_OUT;
    return S_OK;
}

static const IClassFactoryVtbl IClassFactory_Vtbl =
{
    MSO_TO_OO_IClassFactory_QueryInterface,
    MSO_TO_OO_IClassFactory_AddRef,
    MSO_TO_OO_IClassFactory_Release,
    MSO_TO_OO_IClassFactory_CreateInstance,
    MSO_TO_OO_IClassFactory_LockServer
};

ClassFactoryImpl OOFFICE_ClassFactory =
{
    &IClassFactory_Vtbl,
    0
};

