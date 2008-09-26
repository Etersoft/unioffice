/*
 * Main module
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

#include "unioffice_word_private.h"

#ifdef __cplusplus
extern "C"
{
#endif

LONG dll_ref = 0;

__declspec(dllexport) BOOL __stdcall DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    switch (fdwReason)
    {
        case DLL_PROCESS_ATTACH:
            DisableThreadLibraryCalls(hinstDLL);
            break;
        case DLL_PROCESS_DETACH:
            break;
    }

    return TRUE;
}

__declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID iid, LPVOID *ppv)
{
    return E_NOTIMPL;
}

__declspec(dllexport) HRESULT __stdcall DllCanUnloadNow()
{
    return dll_ref != 0 ? S_FALSE : S_OK;
}

__declspec(dllexport) HRESULT __stdcall DllRegisterServer()
{
    return E_NOTIMPL;
}

__declspec(dllexport) HRESULT __stdcall DllUnRegisterServer()
{
    return  E_NOTIMPL;
}

#ifdef __cplusplus
}
#endif
