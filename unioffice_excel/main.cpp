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

#include "unioffice_excel_private.h"
#include "factory.h"

LONG g_cServerLocks = 0;
LONG g_cComponents  = 0;

LONG OOVersion = VER_3;
BOOL write_log = 0;
char buf[MAX_PATH+50];
FILE *trace_file;

#ifdef DEBUG
int __tab = 0;
#endif

#ifdef __cplusplus
extern "C"
{
#endif

__declspec(dllexport) BOOL __stdcall DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    TRACE_IN;
    TRACE("(%p, %d, %p)\n", hinstDLL, fdwReason, lpvReserved);
    switch (fdwReason)
    {
        case DLL_PROCESS_ATTACH:
            DisableThreadLibraryCalls(hinstDLL);
            break;
        case DLL_PROCESS_DETACH:
            break;
    }
    TRACE_OUT;
    return TRUE;
}

__declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID iid, LPVOID *ppv)
{
    *ppv = NULL;
    char file_name[]= {'\\','u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','l','o','g',0};
    int len,i=0;
    TRACE_IN;

    if (IsEqualGUID(rclsid, CLSID_Application)) {
        /*îÁÞÉÎÁÅÍ ÚÁÐÉÓØ ÌÏÇÁ ÅÓÌÉ ÆÁÊÌ ÓÕÝÅÓÔ×ÕÅÔ*/
        len = GetSystemDirectoryA(buf, MAX_PATH);
        if (len) {
            while (file_name[i]!=0) {buf[len+i]=file_name[i];i++;};
            if (GetFileAttributesA(buf) != 0xFFFFFFFF) {
                write_log = 1;
                trace_file = fopen(buf,"w");
                if (trace_file) fclose(trace_file);
            }
        }
        
        // Ñîçäàòü ôàáðèêó êëàññà
        CFactory* pFactory = new CFactory; // Ñ÷åò÷èê ññûëîê óñòàíàâëèâàåòñÿ
        // â êîíñòðóêòîðå â 1
        if (pFactory == NULL)
        {
            TRACE_OUT;
            return E_OUTOFMEMORY;
        }
    
        // Ïîëó÷èòü òðåáóåìûé èíòåðôåéñ
        HRESULT hr = pFactory->QueryInterface(iid, ppv);
        pFactory->Release();
        
        TRACE_OUT;
        return hr;
    }
    TRACE_OUT;
    return CLASS_E_CLASSNOTAVAILABLE;
}

__declspec(dllexport) HRESULT __stdcall DllCanUnloadNow(void)
{
    TRACE("GLOBAL REF = %i \n", g_cComponents);
    TRACE("g_cServerLocks = %i \n", g_cServerLocks);
     
    if ((g_cComponents == 0) && (g_cServerLocks == 0))
    {
         return S_OK;
    }
    else
    {
         return S_FALSE;
    }

}

#ifdef __cplusplus
}
#endif
