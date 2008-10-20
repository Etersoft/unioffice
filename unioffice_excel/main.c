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

#include "mso_to_oo_private.h"

LONG dll_ref = 0;
LONG OOVersion = VER_3;
BOOL write_log = 0;
char buf[MAX_PATH+50];
FILE *trace_file;

extern ITypeInfo *ti_excel;
extern ITypeInfo *ti_font;
extern ITypeInfo *ti_workbooks;
extern ITypeInfo *ti_workbook;
extern ITypeInfo *ti_sheets;
extern ITypeInfo *ti_worksheet;
extern ITypeInfo *ti_range;
extern ITypeInfo *ti_interrior;
extern ITypeInfo *ti_pagesetup;
extern ITypeInfo *ti_borders;
extern ITypeInfo *ti_border;
extern ITypeInfo *ti_name;
extern ITypeInfo *ti_names;
extern ITypeInfo *ti_outline;
extern ITypeInfo *ti_window;
extern ITypeInfo *ti_windows;
extern ITypeInfo *ti_shapes;
extern ITypeInfo *ti_shape;

__declspec(dllexport) BOOL __stdcall DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    TRACE_IN;
    TRACE("(%p, %d, %p)\n", hinstDLL, fdwReason, lpvReserved);
    fprintf(stderr,"WE ARE THERE \n");
    switch (fdwReason)
    {
        case DLL_PROCESS_ATTACH:
            DisableThreadLibraryCalls(hinstDLL);
            break;
        case DLL_PROCESS_DETACH:
            if(ti_excel)
                ti_excel->lpVtbl->Release(ti_excel);
            if(ti_font)
                ti_excel->lpVtbl->Release(ti_font);
            if(ti_workbooks)
                ti_excel->lpVtbl->Release(ti_workbooks);
            if(ti_workbook)
                ti_excel->lpVtbl->Release(ti_workbook);
            if(ti_sheets)
                ti_excel->lpVtbl->Release(ti_sheets);
            if(ti_worksheet)
                ti_excel->lpVtbl->Release(ti_worksheet);
            if(ti_range)
                ti_excel->lpVtbl->Release(ti_range);
            if(ti_interrior)
                ti_excel->lpVtbl->Release(ti_interrior);
            if(ti_pagesetup)
                ti_excel->lpVtbl->Release(ti_pagesetup);
            if(ti_borders)
                ti_excel->lpVtbl->Release(ti_borders);
            if(ti_border)
                ti_excel->lpVtbl->Release(ti_border);
            if(ti_name)
                ti_excel->lpVtbl->Release(ti_name);
            if(ti_names)
                ti_excel->lpVtbl->Release(ti_names);
            if(ti_outline)
                ti_excel->lpVtbl->Release(ti_outline);
            if(ti_window)
                ti_excel->lpVtbl->Release(ti_window);
            if(ti_windows)
                ti_excel->lpVtbl->Release(ti_windows);
            if(ti_shapes)
                ti_excel->lpVtbl->Release(ti_shapes);
            if(ti_shape)
                ti_excel->lpVtbl->Release(ti_shape);
            break;
    }
    TRACE_OUT;
    return TRUE;
}

__declspec(dllexport) STDAPI DllGetClassObject(REFCLSID rclsid, REFIID iid, LPVOID *ppv)
{
    *ppv = NULL;
    char file_name[]= {'\\','u','n','i','o','f','f','i','c','e','.','l','o','g',0};
    int len,i=0;
    TRACE_IN;

    if (IsEqualGUID(rclsid, &CLSID__ApplicationExcel)) {
        /*Начинаем запись лога если файл существует*/
        len = GetSystemDirectoryA(buf, MAX_PATH);
        if (len) {
            while (file_name[i]!=0) {buf[len+i]=file_name[i];i++;};
            if (GetFileAttributesA(buf) != 0xFFFFFFFF) {
                write_log = 1;
                trace_file = fopen(buf,"w");
                if (trace_file) fclose(trace_file);
            }
        }
        TRACE_OUT;
        return IClassFactory_QueryInterface((LPCLASSFACTORY)&OOFFICE_ClassFactory, iid, ppv);
    }
    TRACE_OUT;
    return CLASS_E_CLASSNOTAVAILABLE;
}

__declspec(dllexport) STDAPI DllCanUnloadNow(void)
{
    TRACE("GLOBAL REF = %i \n",dll_ref);
    return dll_ref != 0 ? S_FALSE : S_OK;
}

