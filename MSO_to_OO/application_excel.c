/*
 * Excel.Application interface functions
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
#include <oleauto.h>

static WCHAR const str_usercontrol[] = {
    'U','s','e','r','C','o','n','t','r','o','l',0};
static WCHAR const str_displayalerts[] = {
    'D','i','s','p','l','a','y','A','l','e','r','t','s',0};
static WCHAR const str_windowstate[] = {
    'W','i','n','d','o','w','S','t','a','t','e',0};
static WCHAR const str_visible[] = {
    'V','i','s','i','b','l','e',0};
static WCHAR const str_workbooks[] = {
    'W','o','r','k','b','o','o','k','s',0};
static WCHAR const str_sheets[] = {
    'S','h','e','e','t','s',0};
static WCHAR const str_worksheets[] = {
    'W','o','r','k','s','h','e','e','t','s',0};
static WCHAR const str_cells[] = {
    'C','e','l','l','s',0};
static WCHAR const str_activesheet[] = {
    'A','c','t','i','v','e','S','h','e','e','t',0};
static WCHAR const str_version[] = {
    'V','e','r','s','i','o','n',0};
static WCHAR const str_convertformula[] = {
    'C','o','n','v','e','r','t','F','o','r','m','u','l','a',0};
static WCHAR const str_quit[] = {
    'Q','u','i','t',0};
static WCHAR const str_activecell[] = {
    'A','c','t','i','v','e','C','e','l','l',0};
static WCHAR const str_application[] = {
    'A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR const str_enableevents[] = {
    'E','n','a','b','l','e','E','v','e','n','t','s',0};
static WCHAR const str_screenupdating[] = {
    'S','c','r','e','e','n','U','p','d','a','t','i','n','g',0};
static WCHAR const str_caption[] = {
    'C','a','p','t','i','o','n',0};
static WCHAR const str_activeworkbook[] = {
    'A','c','t','i','v','e','W','o','r','k','b','o','o','k',0};
static WCHAR const str_range[] = {
    'R','a','n','g','e',0};
static WCHAR const str_columns[] = {
    'C','o','l','u','m','n','s',0};
static WCHAR const str_rows[] = {
    'R','o','w','s',0};
static WCHAR const str_selection[] = {
    'S','e','l','e','c','t','i','o','n',0};

#define DEFINE_THIS(class,ifild,iface) ((class*)((BYTE*)(iface)-offsetof(class,p ## ifild ## Vtbl)))


/*IConnectionPoint interface*/

#define CONPOINT_THIS(iface) DEFINE_THIS(_ApplicationExcelImpl,ConnectionPoint,iface);

    /*** IUnknown methods ***/
static HRESULT WINAPI MSO_TO_OO_ConnectionPoint_QueryInterface(
        IConnectionPoint* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationExcelImpl *This = CONPOINT_THIS(iface);
    return I_ApplicationExcel_QueryInterface(APPEXCEL(This), riid, ppvObject);
}

static ULONG WINAPI MSO_TO_OO_ConnectionPoint_AddRef(
        IConnectionPoint* iface)
{
    _ApplicationExcelImpl *This = CONPOINT_THIS(iface);
    return I_ApplicationExcel_AddRef(APPEXCEL(This));
}

static ULONG WINAPI MSO_TO_OO_ConnectionPoint_Release(
        IConnectionPoint* iface)
{
    _ApplicationExcelImpl *This = CONPOINT_THIS(iface);
    return I_ApplicationExcel_Release(APPEXCEL(This));
}

    /*** IConnectionPoint methods ***/
static HRESULT WINAPI MSO_TO_OO_ConnectionPoint_GetConnectionInterface(
        IConnectionPoint* iface,
        IID *pIID)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_ConnectionPoint_GetConnectionPointContainer(
        IConnectionPoint* iface,
        IConnectionPointContainer **ppCPC)
{
    _ApplicationExcelImpl *This = CONPOINT_THIS(iface);

    *ppCPC = (IConnectionPointContainer*)CONPOINTCONT(This);
    if (*ppCPC) {
        IConnectionPointContainer_AddRef(*ppCPC);
        return S_OK;
    }
    TRACE("ERROR \n");
    return E_FAIL;
}

static HRESULT WINAPI MSO_TO_OO_ConnectionPoint_Advise(
        IConnectionPoint* iface,
        IUnknown *pUnkSink,
        DWORD *pdwCookie)
{
    TRACE("Not implemented but return S_OK\n");
    *pdwCookie = 0;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_ConnectionPoint_Unadvise(
        IConnectionPoint* iface,
        DWORD dwCookie)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_ConnectionPoint_EnumConnections(
        IConnectionPoint* iface,
        IEnumConnections **ppEnum)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

const IConnectionPointVtbl MSO_TO_OO_ConnectionPointVtbl = 
{
    MSO_TO_OO_ConnectionPoint_QueryInterface,
    MSO_TO_OO_ConnectionPoint_AddRef,
    MSO_TO_OO_ConnectionPoint_Release,
    MSO_TO_OO_ConnectionPoint_GetConnectionInterface,
    MSO_TO_OO_ConnectionPoint_GetConnectionPointContainer,
    MSO_TO_OO_ConnectionPoint_Advise,
    MSO_TO_OO_ConnectionPoint_Unadvise,
    MSO_TO_OO_ConnectionPoint_EnumConnections
};

#undef CONPOINT_THIS

/*IConnectionPointContainer interface*/

#define CONPOINTCONT_THIS(iface) DEFINE_THIS(_ApplicationExcelImpl,ConnectionPointContainer,iface);

    /*** IUnknown methods ***/
static HRESULT WINAPI MSO_TO_OO_ConnectionPointContainer_QueryInterface(
        IConnectionPointContainer* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationExcelImpl *This = CONPOINTCONT_THIS(iface);
    return I_ApplicationExcel_QueryInterface(APPEXCEL(This), riid, ppvObject);
}

static ULONG WINAPI MSO_TO_OO_ConnectionPointContainer_AddRef(
        IConnectionPointContainer* iface)
{
    _ApplicationExcelImpl *This = CONPOINTCONT_THIS(iface);
    return I_ApplicationExcel_AddRef(APPEXCEL(This));
}

static ULONG WINAPI MSO_TO_OO_ConnectionPointContainer_Release(
        IConnectionPointContainer* iface)
{
    _ApplicationExcelImpl *This = CONPOINTCONT_THIS(iface);
    return I_ApplicationExcel_Release(APPEXCEL(This));
}

    /*** IConnectionPointContainer methods ***/
static HRESULT WINAPI MSO_TO_OO_ConnectionPointContainer_EnumConnectionPoints(
        IConnectionPointContainer* iface,
        IEnumConnectionPoints **ppEnum)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_ConnectionPointContainer_FindConnectionPoint(
        IConnectionPointContainer* iface,
        REFIID riid,
        IConnectionPoint **ppCP)
{
    _ApplicationExcelImpl *This = CONPOINTCONT_THIS(iface);
    WCHAR str_clsid[39];
    StringFromGUID2(riid, str_clsid, 39);
    WTRACE(L"riid = (%s) \n", str_clsid);

    *ppCP = (IConnectionPoint*)CONPOINT(This);
    if (*ppCP) {
        IConnectionPoint_AddRef(*ppCP);
        return S_OK;
    }
    TRACE("ERROR \n");
    return E_FAIL;
}

const IConnectionPointContainerVtbl MSO_TO_OO_ConnectionPointContainerVtbl = 
{
    MSO_TO_OO_ConnectionPointContainer_QueryInterface,
    MSO_TO_OO_ConnectionPointContainer_AddRef,
    MSO_TO_OO_ConnectionPointContainer_Release,
    MSO_TO_OO_ConnectionPointContainer_EnumConnectionPoints,
    MSO_TO_OO_ConnectionPointContainer_FindConnectionPoint
};

#undef CONPOINTCONT_THIS

/*IApplicationExcel interface*/
/*
IUnknown
*/

#define APPEXCEL_THIS(iface) DEFINE_THIS(_ApplicationExcelImpl, ApplicationExcel, iface);

static ULONG WINAPI MSO_TO_OO_I_ApplicationExcel_AddRef(
        I_ApplicationExcel* iface)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    ULONG ref;

    if (This == NULL) {
        TRACE("Object is NULL \n");
        return E_POINTER;
    }

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }

    TRACE("REF=%i \n", This->ref);

    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_QueryInterface(
        I_ApplicationExcel* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    WCHAR str_clsid[39];

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_ApplicationExcel)) {
        TRACE("IApplicationExcel \n");
        *ppvObject = APPEXCEL(This);
    }
    if (IsEqualGUID(riid, &IID_IConnectionPointContainer)) {
        TRACE("IConnectionPointContainer \n");
        *ppvObject = CONPOINTCONT(This);
    }

    if (*ppvObject) {
        I_ApplicationExcel_AddRef(iface);
        return S_OK;
    }

    StringFromGUID2(riid, str_clsid, 39);
    WTRACE(L"(%s) not supported \n", str_clsid);
    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_ApplicationExcel_Release(
        I_ApplicationExcel* iface)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    ULONG ref;

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);

    TRACE("REF = %i \n", This->ref);

    if (ref == 0) {
        if (This->pdOOApp != NULL) {
           IDispatch_Release(This->pdOOApp);
           This->pdOOApp = NULL;
        }
        if (This->pdWorkbooks != NULL) {
           IDispatch_Release(This->pdWorkbooks);
           This->pdWorkbooks = NULL;
        }
        if (This->pdOODesktop != NULL) {
           IDispatch_Release(This->pdOODesktop);
           This->pdOODesktop = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*
I_ApplicationExcel
*/
static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_UserControl(
        I_ApplicationExcel* iface,
        VARIANT_BOOL vbUserControl)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UserControl(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *vbUserControl)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayAlerts(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL vbDisplayAlerts)
{
   TRACE("\n");
   /*Возвращаем успех*/
   return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayAlerts(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *vbDisplayAlerts)
{
   TRACE("\n");
   /*Возвращаем успех*/
   *vbDisplayAlerts = VARIANT_FALSE;
   return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_WindowState(
        I_ApplicationExcel* iface,
        long lcid,
        XlWindowState State)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_WindowState(
        I_ApplicationExcel* iface,
        long lcid,
        XlWindowState *State)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Visible(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL vbVisible)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Visible(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *vbVisible)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Workbooks(
        I_ApplicationExcel* iface,
        IDispatch **ppWorkbooks)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);

    TRACE("\n");

    if (This->pdWorkbooks==NULL)
       return E_POINTER;

    *ppWorkbooks = This->pdWorkbooks;

    I_Workbooks_AddRef(This->pdWorkbooks);

    if (ppWorkbooks==NULL)
       return E_POINTER;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Sheets(
        I_ApplicationExcel* iface,
        IDispatch **ppSheets)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    HRESULT hres;

    TRACE("\n");

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) return E_FAIL;

    hres = I_Workbook_get_Sheets(pwb, ppSheets);
    if (FAILED(hres)) {
        I_Workbook_Release(pwb);
        return E_FAIL;
    }

    I_Workbook_Release(pwb);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Worksheets(
        I_ApplicationExcel* iface,
        IDispatch **ppSheets)
{
   /*Используем Sheets - они выполняют одинаковые функции*/
   TRACE("\n");
   return MSO_TO_OO_I_ApplicationExcel_get_Sheets(iface, ppSheets);
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Cells(
        I_ApplicationExcel* iface,
        IDispatch **ppRange)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    I_Sheets *pSheets;
    I_Worksheet *pworksheet;
    HRESULT hres;

    TRACE("\n");

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) return E_FAIL;

    hres = I_Workbook_get_Sheets(pwb, (IDispatch**) &pSheets);
    if (FAILED(hres)) {
        I_Workbook_Release(pwb);
        *ppRange = NULL;
        return E_FAIL;
    }

    hres = MSO_TO_OO_GetActiveSheet(pSheets, &pworksheet);
    if (FAILED(hres)) {
        I_Workbook_Release(pwb);
        I_Sheets_Release(pSheets);
        *ppRange = NULL;
        return hres;
    }

    hres = I_Worksheet_get_Cells(pworksheet, ppRange);
    if (FAILED(hres)) {
        *ppRange = NULL;
    }

    I_Workbook_Release(pwb);
    I_Sheets_Release(pSheets);
    I_Worksheet_Release(pworksheet);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    I_Sheets *pSheets;
    HRESULT hres;

    TRACE("\n");

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) return E_FAIL;

    hres = I_Workbook_get_Sheets(pwb, (IDispatch**) &pSheets);
    if (FAILED(hres)) {
        *RHS = NULL;
        I_Workbook_Release(pwb);
        return E_FAIL;
    }

    hres = MSO_TO_OO_GetActiveSheet(pSheets, (I_Worksheet**)RHS);
    if (FAILED(hres)) {
        I_Sheets_Release(pSheets);
        I_Workbook_Release(pwb);
        *RHS = NULL;
        return hres;
    }
    I_Sheets_Release(pSheets);
    I_Workbook_Release(pwb);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Version(
        I_ApplicationExcel* iface,
        long Lcid,
        BSTR *pVersion)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    if (pVersion == NULL)
        return E_POINTER;

    *pVersion = SysAllocString(OLESTR("11.0"));

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_ConvertFormula(
        I_ApplicationExcel* iface,
        VARIANT Formula,
        XlReferenceStyle FromReferenceStyle,
        VARIANT ToReferenceStyle,
        VARIANT ToAbsolute,
        VARIANT RelativeTo,
        long Lcid,
        VARIANT *pResult)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);

    TRACE("\n");

    if (This == NULL) return E_POINTER;
/*
RelativeTo и ToAbsolute - пока игнорируются
*/

/*Функция должна преобразовывать представление ячеек
1. из "=R1C1" в "=$A$1"
2. наоборот
3. предусмотреть вариант для области ячеек.
т.е. "=R1C1:R2C2" в "=$A$1:$B$2"
*/
    long tmp;
    VARIANT vtmp;
    VariantInit(&vtmp);
    WCHAR *result;
    WCHAR *sformula;
    WCHAR str[10];
    WCHAR stmp[2];
    int i;
    int row,col;

    /*преобразовываем любой тип к I4*/
    HRESULT hr = VariantChangeTypeEx(&vtmp, &ToReferenceStyle, 0, 0, VT_I4);
    if (FAILED(hr)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    tmp = V_I4(&vtmp);

    result = HeapAlloc(GetProcessHeap(),0,sizeof(WCHAR)*100);
    switch (FromReferenceStyle) {
    case xlA1:
        switch (tmp) {
        case xlR1C1:
            i=0;
            row = 0;
            col = 0;
            result[0]=0;
            stmp[1]=0;
            sformula = V_BSTR(&Formula);
            while (sformula[i]!=0) {
                if (sformula[i]==L'=') {
                    stmp[0]=sformula[i];
                    result=wcscat(result,stmp);
                }
                if ((sformula[i]>=L'A')&&(sformula[i]<=L'Z')) {
                   col=col*26+(sformula[i]-L'A')+1;
                }
                if ((sformula[i]>=L'0')&&(sformula[i]<=L'9')) {
                   row=row*10+(sformula[i]-L'0');
                }
                if (sformula[i]==L':') {
                    stmp[0]=L'R';
                    result=wcscat(result,stmp);
                    swprintf(str,L"%i",row);
                    result=wcscat(result,str);
                    stmp[0]=L'C';
                    result=wcscat(result,stmp);
                    swprintf(str,L"%i",col);
                    result=wcscat(result,str);
                    row = 0;
                    col = 0;
                    /*нужно пристыковать числа*/
                    stmp[0]=sformula[i];
                    result=wcscat(result,stmp);
                }
                i++;
            }
            stmp[0]=L'R';
            result=wcscat(result,stmp);
            swprintf(str,L"%i",row);
            result=wcscat(result,str);
            stmp[0]=L'C';
            result=wcscat(result,stmp);
            swprintf(str,L"%i",col);
            result=wcscat(result,str);
            V_VT(pResult) = VT_BSTR;
            V_BSTR(pResult) = SysAllocString(result);
            break;
        case xlA1:
            V_VT(pResult) = VT_BSTR;
            V_BSTR(pResult) = SysAllocString(V_BSTR(&Formula));
            break;
        }
        break;
    case xlR1C1:
        switch (tmp) {
        case xlA1:
            i=0;
            result[0]=0;
            stmp[1]=0;
            row = 0;
            sformula = V_BSTR(&Formula);
            while (sformula[i]!=0) {
                if (sformula[i]==L'=') {
                    stmp[0]=sformula[i];
                    result=wcscat(result,stmp);
                }
                if (sformula[i]==L'C') {
                   row = col;
                   col = 0;
                }
                if ((sformula[i]>=L'0')&&(sformula[i]<=L'9')) {
                   col=col*10+(sformula[i]-L'0');
                }
                if (sformula[i]==L':') {
                    if (col<26) {
                       stmp[0]=col + L'A' - 1;
                       result=wcscat(result,stmp);
                    } else {
                       stmp[0]=(col / 26) + L'A' - 1;
                       result=wcscat(result,stmp);
                       stmp[0]=(col % 26) + L'A' - 1;
                       result=wcscat(result,stmp);
                    }
                    swprintf(str,L"%i",row);
                    result=wcscat(result,str);
                    row = 0;
                    col = 0;
                    stmp[0]=sformula[i];
                    result=wcscat(result,stmp);
                }
                i++;
            }
            if (col<26) {
               stmp[0]=col + L'A' - 1;
               result=wcscat(result,stmp);
            } else {
               stmp[0]=(col / 26) + L'A' - 1;
               result=wcscat(result,stmp);
               stmp[0]=(col % 26) + L'A' - 1;
               result=wcscat(result,stmp);
               }
            swprintf(str,L"%i",row);
            result=wcscat(result,str);
            V_VT(pResult) = VT_BSTR;
            V_BSTR(pResult) = SysAllocString(result);
            break;
        case xlR1C1:
            V_VT(pResult) = VT_BSTR;
            V_BSTR(pResult) = SysAllocString(V_BSTR(&Formula));
            break;
        }
        break;
    }

    HeapFree(GetProcessHeap(),0,result);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Quit(
        I_ApplicationExcel* iface)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);

    TRACE("\n");

    if (iface==NULL) {
        TRACE("ERROR Object is NULL\n");
        return E_FAIL;
    }
    /*При вызове этого метода вызываем вызываем метод Close объекта WorkBooks*/
    I_Workbooks_Close((I_Workbooks*)(This->pdWorkbooks), 0);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveCell(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);

    HRESULT hres;

    TRACE("\n");

    hres = MSO_TO_OO_GetActiveCells((I_Workbooks*)This->pdWorkbooks, (I_Range**) RHS);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Application(
        I_ApplicationExcel* iface,
        IDispatch **value)
{

    TRACE(" \n");

    if (iface!=NULL) {
        *value = (IDispatch*)iface;
        MSO_TO_OO_I_ApplicationExcel_AddRef(iface);
    } else {
        return E_FAIL;
    }
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_EnableEvents(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *pvbee)
{
    TRACE("\n");
    /*Always return TRUE*/
    *pvbee = VARIANT_TRUE;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_EnableEvents(
        I_ApplicationExcel* iface,
        VARIANT_BOOL vbee)
{
    TRACE(" \n");
    /*Always return S_OK*/
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ScreenUpdating(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL vbscup)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ScreenUpdating(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *vbscup)
{
    TRACE("\n");
    /*Всегда возвращаем TRUE*/
    *vbscup = VARIANT_TRUE;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Caption(
        I_ApplicationExcel* iface,
        VARIANT *vName)
{
    TRACE("\n");
    if (vName==NULL) {
        TRACE("ERROR object is NULL\n");
    }
    V_VT(vName) = VT_BSTR;
    V_BSTR(vName) = SysAllocString(L"Microsoft Excel");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Caption(
        I_ApplicationExcel* iface,
        VARIANT vName)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveWorkbook(
        I_ApplicationExcel* iface,
        IDispatch **result)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    I_Sheets *pSheets;
    HRESULT hres;

    TRACE(" \n");
    if (This==NULL) return E_FAIL;

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) {
        TRACE("ERROR when GetActiveWorkbook\n");
        *result = NULL;
        return S_OK;
    }
    *result = (IDispatch*)pwb;

    I_Workbook_AddRef((I_Workbook*)*result);
    I_Workbook_Release(pwb);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Range(
        I_ApplicationExcel* iface,
        VARIANT Cell1,
        VARIANT Cell2,
        IDispatch **ppRange)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    HRESULT hres; 
    I_Worksheet *wsh;

    TRACE("\n");

    hres = MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(iface, (IDispatch**) &wsh);

    hres = I_Worksheet_get_Range(wsh,Cell1, Cell2, ppRange);

    I_Worksheet_Release(wsh);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Columns(
        I_ApplicationExcel* iface,
        VARIANT param,
        IDispatch **ppRange)
{
    HRESULT hres;
    IDispatch *active_sheet;

    hres = MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(iface, &active_sheet);

    if (FAILED(hres)) {
        TRACE("No active sheet \n");
        return E_FAIL;
    }

    hres = I_Worksheet_get_Columns((I_Worksheet*)active_sheet, param, ppRange);
    if (FAILED(hres)) {
        IDispatch_Release(active_sheet);
        TRACE("FAILED I_Worksheet_get_Columns \n");
        return hres;
    }
    IDispatch_Release(active_sheet);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Rows(
        I_ApplicationExcel* iface,
        VARIANT param,
        IDispatch **ppRange)
{
    HRESULT hres;
    IDispatch *active_sheet;

    hres = MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(iface, &active_sheet);

    if (FAILED(hres)) {
        TRACE("No active sheet \n");
        return E_FAIL;
    }

    hres = I_Worksheet_get_Rows((I_Worksheet*)active_sheet, param, ppRange);
    if (FAILED(hres)) {
        IDispatch_Release(active_sheet);
        TRACE("FAILED I_Worksheet_get_Rows \n");
        return hres;
    }
    IDispatch_Release(active_sheet);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Selection(
        I_ApplicationExcel* iface,
        IDispatch **ppRange)
{
    WorkbookImpl *awb;
    IDispatch *asheet;
    IUnknown *pobj;
    HRESULT hres;
    VARIANT vRes, vRet;
    RangeImpl *range;

    TRACE("\n");

    VariantInit(&vRes);
    VariantInit(&vRet);

    hres = MSO_TO_OO_I_ApplicationExcel_get_ActiveWorkbook(iface, (IDispatch**)&awb);
    if (FAILED(hres)) {
        TRACE("ERROR when get_ActiveWorkbook\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(iface, &asheet);
    if (FAILED(hres)) {
        I_Workbook_Release((I_Workbook*)awb);
        TRACE("ERROR when get_ActiveSheet\n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vRes, awb->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        I_Workbook_Release((I_Workbook*)awb);
        I_Worksheet_Release((I_Worksheet*)asheet);
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vRes), L"getSelection",0);
    if (FAILED(hres)) {
        I_Workbook_Release((I_Workbook*)awb);
        I_Worksheet_Release((I_Worksheet*)asheet);
        TRACE("ERROR when getSelectionr \n");
        return hres;
    }

    hres = _I_RangeConstructor((void**)&pobj);
    if (FAILED(hres)) {
        TRACE("ERROR when _I_RangeConstructor\n");
        VariantClear(&vRes);
        VariantClear(&vRet);
        I_Workbook_Release((I_Workbook*)awb);
        I_Worksheet_Release((I_Worksheet*)asheet);
        return E_FAIL;
    }

    hres = I_Range_QueryInterface(pobj, &IID_I_Range, (void**)ppRange);
    if (FAILED(hres)) {
        TRACE("ERROR when _I_RangeConstructor\n");
        VariantClear(&vRes);
        VariantClear(&vRet);
        I_Workbook_Release((I_Workbook*)awb);
        I_Worksheet_Release((I_Worksheet*)asheet);
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Range_Initialize3((I_Range*)*ppRange, V_DISPATCH(&vRet), asheet, (IDispatch*)iface);
    if (FAILED(hres)) {
        TRACE("ERROR when MSO_TO_OO_I_Range_Initialize2\n");
        VariantClear(&vRes);
        VariantClear(&vRet);
        I_Workbook_Release((I_Workbook*)awb);
        I_Worksheet_Release((I_Worksheet*)asheet);
        return E_FAIL;
    }

    VariantClear(&vRes);
    VariantClear(&vRet);
    I_Worksheet_Release((I_Worksheet*)asheet);
    I_Workbook_Release((I_Workbook*)awb);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Creator(
        I_ApplicationExcel* iface,
        XlCreator *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Parent(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveChart(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveDialog(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveMenuBar(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActivePrinter(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveWindow(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AddIns(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Assistant(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Calculate(
        I_ApplicationExcel* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Charts(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CommandBars(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DDEAppReturnCode(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DDEExecute(
        I_ApplicationExcel* iface,
        long Channel,
        BSTR String,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DDEInitiate(
        I_ApplicationExcel* iface,
        BSTR App,
        BSTR Topic,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DDEPoke(
        I_ApplicationExcel* iface,
        long Channel,
        VARIANT Item,
        VARIANT Data,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DDERequest(
        I_ApplicationExcel* iface,
        long Channel,
        BSTR Item,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DDETerminate(
        I_ApplicationExcel* iface,
        long Channel,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DialogSheets(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Evaluate(
        I_ApplicationExcel* iface,
        VARIANT Name,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel__Evaluate(
        I_ApplicationExcel* iface,
        VARIANT Name,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_ExecuteExcel4Macro(
        I_ApplicationExcel* iface,
        BSTR String,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Intersect(
        I_ApplicationExcel* iface,
        IDispatch *Arg1,
        IDispatch *Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MenuBars(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Modules(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Names(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Run(
        I_ApplicationExcel* iface,
        VARIANT Macro,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel__Run2(
        I_ApplicationExcel* iface,
        VARIANT Macro,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_SendKeys(
        I_ApplicationExcel* iface,
        VARIANT Keys,
        VARIANT Wait,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ShortcutMenus(
        I_ApplicationExcel* iface,
        long Index,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ThisWorkbook(
        I_ApplicationExcel* iface,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Toolbars(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Union(
        I_ApplicationExcel* iface,
        IDispatch *Arg1,
        IDispatch *Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Windows(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_WorksheetFunction(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Worksheets(
        I_ApplicationExcel* iface,
        IDispatch **ppSheets);

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Excel4IntlMacroSheets(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Excel4MacroSheets(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_ActivateMicrosoftApp(
        I_ApplicationExcel* iface,
        XlMSApplication Index,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_AddChartAutoFormat(
        I_ApplicationExcel* iface,
        VARIANT Chart,
        BSTR Name,
        VARIANT Description,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_AddCustomList(
        I_ApplicationExcel* iface,
        VARIANT ListArray,
        VARIANT ByRow,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AlertBeforeOverwriting(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_AlertBeforeOverwriting(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AltStartupPath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_AltStartupPath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AskToUpdateLinks(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_AskToUpdateLinks(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_EnableAnimations(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_EnableAnimations(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AutoCorrect(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Build(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CalculateBeforeSave(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_CalculateBeforeSave(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Calculation(
        I_ApplicationExcel* iface,
        long lcid,
        XlCalculation *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Calculation(
        I_ApplicationExcel* iface,
        long lcid,
        XlCalculation RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Caller(
        I_ApplicationExcel* iface,
        VARIANT Index,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CanPlaySounds(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CanRecordSounds(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CellDragAndDrop(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_CellDragAndDrop(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_CentimetersToPoints(
        I_ApplicationExcel* iface,
        double Centimeters,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_CheckSpelling(
        I_ApplicationExcel* iface,
        BSTR Word,
        VARIANT CustomDictionary,
        VARIANT IgnoreUppercase,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ClipboardFormats(
        I_ApplicationExcel* iface,
        VARIANT Index,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayClipboardWindow(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayClipboardWindow(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ColorButtons(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ColorButtons(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CommandUnderlines(
        I_ApplicationExcel* iface,
        long lcid,
        XlCommandUnderlines *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_CommandUnderlines(
        I_ApplicationExcel* iface,
        long lcid,
        XlCommandUnderlines RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ConstrainNumeric(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ConstrainNumeric(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CopyObjectsWithCells(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_CopyObjectsWithCells(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Cursor(
        I_ApplicationExcel* iface,
        long lcid,
        XlMousePointer *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Cursor(
        I_ApplicationExcel* iface,
        long lcid,
        XlMousePointer RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CustomListCount(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CutCopyMode(
        I_ApplicationExcel* iface,
        long lcid,
        XlCutCopyMode *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_CutCopyMode(
        I_ApplicationExcel* iface,
        long lcid,
        XlCutCopyMode RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}
static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DataEntryMode(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DataEntryMode(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy1(
        I_ApplicationExcel* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy2(
        I_ApplicationExcel* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy3(
        I_ApplicationExcel* iface,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy4(
        I_ApplicationExcel* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy5(
        I_ApplicationExcel* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy6(
        I_ApplicationExcel* iface,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy7(
        I_ApplicationExcel* iface,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy8(
        I_ApplicationExcel* iface,
        VARIANT Arg1,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy9(
        I_ApplicationExcel* iface,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy10(
        I_ApplicationExcel* iface,
        VARIANT arg,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy11(
        I_ApplicationExcel* This)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get__Default(
        I_ApplicationExcel* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DefaultFilePath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DefaultFilePath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DeleteChartAutoFormat(
        I_ApplicationExcel* iface,
        BSTR Name,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DeleteCustomList(
        I_ApplicationExcel* iface,
        long ListNum,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Dialogs(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayFormulaBar(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayFormulaBar(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayFullScreen(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayFullScreen(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayNoteIndicator(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayNoteIndicator(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayCommentIndicator(
        I_ApplicationExcel* iface,
        XlCommentDisplayMode *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayCommentIndicator(
        I_ApplicationExcel* iface,
        XlCommentDisplayMode RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayExcel4Menus(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayExcel4Menus(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayRecentFiles(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayRecentFiles(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayScrollBars(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayScrollBars(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayStatusBar(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayStatusBar(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DoubleClick(
        I_ApplicationExcel* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_EditDirectlyInCell(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_EditDirectlyInCell(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_EnableAutoComplete(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_EnableAutoComplete(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}
static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_EnableCancelKey(
        I_ApplicationExcel* iface,
        long lcid,
        XlEnableCancelKey *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_EnableCancelKey(
        I_ApplicationExcel* iface,
        long lcid,
        XlEnableCancelKey RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_EnableSound(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_EnableSound(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_EnableTipWizard(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_EnableTipWizard(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FileConverters(
        I_ApplicationExcel* iface,
        VARIANT Index1,
        VARIANT Index2,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FileSearch(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FileFind(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel__FindFile(
        I_ApplicationExcel* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FixedDecimal(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_FixedDecimal(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FixedDecimalPlaces(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_FixedDecimalPlaces(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetCustomListContents(
        I_ApplicationExcel* iface,
        long ListNum,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetCustomListNum(
        I_ApplicationExcel* iface,
        VARIANT ListArray,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetOpenFilename(
        I_ApplicationExcel* iface,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        VARIANT MultiSelect,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetSaveAsFilename(
        I_ApplicationExcel* iface,
        VARIANT InitialFilename,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Goto(
        I_ApplicationExcel* iface,
        VARIANT Reference,
        VARIANT Scroll,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Height(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Height(
        I_ApplicationExcel* iface,
        long lcid,
        double RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Help(
        I_ApplicationExcel* iface,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_IgnoreRemoteRequests(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_IgnoreRemoteRequests(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_InchesToPoints(
        I_ApplicationExcel* iface,
        double Inches,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_InputBox(
        I_ApplicationExcel* iface,
        BSTR Prompt,
        VARIANT Title,
        VARIANT Default,
        VARIANT Left,
        VARIANT Top,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        VARIANT Type,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Interactive(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Interactive(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_International(
        I_ApplicationExcel* iface,
        VARIANT Index,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Iteration(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Iteration(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_LargeButtons(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_LargeButtons(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Left(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Left(
        I_ApplicationExcel* iface,
        long lcid,
        double RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_LibraryPath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_MacroOptions(
        I_ApplicationExcel* iface,
        VARIANT Macro,
        VARIANT Description,
        VARIANT HasMenu,
        VARIANT MenuText,
        VARIANT HasShortcutKey,
        VARIANT ShortcutKey,
        VARIANT Category,
        VARIANT StatusBar,
        VARIANT HelpContextID,
        VARIANT HelpFile,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_MailLogoff(
        I_ApplicationExcel* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_MailLogon(
        I_ApplicationExcel* iface,
        VARIANT Name,
        VARIANT Password,
        VARIANT DownloadNewMail,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MailSession(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MailSystem(
        I_ApplicationExcel* iface,
        long lcid,
        XlMailSystem *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MathCoprocessorAvailable(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MaxChange(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_MaxChange(
        I_ApplicationExcel* iface,
        long lcid,
        double RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MaxIterations(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_MaxIterations(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MemoryFree(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MemoryTotal(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MemoryUsed(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MouseAvailable(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MoveAfterReturn(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_MoveAfterReturn(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MoveAfterReturnDirection(
        I_ApplicationExcel* iface,
        long lcid,
        XlDirection *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_MoveAfterReturnDirection(
        I_ApplicationExcel* iface,
        long lcid,
        XlDirection RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_RecentFiles(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Name(
        I_ApplicationExcel* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_NextLetter(
        I_ApplicationExcel* iface,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_NetworkTemplatesPath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ODBCErrors(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ODBCTimeout(
        I_ApplicationExcel* iface,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ODBCTimeout(
        I_ApplicationExcel* iface,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OnCalculate(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_OnCalculate(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OnData(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_OnData(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OnDoubleClick(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_OnDoubleClick(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OnEntry(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_OnEntry(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_OnKey(
        I_ApplicationExcel* iface,
        BSTR Key,
        VARIANT Procedure,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_OnRepeat(
        I_ApplicationExcel* iface,
        BSTR Text,
        BSTR Procedure,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OnSheetActivate(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_OnSheetActivate(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OnSheetDeactivate(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_OnSheetDeactivate(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_OnTime(
        I_ApplicationExcel* iface,
        VARIANT EarliestTime,
        BSTR Procedure,
        VARIANT LatestTime,
        VARIANT Schedule,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_OnUndo(
        I_ApplicationExcel* iface,
        BSTR Text,
        BSTR Procedure,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OnWindow(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_OnWindow(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OperatingSystem(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OrganizationName(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Path(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_PathSeparator(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_PreviousSelections(
        I_ApplicationExcel* iface,
        VARIANT Index,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_PivotTableSelection(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_PivotTableSelection(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_PromptForSummaryInfo(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_PromptForSummaryInfo(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_RecordMacro(
        I_ApplicationExcel* iface,
        VARIANT BasicCode,
        VARIANT XlmCode,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_RecordRelative(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ReferenceStyle(
        I_ApplicationExcel* iface,
        long lcid,
        XlReferenceStyle *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ReferenceStyle(
        I_ApplicationExcel* iface,
        long lcid,
        XlReferenceStyle RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_RegisteredFunctions(
        I_ApplicationExcel* iface,
        VARIANT Index1,
        VARIANT Index2,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_RegisterXLL(
        I_ApplicationExcel* iface,
        BSTR Filename,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Repeat(
        I_ApplicationExcel* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_ResetTipWizard(
        I_ApplicationExcel* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_RollZoom(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_RollZoom(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Save(
        I_ApplicationExcel* iface,
        VARIANT Filename,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_SaveWorkspace(
        I_ApplicationExcel* iface,
        VARIANT Filename,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_SetDefaultChart(
        I_ApplicationExcel* iface,
        VARIANT FormatName,
        VARIANT Gallery)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_SheetsInNewWorkbook(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_SheetsInNewWorkbook(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ShowChartTipNames(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ShowChartTipNames(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ShowChartTipValues(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ShowChartTipValues(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_StandardFont(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_StandardFont(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_StandardFontSize(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_StandardFontSize(
        I_ApplicationExcel* iface,
        long lcid,
        double RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_StartupPath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_StatusBar(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_StatusBar(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_TemplatesPath(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ShowToolTips(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ShowToolTips(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Top(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Top(
        I_ApplicationExcel* iface,
        long lcid,
        double RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DefaultSaveFormat(
        I_ApplicationExcel* iface,
        XlFileFormat *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DefaultSaveFormat(
        I_ApplicationExcel* iface,
        XlFileFormat RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_TransitionMenuKey(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_TransitionMenuKey(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_TransitionMenuKeyAction(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_TransitionMenuKeyAction(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_TransitionNavigKeys(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_TransitionNavigKeys(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Undo(
        I_ApplicationExcel* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UsableHeight(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UsableWidth(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UserName(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_UserName(
        I_ApplicationExcel* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Value(
        I_ApplicationExcel* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_VBE(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Volatile(
        I_ApplicationExcel* iface,
        VARIANT Volatile,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel__Wait(
        I_ApplicationExcel* iface,
        VARIANT Time,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Width(
        I_ApplicationExcel* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Width(
        I_ApplicationExcel* iface,
        long lcid,
        double RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_WindowsForPens(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UILanguage(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_UILanguage(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DefaultSheetDirection(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DefaultSheetDirection(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CursorMovement(
        I_ApplicationExcel* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_CursorMovement(
        I_ApplicationExcel* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ControlCharacters(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ControlCharacters(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel__WSFunction(
        I_ApplicationExcel* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayInfoWindow(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayInfoWindow(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Wait(
        I_ApplicationExcel* iface,
        VARIANT Time,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ExtendList(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ExtendList(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_OLEDBErrors(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetPhonetic(
        I_ApplicationExcel* iface,
        VARIANT Text,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_COMAddIns(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DefaultWebOptions(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ProductCode(
        I_ApplicationExcel* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UserLibraryPath(
        I_ApplicationExcel* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AutoPercentEntry(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_AutoPercentEntry(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_LanguageSettings(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Dummy101(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy12(
        I_ApplicationExcel* iface,
        IDispatch *p1,
        IDispatch *p2)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AnswerWizard(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_CalculateFull(
        I_ApplicationExcel* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_FindFile(
        I_ApplicationExcel* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CalculationVersion(
        I_ApplicationExcel* iface,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ShowWindowsInTaskbar(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ShowWindowsInTaskbar(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FeatureInstall(
        I_ApplicationExcel* iface,
        MsoFeatureInstall *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_FeatureInstall(
        I_ApplicationExcel* iface,
        MsoFeatureInstall RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Ready(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy13(
        I_ApplicationExcel* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FindFormat(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_putref_FindFormat(
        I_ApplicationExcel* iface,
        IDispatch *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ReplaceFormat(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_putref_ReplaceFormat(
        I_ApplicationExcel* iface,
        IDispatch *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UsedObjects(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CalculationState(
        I_ApplicationExcel* iface,
        XlCalculationState *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_CalculationInterruptKey(
        I_ApplicationExcel* iface,
        XlCalculationInterruptKey *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_CalculationInterruptKey(
        I_ApplicationExcel* iface,
        XlCalculationInterruptKey RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Watches(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayFunctionToolTips(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayFunctionToolTips(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AutomationSecurity(
        I_ApplicationExcel* iface,
        MsoAutomationSecurity *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_AutomationSecurity(
        I_ApplicationExcel* iface,
        MsoAutomationSecurity RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_FileDialog(
        I_ApplicationExcel* iface,
        MsoFileDialogType fileDialogType,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Dummy14(
        I_ApplicationExcel* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_CalculateFullRebuild(
        I_ApplicationExcel* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayPasteOptions(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayPasteOptions(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayInsertOptions(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
};

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayInsertOptions(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_GenerateGetPivotData(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_GenerateGetPivotData(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AutoRecover(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Hwnd(
        I_ApplicationExcel* iface,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Hinstance(
        I_ApplicationExcel* iface,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_CheckAbort(
        I_ApplicationExcel* iface,
        VARIANT KeepAbort)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ErrorCheckingOptions(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_AutoFormatAsYouTypeReplaceHyperlinks(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_AutoFormatAsYouTypeReplaceHyperlinks(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_SmartTagRecognizers(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_NewWorkbook(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_SpellingOptions(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Speech(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_MapPaperSize(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_MapPaperSize(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ShowStartupDialog(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ShowStartupDialog(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DecimalSeparator(
        I_ApplicationExcel* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DecimalSeparator(
        I_ApplicationExcel* iface,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ThousandsSeparator(
        I_ApplicationExcel* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_ThousandsSeparator(
        I_ApplicationExcel* iface,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_UseSystemSeparators(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_UseSystemSeparators(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ThisCell(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_RTD(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayDocumentActionTaskPane(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_DisplayDocumentActionTaskPane(
        I_ApplicationExcel* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_DisplayXMLSourcePane(
        I_ApplicationExcel* iface,
        VARIANT XmlMap)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ArbitraryXMLSupportAvailable(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Support(
        I_ApplicationExcel* iface,
        IDispatch *Object,
        long ID,
        VARIANT arg,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

/*
IDispatch
*/
static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetTypeInfoCount(
        I_ApplicationExcel* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetTypeInfo(
        I_ApplicationExcel* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_GetIDsOfNames(
        I_ApplicationExcel* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    if (!lstrcmpiW(*rgszNames, str_usercontrol)) {
        *rgDispId = dispid_application_usercontrol;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_displayalerts)) {
        *rgDispId = dispid_application_displayalerts;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_windowstate)) {
        *rgDispId = dispid_application_windowstate;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_visible)) {
        *rgDispId = dispid_application_visible;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_workbooks)) {
        *rgDispId = dispid_application_workbooks;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_sheets)) {
        *rgDispId = dispid_application_sheets;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_worksheets)) {
        *rgDispId = dispid_application_worksheets;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_cells)) {
        *rgDispId = dispid_application_cells;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_activesheet)) {
        *rgDispId = dispid_application_activesheet;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_version)) {
        *rgDispId = dispid_application_version;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_convertformula)) {
        *rgDispId = dispid_application_convertformula;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_quit)) {
        *rgDispId = dispid_application_quit;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_activecell)) {
        *rgDispId = dispid_application_activecell;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = dispid_application_application;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_enableevents)) {
        *rgDispId = dispid_application_enableevents;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_screenupdating)) {
        *rgDispId = dispid_application_screenupdating;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_caption)) {
        *rgDispId = dispid_application_caption;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_activeworkbook)) {
        *rgDispId = dispid_application_activeworkbook;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_range)) {
        *rgDispId = dispid_application_range;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_columns)) {
        *rgDispId = dispid_application_columns;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_rows)) {
        *rgDispId = dispid_application_rows;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_selection)) {
        *rgDispId = dispid_application_selection;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L" %s NOT REALIZE\n", *rgszNames);
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_Invoke(
        I_ApplicationExcel* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    _ApplicationExcelImpl *This = APPEXCEL_THIS(iface);
    HRESULT hr;
    BSTR pVersion;
    IDispatch *pdisp;
    IDispatch *pretdisp;
    long tmp;
    VARIANT vNull;
    VARIANT vRet,vtmp,cell1,cell2;
    VARIANT_BOOL vbin;

    VariantInit(&vNull);
    VariantInit(&vtmp);
    VariantInit(&cell1);
    VariantInit(&cell2);
    VariantInit(&vRet);

    TRACE(" %i\n", dispIdMember);

    if (This == NULL) return E_POINTER;

    switch (dispIdMember)
    {
    case dispid_application_usercontrol:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 1) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_ApplicationExcel_put_UserControl(iface,vbin);
        } else {
            return E_NOTIMPL;
        }
    case dispid_application_displayalerts:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 2) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_ApplicationExcel_put_DisplayAlerts(iface, 0, vbin);
        } else {
            return E_NOTIMPL;
        }
    case dispid_application_windowstate:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            /*преобразовываем любой тип к I4*/
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hr)) {
                TRACE(" (case 3) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            tmp = V_I4(&vtmp);
            MSO_TO_OO_I_ApplicationExcel_put_WindowState(iface, 0, tmp);
            return S_OK;
        } else {
            return E_NOTIMPL;
        }
    case dispid_application_visible:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 4) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_ApplicationExcel_put_Visible(iface, 0, vbin);
        } else {
            return E_NOTIMPL;
        }
    case dispid_application_workbooks:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_Workbooks(iface,&pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
            }
            return hr;
        }
    case dispid_application_sheets:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_Sheets(iface,&pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
               }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
            }
            if (pDispParams->cArgs==1) {
                I_Sheets_get__Default((I_Sheets*)pdisp, pDispParams->rgvarg[0], &pretdisp);
                I_Sheets_Release((I_Sheets*)pdisp);
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)pretdisp;
                } else {
                    I_Worksheet_Release((I_Worksheet*)pretdisp);
                }
            }
            return S_OK;
        }
    case dispid_application_worksheets:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_Sheets(iface,&pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
            }
            if (pDispParams->cArgs==1) {
                I_Sheets_get__Default((I_Sheets*)pdisp, pDispParams->rgvarg[0], &pretdisp);
                I_Sheets_Release((I_Sheets*)pdisp);
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)pretdisp;
                } else {
                    I_Worksheet_Release((I_Worksheet*)pretdisp);
                }
            }
            return S_OK;
        }
    case dispid_application_cells:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            switch (pDispParams->cArgs) {
            case 3:
                hr = MSO_TO_OO_I_ApplicationExcel_get_Cells(iface,&pdisp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                /*необходимо привести к значению , т.к. иногда присылаются ссылки*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &cell1);
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &cell2);
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
                I_Range_get__Default((I_Range*)pdisp,cell1, cell2,&pretdisp);
                I_Range_put_Value((I_Range*)pretdisp, vNull, 0, vtmp);
                IDispatch_Release(pdisp);
                IDispatch_Release(pretdisp);
                return S_OK;
            }
            TRACE(" (case 8) (PUT) only realized with 3 parameters \n");
            return E_NOTIMPL;
        } else {
            TRACE(" (case 8 (cells))\n");
            hr = MSO_TO_OO_I_ApplicationExcel_get_Cells(iface,&pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
	    /*здесь надо проверить параметры если они есть, то вызвать метод Range->_Default.*/
            switch(pDispParams->cArgs) {
            case 0:
                if (pVarResult!=NULL){
        	    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
                }
                break;
            case 1:
                if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1))) return E_FAIL;
                I_Range_get__Default((I_Range*)pdisp,cell1,vNull,&pretdisp);
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)pretdisp;
                } else {
                    I_Range_Release((I_Range*)pretdisp);
                }
                I_Range_Release((I_Range*)pdisp);
                break;
            case 2:
                if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1))) return E_FAIL;
                if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &cell2))) return E_FAIL;
                I_Range_get__Default((I_Range*)pdisp,cell2,cell1, &pretdisp);
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)pretdisp;
                } else {
                    I_Range_Release((I_Range*)pretdisp);
                }
                I_Range_Release((I_Range*)pdisp);
                break;
            }
            return hr;
        }
    case dispid_application_activesheet:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(iface, &pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
            } else {
                I_Worksheet_Release((I_Worksheet*)pdisp);
            }
            return S_OK;
        }
    case dispid_application_version:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_Version(iface,0,&pVersion);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_BSTR;
                V_BSTR(pVarResult)=pVersion;
            }
            return hr;
        }
    case dispid_application_convertformula:
        /*MSO_TO_OO_I_ApplicationExcel_ConvertFormula*/
        if (pDispParams->cArgs<3) return E_FAIL;

        /*преобразовываем любой тип к I4*/
        hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[pDispParams->cArgs-2]), 0, 0, VT_I4);
        if (FAILED(hr)) {
            TRACE(" (case 11) ERROR when VariantChangeTypeEx\n");
            return E_FAIL;
        }
        tmp = V_I4(&vtmp);

        hr = MSO_TO_OO_I_ApplicationExcel_ConvertFormula(iface, pDispParams->rgvarg[pDispParams->cArgs-1], tmp, pDispParams->rgvarg[pDispParams->cArgs-3], vNull, vNull, tmp, &vRet);
        if (FAILED(hr)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hr;
        }
        if (pVarResult!=NULL){
            V_VT(pVarResult)=VT_BSTR;
            V_BSTR(pVarResult)=V_BSTR(&vRet);
        }
        return S_OK;
    case dispid_application_quit:
        return MSO_TO_OO_I_ApplicationExcel_Quit(iface);
    case dispid_application_activecell:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_ActiveCell(iface, &pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
            } else {
                I_Range_Release((I_Range*)pdisp);
            }
            return hr;
        }
    case dispid_application_application:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_Application(iface,&pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
            }
            return S_OK;
        }
    case dispid_application_enableevents:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 15) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
           return MSO_TO_OO_I_ApplicationExcel_put_EnableEvents(iface, vbin);
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_EnableEvents(iface,&vbin);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = vbin;
            }
            return S_OK;
        }
    case dispid_application_screenupdating:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 16) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_ApplicationExcel_put_ScreenUpdating(iface, 0, vbin);
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_ScreenUpdating(iface, 0, &vbin);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = vbin;
            }
            return S_OK;
        }
    case dispid_application_caption:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BSTR);
            if (FAILED(hr)) {
                TRACE(" (case 17) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            hr = MSO_TO_OO_I_ApplicationExcel_put_Caption(iface, vtmp);
            return hr;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_Caption(iface, pVarResult);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hr;
        }
    case dispid_application_activeworkbook:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE(" (case 18) ERROR when (PUT)\n");
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_ActiveWorkbook(iface, &pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
            }
            return hr;
        }
    case dispid_application_range:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE(" (case 19) ERROR when (PUT)\n");
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 1:
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2);

                hr = MSO_TO_OO_I_ApplicationExcel_get_Range(iface, cell2, vNull, &pdisp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
                }
                return S_OK;
            case 2:
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &cell1);
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2);

                hr = MSO_TO_OO_I_ApplicationExcel_get_Range(iface, cell1, cell2, &pdisp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=(IDispatch *)pdisp;
                }
                return S_OK;
            default :
                TRACE(" (case 3) ERROR PARAMETR IS SEND\n");
                return E_FAIL;
            }
        }
    case dispid_application_columns:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE(" (case 20) ERROR when (PUT)\n");
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 0:
                TRACE("(case 20) 0 Parameter\n");
                hr = MSO_TO_OO_I_ApplicationExcel_get_Columns(iface, vNull, &pdisp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = pdisp;
                }
                return S_OK;
            case 1:
                TRACE("(case 20) 1 Parameter\n");
                /*Привести параметры к типу VARIANT если они переданы по ссылке*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1);
                hr = MSO_TO_OO_I_ApplicationExcel_get_Columns(iface, cell1, &pdisp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = pdisp;
                }
                return S_OK;
            }
        }
    case dispid_application_rows:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE(" (case 21) ERROR when (PUT)\n");
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 0:
                TRACE("(case 21) 0 Parameter\n");
                hr = MSO_TO_OO_I_ApplicationExcel_get_Rows(iface, vNull, &pdisp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = pdisp;
                }
                return S_OK;
            case 1:
                TRACE("(case 21) 1 Parameter\n");
                /*Привести параметры к типу VARIANT если они переданы по ссылке*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1);
                hr = MSO_TO_OO_I_ApplicationExcel_get_Rows(iface, cell1, &pdisp);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = pdisp;
                }
                return S_OK;
            }
        }
    case dispid_application_selection://Selection
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE("\n");
            return E_NOTIMPL;
        } else {
            TRACE("(case 22) 0 Parameter\n");
            hr = MSO_TO_OO_I_ApplicationExcel_get_Selection(iface, &pdisp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_DISPATCH;
                V_DISPATCH(pVarResult) = pdisp;
            }
            return S_OK;
        }
    }

    return E_NOTIMPL;
}

#undef APPEXCEL_THIS


const I_ApplicationExcelVtbl MSO_TO_OO_I_ApplicationExcel_Vtbl =
{
    MSO_TO_OO_I_ApplicationExcel_QueryInterface,
    MSO_TO_OO_I_ApplicationExcel_AddRef,
    MSO_TO_OO_I_ApplicationExcel_Release,
    MSO_TO_OO_I_ApplicationExcel_GetTypeInfoCount,
    MSO_TO_OO_I_ApplicationExcel_GetTypeInfo,
    MSO_TO_OO_I_ApplicationExcel_GetIDsOfNames,
    MSO_TO_OO_I_ApplicationExcel_Invoke,
    MSO_TO_OO_I_ApplicationExcel_get_Application,
    MSO_TO_OO_I_ApplicationExcel_get_Creator,
    MSO_TO_OO_I_ApplicationExcel_get_Parent,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveCell,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveChart,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveDialog,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveMenuBar,
    MSO_TO_OO_I_ApplicationExcel_get_ActivePrinter,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveWindow,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveWorkbook,
    MSO_TO_OO_I_ApplicationExcel_get_AddIns,
    MSO_TO_OO_I_ApplicationExcel_get_Assistant,
    MSO_TO_OO_I_ApplicationExcel_Calculate,
    MSO_TO_OO_I_ApplicationExcel_get_Cells,
    MSO_TO_OO_I_ApplicationExcel_get_Charts,
    MSO_TO_OO_I_ApplicationExcel_get_Columns,
    MSO_TO_OO_I_ApplicationExcel_get_CommandBars,
    MSO_TO_OO_I_ApplicationExcel_get_DDEAppReturnCode,
    MSO_TO_OO_I_ApplicationExcel_DDEExecute,
    MSO_TO_OO_I_ApplicationExcel_DDEInitiate,
    MSO_TO_OO_I_ApplicationExcel_DDEPoke,
    MSO_TO_OO_I_ApplicationExcel_DDERequest,
    MSO_TO_OO_I_ApplicationExcel_DDETerminate,
    MSO_TO_OO_I_ApplicationExcel_get_DialogSheets,
    MSO_TO_OO_I_ApplicationExcel_Evaluate,
    MSO_TO_OO_I_ApplicationExcel__Evaluate,
    MSO_TO_OO_I_ApplicationExcel_ExecuteExcel4Macro,
    MSO_TO_OO_I_ApplicationExcel_Intersect,
    MSO_TO_OO_I_ApplicationExcel_get_MenuBars,
    MSO_TO_OO_I_ApplicationExcel_get_Modules,
    MSO_TO_OO_I_ApplicationExcel_get_Names,
    MSO_TO_OO_I_ApplicationExcel_get_Range,
    MSO_TO_OO_I_ApplicationExcel_get_Rows,
    MSO_TO_OO_I_ApplicationExcel_Run,
    MSO_TO_OO_I_ApplicationExcel__Run2,
    MSO_TO_OO_I_ApplicationExcel_get_Selection,
    MSO_TO_OO_I_ApplicationExcel_SendKeys,
    MSO_TO_OO_I_ApplicationExcel_get_Sheets,
    MSO_TO_OO_I_ApplicationExcel_get_ShortcutMenus,
    MSO_TO_OO_I_ApplicationExcel_get_ThisWorkbook,
    MSO_TO_OO_I_ApplicationExcel_get_Toolbars,
    MSO_TO_OO_I_ApplicationExcel_Union,
    MSO_TO_OO_I_ApplicationExcel_get_Windows,
    MSO_TO_OO_I_ApplicationExcel_get_Workbooks,
    MSO_TO_OO_I_ApplicationExcel_get_WorksheetFunction,
    MSO_TO_OO_I_ApplicationExcel_get_Worksheets,
    MSO_TO_OO_I_ApplicationExcel_get_Excel4IntlMacroSheets,
    MSO_TO_OO_I_ApplicationExcel_get_Excel4MacroSheets,
    MSO_TO_OO_I_ApplicationExcel_ActivateMicrosoftApp,
    MSO_TO_OO_I_ApplicationExcel_AddChartAutoFormat,
    MSO_TO_OO_I_ApplicationExcel_AddCustomList,
    MSO_TO_OO_I_ApplicationExcel_get_AlertBeforeOverwriting,
    MSO_TO_OO_I_ApplicationExcel_put_AlertBeforeOverwriting,
    MSO_TO_OO_I_ApplicationExcel_get_AltStartupPath,
    MSO_TO_OO_I_ApplicationExcel_put_AltStartupPath,
    MSO_TO_OO_I_ApplicationExcel_get_AskToUpdateLinks,
    MSO_TO_OO_I_ApplicationExcel_put_AskToUpdateLinks,
    MSO_TO_OO_I_ApplicationExcel_get_EnableAnimations,
    MSO_TO_OO_I_ApplicationExcel_put_EnableAnimations,
    MSO_TO_OO_I_ApplicationExcel_get_AutoCorrect,
    MSO_TO_OO_I_ApplicationExcel_get_Build,
    MSO_TO_OO_I_ApplicationExcel_get_CalculateBeforeSave,
    MSO_TO_OO_I_ApplicationExcel_put_CalculateBeforeSave,
    MSO_TO_OO_I_ApplicationExcel_get_Calculation,
    MSO_TO_OO_I_ApplicationExcel_put_Calculation,
    MSO_TO_OO_I_ApplicationExcel_get_Caller,
    MSO_TO_OO_I_ApplicationExcel_get_CanPlaySounds,
    MSO_TO_OO_I_ApplicationExcel_get_CanRecordSounds,
    MSO_TO_OO_I_ApplicationExcel_get_Caption,
    MSO_TO_OO_I_ApplicationExcel_put_Caption,
    MSO_TO_OO_I_ApplicationExcel_get_CellDragAndDrop,
    MSO_TO_OO_I_ApplicationExcel_put_CellDragAndDrop,
    MSO_TO_OO_I_ApplicationExcel_CentimetersToPoints,
    MSO_TO_OO_I_ApplicationExcel_CheckSpelling,
    MSO_TO_OO_I_ApplicationExcel_get_ClipboardFormats,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayClipboardWindow,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayClipboardWindow,
    MSO_TO_OO_I_ApplicationExcel_get_ColorButtons,
    MSO_TO_OO_I_ApplicationExcel_put_ColorButtons,
    MSO_TO_OO_I_ApplicationExcel_get_CommandUnderlines,
    MSO_TO_OO_I_ApplicationExcel_put_CommandUnderlines,
    MSO_TO_OO_I_ApplicationExcel_get_ConstrainNumeric,
    MSO_TO_OO_I_ApplicationExcel_put_ConstrainNumeric,
    MSO_TO_OO_I_ApplicationExcel_ConvertFormula,
    MSO_TO_OO_I_ApplicationExcel_get_CopyObjectsWithCells,
    MSO_TO_OO_I_ApplicationExcel_put_CopyObjectsWithCells,
    MSO_TO_OO_I_ApplicationExcel_get_Cursor,
    MSO_TO_OO_I_ApplicationExcel_put_Cursor,
    MSO_TO_OO_I_ApplicationExcel_get_CustomListCount,
    MSO_TO_OO_I_ApplicationExcel_get_CutCopyMode,
    MSO_TO_OO_I_ApplicationExcel_put_CutCopyMode,
    MSO_TO_OO_I_ApplicationExcel_get_DataEntryMode,
    MSO_TO_OO_I_ApplicationExcel_put_DataEntryMode,
    MSO_TO_OO_I_ApplicationExcel_Dummy1,
    MSO_TO_OO_I_ApplicationExcel_Dummy2,
    MSO_TO_OO_I_ApplicationExcel_Dummy3,
    MSO_TO_OO_I_ApplicationExcel_Dummy4,
    MSO_TO_OO_I_ApplicationExcel_Dummy5,
    MSO_TO_OO_I_ApplicationExcel_Dummy6,
    MSO_TO_OO_I_ApplicationExcel_Dummy7,
    MSO_TO_OO_I_ApplicationExcel_Dummy8,
    MSO_TO_OO_I_ApplicationExcel_Dummy9,
    MSO_TO_OO_I_ApplicationExcel_Dummy10,
    MSO_TO_OO_I_ApplicationExcel_Dummy11,
    MSO_TO_OO_I_ApplicationExcel_get__Default,
    MSO_TO_OO_I_ApplicationExcel_get_DefaultFilePath,
    MSO_TO_OO_I_ApplicationExcel_put_DefaultFilePath,
    MSO_TO_OO_I_ApplicationExcel_DeleteChartAutoFormat,
    MSO_TO_OO_I_ApplicationExcel_DeleteCustomList,
    MSO_TO_OO_I_ApplicationExcel_get_Dialogs,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayAlerts,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayAlerts,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayFormulaBar,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayFormulaBar,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayFullScreen,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayFullScreen,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayNoteIndicator,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayNoteIndicator,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayCommentIndicator,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayCommentIndicator,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayExcel4Menus,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayExcel4Menus,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayRecentFiles,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayRecentFiles,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayScrollBars,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayScrollBars,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayStatusBar,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayStatusBar,
    MSO_TO_OO_I_ApplicationExcel_DoubleClick,
    MSO_TO_OO_I_ApplicationExcel_get_EditDirectlyInCell,
    MSO_TO_OO_I_ApplicationExcel_put_EditDirectlyInCell,
    MSO_TO_OO_I_ApplicationExcel_get_EnableAutoComplete,
    MSO_TO_OO_I_ApplicationExcel_put_EnableAutoComplete,
    MSO_TO_OO_I_ApplicationExcel_get_EnableCancelKey,
    MSO_TO_OO_I_ApplicationExcel_put_EnableCancelKey,
    MSO_TO_OO_I_ApplicationExcel_get_EnableSound,
    MSO_TO_OO_I_ApplicationExcel_put_EnableSound,
    MSO_TO_OO_I_ApplicationExcel_get_EnableTipWizard,
    MSO_TO_OO_I_ApplicationExcel_put_EnableTipWizard,
    MSO_TO_OO_I_ApplicationExcel_get_FileConverters,
    MSO_TO_OO_I_ApplicationExcel_get_FileSearch,
    MSO_TO_OO_I_ApplicationExcel_get_FileFind,
    MSO_TO_OO_I_ApplicationExcel__FindFile,
    MSO_TO_OO_I_ApplicationExcel_get_FixedDecimal,
    MSO_TO_OO_I_ApplicationExcel_put_FixedDecimal,
    MSO_TO_OO_I_ApplicationExcel_get_FixedDecimalPlaces,
    MSO_TO_OO_I_ApplicationExcel_put_FixedDecimalPlaces,
    MSO_TO_OO_I_ApplicationExcel_GetCustomListContents,
    MSO_TO_OO_I_ApplicationExcel_GetCustomListNum,
    MSO_TO_OO_I_ApplicationExcel_GetOpenFilename,
    MSO_TO_OO_I_ApplicationExcel_GetSaveAsFilename,
    MSO_TO_OO_I_ApplicationExcel_Goto,
    MSO_TO_OO_I_ApplicationExcel_get_Height,
    MSO_TO_OO_I_ApplicationExcel_put_Height,
    MSO_TO_OO_I_ApplicationExcel_Help,
    MSO_TO_OO_I_ApplicationExcel_get_IgnoreRemoteRequests,
    MSO_TO_OO_I_ApplicationExcel_put_IgnoreRemoteRequests,
    MSO_TO_OO_I_ApplicationExcel_InchesToPoints,
    MSO_TO_OO_I_ApplicationExcel_InputBox,
    MSO_TO_OO_I_ApplicationExcel_get_Interactive,
    MSO_TO_OO_I_ApplicationExcel_put_Interactive,
    MSO_TO_OO_I_ApplicationExcel_get_International,
    MSO_TO_OO_I_ApplicationExcel_get_Iteration,
    MSO_TO_OO_I_ApplicationExcel_put_Iteration,
    MSO_TO_OO_I_ApplicationExcel_get_LargeButtons,
    MSO_TO_OO_I_ApplicationExcel_put_LargeButtons,
    MSO_TO_OO_I_ApplicationExcel_get_Left,
    MSO_TO_OO_I_ApplicationExcel_put_Left,
    MSO_TO_OO_I_ApplicationExcel_get_LibraryPath,
    MSO_TO_OO_I_ApplicationExcel_MacroOptions,
    MSO_TO_OO_I_ApplicationExcel_MailLogoff,
    MSO_TO_OO_I_ApplicationExcel_MailLogon,
    MSO_TO_OO_I_ApplicationExcel_get_MailSession,
    MSO_TO_OO_I_ApplicationExcel_get_MailSystem,
    MSO_TO_OO_I_ApplicationExcel_get_MathCoprocessorAvailable,
    MSO_TO_OO_I_ApplicationExcel_get_MaxChange,
    MSO_TO_OO_I_ApplicationExcel_put_MaxChange,
    MSO_TO_OO_I_ApplicationExcel_get_MaxIterations,
    MSO_TO_OO_I_ApplicationExcel_put_MaxIterations,
    MSO_TO_OO_I_ApplicationExcel_get_MemoryFree,
    MSO_TO_OO_I_ApplicationExcel_get_MemoryTotal,
    MSO_TO_OO_I_ApplicationExcel_get_MemoryUsed,
    MSO_TO_OO_I_ApplicationExcel_get_MouseAvailable,
    MSO_TO_OO_I_ApplicationExcel_get_MoveAfterReturn,
    MSO_TO_OO_I_ApplicationExcel_put_MoveAfterReturn,
    MSO_TO_OO_I_ApplicationExcel_get_MoveAfterReturnDirection,
    MSO_TO_OO_I_ApplicationExcel_put_MoveAfterReturnDirection,
    MSO_TO_OO_I_ApplicationExcel_get_RecentFiles,
    MSO_TO_OO_I_ApplicationExcel_get_Name,
    MSO_TO_OO_I_ApplicationExcel_NextLetter,
    MSO_TO_OO_I_ApplicationExcel_get_NetworkTemplatesPath,
    MSO_TO_OO_I_ApplicationExcel_get_ODBCErrors,
    MSO_TO_OO_I_ApplicationExcel_get_ODBCTimeout,
    MSO_TO_OO_I_ApplicationExcel_put_ODBCTimeout,
    MSO_TO_OO_I_ApplicationExcel_get_OnCalculate,
    MSO_TO_OO_I_ApplicationExcel_put_OnCalculate,
    MSO_TO_OO_I_ApplicationExcel_get_OnData,
    MSO_TO_OO_I_ApplicationExcel_put_OnData,
    MSO_TO_OO_I_ApplicationExcel_get_OnDoubleClick,
    MSO_TO_OO_I_ApplicationExcel_put_OnDoubleClick,
    MSO_TO_OO_I_ApplicationExcel_get_OnEntry,
    MSO_TO_OO_I_ApplicationExcel_put_OnEntry,
    MSO_TO_OO_I_ApplicationExcel_OnKey,
    MSO_TO_OO_I_ApplicationExcel_OnRepeat,
    MSO_TO_OO_I_ApplicationExcel_get_OnSheetActivate,
    MSO_TO_OO_I_ApplicationExcel_put_OnSheetActivate,
    MSO_TO_OO_I_ApplicationExcel_get_OnSheetDeactivate,
    MSO_TO_OO_I_ApplicationExcel_put_OnSheetDeactivate,
    MSO_TO_OO_I_ApplicationExcel_OnTime,
    MSO_TO_OO_I_ApplicationExcel_OnUndo,
    MSO_TO_OO_I_ApplicationExcel_get_OnWindow,
    MSO_TO_OO_I_ApplicationExcel_put_OnWindow,
    MSO_TO_OO_I_ApplicationExcel_get_OperatingSystem,
    MSO_TO_OO_I_ApplicationExcel_get_OrganizationName,
    MSO_TO_OO_I_ApplicationExcel_get_Path,
    MSO_TO_OO_I_ApplicationExcel_get_PathSeparator,
    MSO_TO_OO_I_ApplicationExcel_get_PreviousSelections,
    MSO_TO_OO_I_ApplicationExcel_get_PivotTableSelection,
    MSO_TO_OO_I_ApplicationExcel_put_PivotTableSelection,
    MSO_TO_OO_I_ApplicationExcel_get_PromptForSummaryInfo,
    MSO_TO_OO_I_ApplicationExcel_put_PromptForSummaryInfo,
    MSO_TO_OO_I_ApplicationExcel_Quit,
    MSO_TO_OO_I_ApplicationExcel_RecordMacro,
    MSO_TO_OO_I_ApplicationExcel_get_RecordRelative,
    MSO_TO_OO_I_ApplicationExcel_get_ReferenceStyle,
    MSO_TO_OO_I_ApplicationExcel_put_ReferenceStyle,
    MSO_TO_OO_I_ApplicationExcel_get_RegisteredFunctions,
    MSO_TO_OO_I_ApplicationExcel_RegisterXLL,
    MSO_TO_OO_I_ApplicationExcel_Repeat,
    MSO_TO_OO_I_ApplicationExcel_ResetTipWizard,
    MSO_TO_OO_I_ApplicationExcel_get_RollZoom,
    MSO_TO_OO_I_ApplicationExcel_put_RollZoom,
    MSO_TO_OO_I_ApplicationExcel_Save,
    MSO_TO_OO_I_ApplicationExcel_SaveWorkspace,
    MSO_TO_OO_I_ApplicationExcel_get_ScreenUpdating,
    MSO_TO_OO_I_ApplicationExcel_put_ScreenUpdating,
    MSO_TO_OO_I_ApplicationExcel_SetDefaultChart,
    MSO_TO_OO_I_ApplicationExcel_get_SheetsInNewWorkbook,
    MSO_TO_OO_I_ApplicationExcel_put_SheetsInNewWorkbook,
    MSO_TO_OO_I_ApplicationExcel_get_ShowChartTipNames,
    MSO_TO_OO_I_ApplicationExcel_put_ShowChartTipNames,
    MSO_TO_OO_I_ApplicationExcel_get_ShowChartTipValues,
    MSO_TO_OO_I_ApplicationExcel_put_ShowChartTipValues,
    MSO_TO_OO_I_ApplicationExcel_get_StandardFont,
    MSO_TO_OO_I_ApplicationExcel_put_StandardFont,
    MSO_TO_OO_I_ApplicationExcel_get_StandardFontSize,
    MSO_TO_OO_I_ApplicationExcel_put_StandardFontSize,
    MSO_TO_OO_I_ApplicationExcel_get_StartupPath,
    MSO_TO_OO_I_ApplicationExcel_get_StatusBar,
    MSO_TO_OO_I_ApplicationExcel_put_StatusBar,
    MSO_TO_OO_I_ApplicationExcel_get_TemplatesPath,
    MSO_TO_OO_I_ApplicationExcel_get_ShowToolTips,
    MSO_TO_OO_I_ApplicationExcel_put_ShowToolTips,
    MSO_TO_OO_I_ApplicationExcel_get_Top,
    MSO_TO_OO_I_ApplicationExcel_put_Top,
    MSO_TO_OO_I_ApplicationExcel_get_DefaultSaveFormat,
    MSO_TO_OO_I_ApplicationExcel_put_DefaultSaveFormat,
    MSO_TO_OO_I_ApplicationExcel_get_TransitionMenuKey,
    MSO_TO_OO_I_ApplicationExcel_put_TransitionMenuKey,
    MSO_TO_OO_I_ApplicationExcel_get_TransitionMenuKeyAction,
    MSO_TO_OO_I_ApplicationExcel_put_TransitionMenuKeyAction,
    MSO_TO_OO_I_ApplicationExcel_get_TransitionNavigKeys,
    MSO_TO_OO_I_ApplicationExcel_put_TransitionNavigKeys,
    MSO_TO_OO_I_ApplicationExcel_Undo,
    MSO_TO_OO_I_ApplicationExcel_get_UsableHeight,
    MSO_TO_OO_I_ApplicationExcel_get_UsableWidth,
    MSO_TO_OO_I_ApplicationExcel_put_UserControl,
    MSO_TO_OO_I_ApplicationExcel_get_UserControl,
    MSO_TO_OO_I_ApplicationExcel_get_UserName,
    MSO_TO_OO_I_ApplicationExcel_put_UserName,
    MSO_TO_OO_I_ApplicationExcel_get_Value,
    MSO_TO_OO_I_ApplicationExcel_get_VBE,
    MSO_TO_OO_I_ApplicationExcel_get_Version,
    MSO_TO_OO_I_ApplicationExcel_put_Visible,
    MSO_TO_OO_I_ApplicationExcel_get_Visible,
    MSO_TO_OO_I_ApplicationExcel_Volatile,
    MSO_TO_OO_I_ApplicationExcel__Wait,
    MSO_TO_OO_I_ApplicationExcel_get_Width,
    MSO_TO_OO_I_ApplicationExcel_put_Width,
    MSO_TO_OO_I_ApplicationExcel_get_WindowsForPens,
    MSO_TO_OO_I_ApplicationExcel_put_WindowState,
    MSO_TO_OO_I_ApplicationExcel_get_WindowState,
    MSO_TO_OO_I_ApplicationExcel_get_UILanguage,
    MSO_TO_OO_I_ApplicationExcel_put_UILanguage,
    MSO_TO_OO_I_ApplicationExcel_get_DefaultSheetDirection,
    MSO_TO_OO_I_ApplicationExcel_put_DefaultSheetDirection,
    MSO_TO_OO_I_ApplicationExcel_get_CursorMovement,
    MSO_TO_OO_I_ApplicationExcel_put_CursorMovement,
    MSO_TO_OO_I_ApplicationExcel_get_ControlCharacters,
    MSO_TO_OO_I_ApplicationExcel_put_ControlCharacters,
    MSO_TO_OO_I_ApplicationExcel__WSFunction,
    MSO_TO_OO_I_ApplicationExcel_get_EnableEvents,
    MSO_TO_OO_I_ApplicationExcel_put_EnableEvents,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayInfoWindow,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayInfoWindow,
    MSO_TO_OO_I_ApplicationExcel_Wait,
    MSO_TO_OO_I_ApplicationExcel_get_ExtendList,
    MSO_TO_OO_I_ApplicationExcel_put_ExtendList,
    MSO_TO_OO_I_ApplicationExcel_get_OLEDBErrors,
    MSO_TO_OO_I_ApplicationExcel_GetPhonetic,
    MSO_TO_OO_I_ApplicationExcel_get_COMAddIns,
    MSO_TO_OO_I_ApplicationExcel_get_DefaultWebOptions,
    MSO_TO_OO_I_ApplicationExcel_get_ProductCode,
    MSO_TO_OO_I_ApplicationExcel_get_UserLibraryPath,
    MSO_TO_OO_I_ApplicationExcel_get_AutoPercentEntry,
    MSO_TO_OO_I_ApplicationExcel_put_AutoPercentEntry,
    MSO_TO_OO_I_ApplicationExcel_get_LanguageSettings,
    MSO_TO_OO_I_ApplicationExcel_get_Dummy101,
    MSO_TO_OO_I_ApplicationExcel_Dummy12,
    MSO_TO_OO_I_ApplicationExcel_get_AnswerWizard,
    MSO_TO_OO_I_ApplicationExcel_CalculateFull,
    MSO_TO_OO_I_ApplicationExcel_FindFile,
    MSO_TO_OO_I_ApplicationExcel_get_CalculationVersion,
    MSO_TO_OO_I_ApplicationExcel_get_ShowWindowsInTaskbar,
    MSO_TO_OO_I_ApplicationExcel_put_ShowWindowsInTaskbar,
    MSO_TO_OO_I_ApplicationExcel_get_FeatureInstall,
    MSO_TO_OO_I_ApplicationExcel_put_FeatureInstall,
    MSO_TO_OO_I_ApplicationExcel_get_Ready,
    MSO_TO_OO_I_ApplicationExcel_Dummy13,
    MSO_TO_OO_I_ApplicationExcel_get_FindFormat,
    MSO_TO_OO_I_ApplicationExcel_putref_FindFormat,
    MSO_TO_OO_I_ApplicationExcel_get_ReplaceFormat,
    MSO_TO_OO_I_ApplicationExcel_putref_ReplaceFormat,
    MSO_TO_OO_I_ApplicationExcel_get_UsedObjects,
    MSO_TO_OO_I_ApplicationExcel_get_CalculationState,
    MSO_TO_OO_I_ApplicationExcel_get_CalculationInterruptKey,
    MSO_TO_OO_I_ApplicationExcel_put_CalculationInterruptKey,
    MSO_TO_OO_I_ApplicationExcel_get_Watches,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayFunctionToolTips,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayFunctionToolTips,
    MSO_TO_OO_I_ApplicationExcel_get_AutomationSecurity,
    MSO_TO_OO_I_ApplicationExcel_put_AutomationSecurity,
    MSO_TO_OO_I_ApplicationExcel_get_FileDialog,
    MSO_TO_OO_I_ApplicationExcel_Dummy14,
    MSO_TO_OO_I_ApplicationExcel_CalculateFullRebuild,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayPasteOptions,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayPasteOptions,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayInsertOptions,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayInsertOptions,
    MSO_TO_OO_I_ApplicationExcel_get_GenerateGetPivotData,
    MSO_TO_OO_I_ApplicationExcel_put_GenerateGetPivotData,
    MSO_TO_OO_I_ApplicationExcel_get_AutoRecover,
    MSO_TO_OO_I_ApplicationExcel_get_Hwnd,
    MSO_TO_OO_I_ApplicationExcel_get_Hinstance,
    MSO_TO_OO_I_ApplicationExcel_CheckAbort,
    MSO_TO_OO_I_ApplicationExcel_get_ErrorCheckingOptions,
    MSO_TO_OO_I_ApplicationExcel_get_AutoFormatAsYouTypeReplaceHyperlinks,
    MSO_TO_OO_I_ApplicationExcel_put_AutoFormatAsYouTypeReplaceHyperlinks,
    MSO_TO_OO_I_ApplicationExcel_get_SmartTagRecognizers,
    MSO_TO_OO_I_ApplicationExcel_get_NewWorkbook,
    MSO_TO_OO_I_ApplicationExcel_get_SpellingOptions,
    MSO_TO_OO_I_ApplicationExcel_get_Speech,
    MSO_TO_OO_I_ApplicationExcel_get_MapPaperSize,
    MSO_TO_OO_I_ApplicationExcel_put_MapPaperSize,
    MSO_TO_OO_I_ApplicationExcel_get_ShowStartupDialog,
    MSO_TO_OO_I_ApplicationExcel_put_ShowStartupDialog,
    MSO_TO_OO_I_ApplicationExcel_get_DecimalSeparator,
    MSO_TO_OO_I_ApplicationExcel_put_DecimalSeparator,
    MSO_TO_OO_I_ApplicationExcel_get_ThousandsSeparator,
    MSO_TO_OO_I_ApplicationExcel_put_ThousandsSeparator,
    MSO_TO_OO_I_ApplicationExcel_get_UseSystemSeparators,
    MSO_TO_OO_I_ApplicationExcel_put_UseSystemSeparators,
    MSO_TO_OO_I_ApplicationExcel_get_ThisCell,
    MSO_TO_OO_I_ApplicationExcel_get_RTD,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayDocumentActionTaskPane,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayDocumentActionTaskPane,
    MSO_TO_OO_I_ApplicationExcel_DisplayXMLSourcePane,
    MSO_TO_OO_I_ApplicationExcel_get_ArbitraryXMLSupportAvailable,
    MSO_TO_OO_I_ApplicationExcel_Support,
};

HRESULT _ApplicationExcelConstructor(LPVOID *ppObj)
{
    _ApplicationExcelImpl *_applicationexcell;
    CLSID clsid;
    HRESULT hres;
    VARIANT result;
    VARIANT param1;
    IUnknown *punk = NULL;

    TRACE("(%p) \n", ppObj);

    _applicationexcell = HeapAlloc(GetProcessHeap(), 0, sizeof(*_applicationexcell));
    if (!_applicationexcell) {
        return E_OUTOFMEMORY;
    }

    _applicationexcell->pApplicationExcelVtbl = &MSO_TO_OO_I_ApplicationExcel_Vtbl;
    _applicationexcell->pConnectionPointContainerVtbl = &MSO_TO_OO_ConnectionPointContainerVtbl;
    _applicationexcell->pConnectionPointVtbl = &MSO_TO_OO_ConnectionPointVtbl;
    _applicationexcell->ref = 0;
    _applicationexcell->pdOOApp = NULL;
    _applicationexcell->pdOODesktop = NULL;
    _applicationexcell->pdWorkbooks = NULL;

    /*Создание указателей на объекты openOfffice 
    Create OpenOffice Service Manager */
    hres = CLSIDFromProgID(L"com.sun.star.ServiceManager", &clsid);
    if (FAILED(hres))
        return E_NOINTERFACE;

    /* Start server and get IDispatch...*/
    hres = CoCreateInstance(&clsid, NULL, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER, &IID_IDispatch, (void**) &(_applicationexcell->pdOOApp));
    if (FAILED(hres))
        return E_NOINTERFACE;

    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"com.sun.star.frame.Desktop");
    /* Get Desktop and its assoc. IDispatch...*/
    hres = AutoWrap(DISPATCH_METHOD, &result, _applicationexcell->pdOOApp, L"CreateInstance", 1, param1);

    if (FAILED(hres))
        return E_NOINTERFACE;

    _applicationexcell->pdOODesktop = result.pdispVal;
    hres = IDispatch_AddRef(_applicationexcell->pdOODesktop);

    hres = _I_WorkbooksConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Workbooks_QueryInterface(punk, &IID_I_Workbooks, (void**) &(_applicationexcell->pdWorkbooks));
    if (FAILED(hres)) return E_NOINTERFACE;
/*    I_Workbooks_Release(punk);*/

    MSO_TO_OO_I_Workbooks_Initialize((I_Workbooks*)(_applicationexcell->pdWorkbooks), (I_ApplicationExcel*)_applicationexcell);

    *ppObj = APPEXCEL(_applicationexcell);

    /*освобождаем память выделенную под строку*/
    SysFreeString(V_BSTR(&param1));
    VariantClear(&result);
    return S_OK;
}


