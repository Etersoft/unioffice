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

ITypeInfo *ti_excel = NULL;

HRESULT get_typeinfo_application(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if(ti_excel) {
        *typeinfo = ti_excel;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID__Application, &ti_excel);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_excel;

    return hres;
}

/*IConnectionPoint interface*/

#define CONPOINT_THIS(iface) DEFINE_THIS(_ApplicationImpl,ConnectionPoint,iface);

    /*** IUnknown methods ***/
static HRESULT WINAPI MSO_TO_OO_ConnectionPoint_QueryInterface(
        IConnectionPoint* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationImpl *This = CONPOINT_THIS(iface);
    return _Application_QueryInterface(APPEXCEL(This), riid, ppvObject);
}

static ULONG WINAPI MSO_TO_OO_ConnectionPoint_AddRef(
        IConnectionPoint* iface)
{
    _ApplicationImpl *This = CONPOINT_THIS(iface);
    return _Application_AddRef(APPEXCEL(This));
}

static ULONG WINAPI MSO_TO_OO_ConnectionPoint_Release(
        IConnectionPoint* iface)
{
    _ApplicationImpl *This = CONPOINT_THIS(iface);
    return _Application_Release(APPEXCEL(This));
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
    _ApplicationImpl *This = CONPOINT_THIS(iface);
    TRACE_IN;

    *ppCPC = (IConnectionPointContainer*)CONPOINTCONT(This);
    if (*ppCPC) {
        IConnectionPointContainer_AddRef(*ppCPC);
        TRACE_OUT;
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
    TRACE_IN;
    *pdwCookie = 0;
    TRACE_OUT;
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

#define CONPOINTCONT_THIS(iface) DEFINE_THIS(_ApplicationImpl,ConnectionPointContainer,iface);

    /*** IUnknown methods ***/
static HRESULT WINAPI MSO_TO_OO_ConnectionPointContainer_QueryInterface(
        IConnectionPointContainer* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationImpl *This = CONPOINTCONT_THIS(iface);
    return _Application_QueryInterface(APPEXCEL(This), riid, ppvObject);
}

static ULONG WINAPI MSO_TO_OO_ConnectionPointContainer_AddRef(
        IConnectionPointContainer* iface)
{
    _ApplicationImpl *This = CONPOINTCONT_THIS(iface);
    return _Application_AddRef(APPEXCEL(This));
}

static ULONG WINAPI MSO_TO_OO_ConnectionPointContainer_Release(
        IConnectionPointContainer* iface)
{
    _ApplicationImpl *This = CONPOINTCONT_THIS(iface);
    return _Application_Release(APPEXCEL(This));
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
    _ApplicationImpl *This = CONPOINTCONT_THIS(iface);
    WCHAR str_clsid[39];
    StringFromGUID2(riid, str_clsid, 39);
    TRACE_IN;
    WTRACE(L"riid = (%s) \n", str_clsid);

    *ppCP = (IConnectionPoint*)CONPOINT(This);
    if (*ppCP) {
        IConnectionPoint_AddRef(*ppCP);
        TRACE_OUT;
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

/*_Application interface*/

/*
IUnknown
*/

#define APPEXCEL_THIS(iface) DEFINE_THIS(_ApplicationImpl, Application, iface);

static ULONG WINAPI MSO_TO_OO__Application_AddRef(
        _Application* iface)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    ULONG ref;

    if (This == NULL) {
        TRACE("Object is NULL \n");
        return E_POINTER;
    }

    TRACE("REF=%i \n", This->ref);

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }

    return ref;
}

static HRESULT WINAPI MSO_TO_OO__Application_QueryInterface(
        _Application* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    WCHAR str_clsid[39];

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    *ppvObject = NULL;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID__Application)) {
        TRACE("IApplicationExcel \n");
        *ppvObject = APPEXCEL(This);
    }
    if (IsEqualGUID(riid, &IID_IConnectionPointContainer)) {
        TRACE("IConnectionPointContainer \n");
        *ppvObject = CONPOINTCONT(This);
    }

    if (*ppvObject) {
        _Application_AddRef(iface);
        return S_OK;
    }

    StringFromGUID2(riid, str_clsid, 39);
    WTRACE(L"(%s) not supported \n", str_clsid);
    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO__Application_Release(
        _Application* iface)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    ULONG ref;

    if (This == NULL) return E_POINTER;

    TRACE("REF = %i \n", This->ref);

    ref = InterlockedDecrement(&This->ref);

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
        DELETE_OBJECT;
    }

    return ref;
}

/*
_Application
*/
static HRESULT WINAPI MSO_TO_OO__Application_put_UserControl(
        _Application* iface,
        VARIANT_BOOL vbUserControl)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UserControl(
        _Application* iface,
        VARIANT_BOOL *vbUserControl)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayAlerts(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL vbDisplayAlerts)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    This->displayalerts = vbDisplayAlerts;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayAlerts(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *vbDisplayAlerts)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    *vbDisplayAlerts = This->displayalerts;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_WindowState(
        _Application* iface,
        LCID lcid,
        XlWindowState State)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_WindowState(
        _Application* iface,
        LCID lcid,
        XlWindowState *State)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Visible(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL vbVisible)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    WorkbooksImpl* wbs = (WorkbooksImpl*)This->pdWorkbooks;
    int i;
    TRACE_IN;

    for (i=0; i<wbs->count_workbooks;i++) {
        MSO_TO_OO_Workbook_SetVisible((I_Workbook*)(wbs->pworkbook[i]), vbVisible);
    }

    This->visible = vbVisible;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Visible(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *vbVisible)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

   *vbVisible = This->visible;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Workbooks(
        _Application* iface,
        IDispatch **ppWorkbooks)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    if (This->pdWorkbooks==NULL)
       return E_POINTER;

    *ppWorkbooks = This->pdWorkbooks;

    I_Workbooks_AddRef(This->pdWorkbooks);

    if (ppWorkbooks==NULL)
       return E_POINTER;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Sheets(
        _Application* iface,
        IDispatch **ppSheets)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    HRESULT hres;
    TRACE_IN;

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) return E_FAIL;

    hres = I_Workbook_get_Sheets(pwb, ppSheets);
    if (FAILED(hres)) {
        I_Workbook_Release(pwb);
        return E_FAIL;
    }

    I_Workbook_Release(pwb);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Worksheets(
        _Application* iface,
        IDispatch **ppSheets)
{
   /*Используем Sheets - они выполняют одинаковые функции*/
   TRACE(" ----> get_Sheets");
   return MSO_TO_OO__Application_get_Sheets(iface, ppSheets);
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Cells(
        _Application* iface,
        IDispatch **ppRange)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    I_Sheets *pSheets;
    I_Worksheet *pworksheet;
    HRESULT hres;
    TRACE_IN;

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
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActiveSheet(
        _Application* iface,
        IDispatch **RHS)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    I_Sheets *pSheets;
    HRESULT hres;
    TRACE_IN;

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

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Version(
        _Application* iface,
        long Lcid,
        BSTR *pVersion)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    if (pVersion == NULL)
        return E_POINTER;

    *pVersion = SysAllocString(OLESTR("11.0"));

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_ConvertFormula(
        _Application* iface,
        VARIANT Formula,
        XlReferenceStyle FromReferenceStyle,
        VARIANT ToReferenceStyle,
        VARIANT ToAbsolute,
        VARIANT RelativeTo,
        long Lcid,
        VARIANT *pResult)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    MSO_TO_OO_CorrectArg(Formula, &Formula);
    MSO_TO_OO_CorrectArg(ToReferenceStyle, &ToReferenceStyle);
    MSO_TO_OO_CorrectArg(ToAbsolute, &ToAbsolute);
    MSO_TO_OO_CorrectArg(RelativeTo, &RelativeTo);

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_Quit(
        _Application* iface)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    VARIANT res;
    TRACE_IN;

    VariantInit(&res);
    if (iface==NULL) {
        TRACE("ERROR Object is NULL\n");
        return E_FAIL;
    }
    /*При вызове этого метода вызываем метод Close объекта WorkBooks*/
    I_Workbooks_Close((I_Workbooks*)(This->pdWorkbooks), 0);

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &res, This->pdOODesktop, L"terminate", 0);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActiveCell(
        _Application* iface,
        IDispatch **RHS)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);

    HRESULT hres;
    TRACE_IN;

    hres = MSO_TO_OO_GetActiveCells((I_Workbooks*)This->pdWorkbooks, (I_Range**) RHS);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Application(
        _Application* iface,
        IDispatch **value)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    if (iface!=NULL) {
        *value = (IDispatch*)APPEXCEL(This);
        MSO_TO_OO__Application_AddRef((_Application*)*value);
    } else {
        return E_FAIL;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_EnableEvents(
        _Application* iface,
        VARIANT_BOOL *pvbee)
{
    TRACE_IN;
    /*Always return TRUE*/
    *pvbee = VARIANT_TRUE;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_EnableEvents(
        _Application* iface,
        VARIANT_BOOL vbee)
{
    TRACE_NOTIMPL;
    /*Always return S_OK*/
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ScreenUpdating(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL vbscup)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    HRESULT hres;
    IDispatch *wb;
    VARIANT tmp;
    TRACE_IN;

    VariantInit(&tmp);

    if (vbscup == VARIANT_TRUE) {
        hres = _Application_get_ActiveWorkbook(iface, &wb);
        WorkbookImpl *wbi = (WorkbookImpl*)wb;
        if (FAILED(hres)) {
            TRACE("ERROR when get ActiveWorkbook \n");
            return E_FAIL;
        }

        hres = AutoWrap(DISPATCH_METHOD, &tmp, wbi->pDoc, L"unLockControllers", 0);
        if (FAILED(hres)) {
            TRACE("ERROR When unLockControllers\n");
            IDispatch_Release(wb);
            return E_FAIL;
        }

        VariantClear(&tmp);
        hres = AutoWrap(DISPATCH_METHOD, &tmp, wbi->pDoc, L"removeActionLock", 0);
        if (FAILED(hres)) {
            TRACE("ERROR When removeActionLock\n");
            IDispatch_Release(wb);
            return E_FAIL;
        }
    } else {
/*
Отключение вывода
Document.OleFunction("lockControllers");
Document.OleFunction("addActionLock");
*/
        hres = _Application_get_ActiveWorkbook(iface, &wb);
        WorkbookImpl *wbi = (WorkbookImpl*)wb;
        if (FAILED(hres)) {
            TRACE("ERROR when get ActiveWorkbook \n");
            return E_FAIL;
        }

        hres = AutoWrap(DISPATCH_METHOD, &tmp, wbi->pDoc, L"lockControllers", 0);
        if (FAILED(hres)) {
            TRACE("ERROR When lockControllers\n");
            IDispatch_Release(wb);
            return E_FAIL;
        }

        VariantClear(&tmp);
        hres = AutoWrap(DISPATCH_METHOD, &tmp, wbi->pDoc, L"addActionLock", 0);
        if (FAILED(hres)) {
            TRACE("ERROR When addActionLock\n");
            IDispatch_Release(wb);
            return E_FAIL;
        }


    }

    This->screenupdating = vbscup;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ScreenUpdating(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *vbscup)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    *vbscup = This->screenupdating;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Caption(
        _Application* iface,
        VARIANT *vName)
{
    TRACE_IN;
    if (vName==NULL) {
        TRACE("ERROR object is NULL\n");
    }
    V_VT(vName) = VT_BSTR;
    V_BSTR(vName) = SysAllocString(L"Microsoft Excel");
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Caption(
        _Application* iface,
        VARIANT vName)
{
    TRACE_NOTIMPL;
    MSO_TO_OO_CorrectArg(vName, &vName);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActiveWorkbook(
        _Application* iface,
        IDispatch **result)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    I_Sheets *pSheets;
    HRESULT hres;
    TRACE_IN;

    if (This==NULL) return E_FAIL;

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) {
        TRACE("ERROR when GetActiveWorkbook\n");
        *result = NULL;
        return hres;
    }
    *result = (IDispatch*)pwb;

    I_Workbook_AddRef((I_Workbook*)*result);
    I_Workbook_Release(pwb);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Range(
        _Application* iface,
        VARIANT Cell1,
        VARIANT Cell2,
        IDispatch **ppRange)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    HRESULT hres; 
    I_Worksheet *wsh;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(Cell1, &Cell1);
    MSO_TO_OO_CorrectArg(Cell2, &Cell2);

    hres = MSO_TO_OO__Application_get_ActiveSheet(iface, (IDispatch**) &wsh);

    hres = I_Worksheet_get_Range(wsh,Cell1, Cell2, ppRange);

    I_Worksheet_Release(wsh);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Columns(
        _Application* iface,
        VARIANT param,
        IDispatch **ppRange)
{
    HRESULT hres;
    IDispatch *active_sheet;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    hres = MSO_TO_OO__Application_get_ActiveSheet(iface, &active_sheet);

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Rows(
        _Application* iface,
        VARIANT param,
        IDispatch **ppRange)
{
    HRESULT hres;
    IDispatch *active_sheet;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    hres = MSO_TO_OO__Application_get_ActiveSheet(iface, &active_sheet);

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Selection(
        _Application* iface,
        IDispatch **ppRange)
{
    WorkbookImpl *awb;
    IDispatch *asheet;
    IUnknown *pobj;
    HRESULT hres;
    VARIANT vRes, vRet;
    RangeImpl *range;
    TRACE_IN;

    VariantInit(&vRes);
    VariantInit(&vRet);

    hres = MSO_TO_OO__Application_get_ActiveWorkbook(iface, (IDispatch**)&awb);
    if (FAILED(hres)) {
        TRACE("ERROR when get_ActiveWorkbook\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO__Application_get_ActiveSheet(iface, &asheet);
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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Creator(
        _Application* iface,
        XlCreator *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Parent(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActiveChart(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActiveDialog(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActiveMenuBar(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActivePrinter(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ActivePrinter(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ActiveWindow(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AddIns(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Assistant(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Calculate(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Charts(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CommandBars(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DDEAppReturnCode(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DDEExecute(
        _Application* iface,
        long Channel,
        BSTR String,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DDEInitiate(
        _Application* iface,
        BSTR App,
        BSTR Topic,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DDEPoke(
        _Application* iface,
        long Channel,
        VARIANT Item,
        VARIANT Data,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DDERequest(
        _Application* iface,
        long Channel,
        BSTR Item,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DDETerminate(
        _Application* iface,
        long Channel,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DialogSheets(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Evaluate(
        _Application* iface,
        VARIANT Name,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application__Evaluate(
        _Application* iface,
        VARIANT Name,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_ExecuteExcel4Macro(
        _Application* iface,
        BSTR String,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Intersect(
        _Application* iface,
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
        LCID lcid,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MenuBars(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Modules(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Names(
        _Application* iface,
        IDispatch **RHS)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    I_Workbook *pwb;
    HRESULT hres;
    TRACE_IN;

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) return E_FAIL;

    hres = I_Workbook_get_Names(pwb, RHS);
    if (FAILED(hres)) {
        TRACE("ERROR get_Names\n");
        *RHS = NULL;
    }
    I_Workbook_Release(pwb);
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_Run(
        _Application* iface,
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application__Run2(
        _Application* iface,
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
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_SendKeys(
        _Application* iface,
        VARIANT Keys,
        VARIANT Wait,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ShortcutMenus(
        _Application* iface,
        long Index,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ThisWorkbook(
        _Application* iface,
        LCID lcid,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Toolbars(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Union(
        _Application* iface,
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
        LCID lcid,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Windows(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_WorksheetFunction(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Worksheets(
        _Application* iface,
        IDispatch **ppSheets);

static HRESULT WINAPI MSO_TO_OO__Application_get_Excel4IntlMacroSheets(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Excel4MacroSheets(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_ActivateMicrosoftApp(
        _Application* iface,
        XlMSApplication Index,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_AddChartAutoFormat(
        _Application* iface,
        VARIANT Chart,
        BSTR Name,
        VARIANT Description,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_AddCustomList(
        _Application* iface,
        VARIANT ListArray,
        VARIANT ByRow,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AlertBeforeOverwriting(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_AlertBeforeOverwriting(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AltStartupPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_AltStartupPath(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AskToUpdateLinks(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_AskToUpdateLinks(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_EnableAnimations(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_EnableAnimations(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AutoCorrect(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Build(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CalculateBeforeSave(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_CalculateBeforeSave(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Calculation(
        _Application* iface,
        LCID lcid,
        XlCalculation *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Calculation(
        _Application* iface,
        LCID lcid,
        XlCalculation RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Caller(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CanPlaySounds(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CanRecordSounds(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CellDragAndDrop(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_CellDragAndDrop(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_CentimetersToPoints(
        _Application* iface,
        double Centimeters,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_CheckSpelling(
        _Application* iface,
        BSTR Word,
        VARIANT CustomDictionary,
        VARIANT IgnoreUppercase,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ClipboardFormats(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayClipboardWindow(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayClipboardWindow(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ColorButtons(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ColorButtons(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CommandUnderlines(
        _Application* iface,
        LCID lcid,
        XlCommandUnderlines *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_CommandUnderlines(
        _Application* iface,
        LCID lcid,
        XlCommandUnderlines RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ConstrainNumeric(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ConstrainNumeric(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CopyObjectsWithCells(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_CopyObjectsWithCells(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Cursor(
        _Application* iface,
        LCID lcid,
        XlMousePointer *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Cursor(
        _Application* iface,
        LCID lcid,
        XlMousePointer RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CustomListCount(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CutCopyMode(
        _Application* iface,
        LCID lcid,
        XlCutCopyMode *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_CutCopyMode(
        _Application* iface,
        LCID lcid,
        XlCutCopyMode RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}
static HRESULT WINAPI MSO_TO_OO__Application_get_DataEntryMode(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DataEntryMode(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy1(
        _Application* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy2(
        _Application* iface,
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy3(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy4(
        _Application* iface,
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy5(
        _Application* iface,
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy6(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy7(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy8(
        _Application* iface,
        VARIANT Arg1,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy9(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy10(
        _Application* iface,
        VARIANT arg,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy11(
        _Application* This)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get__Default(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DefaultFilePath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DefaultFilePath(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DeleteChartAutoFormat(
        _Application* iface,
        BSTR Name,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DeleteCustomList(
        _Application* iface,
        long ListNum,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Dialogs(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayFormulaBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayFormulaBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayFullScreen(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayFullScreen(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayNoteIndicator(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayNoteIndicator(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayCommentIndicator(
        _Application* iface,
        XlCommentDisplayMode *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayCommentIndicator(
        _Application* iface,
        XlCommentDisplayMode RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayExcel4Menus(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayExcel4Menus(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayRecentFiles(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayRecentFiles(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayScrollBars(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayScrollBars(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayStatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayStatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DoubleClick(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_EditDirectlyInCell(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_EditDirectlyInCell(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_EnableAutoComplete(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_EnableAutoComplete(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}
static HRESULT WINAPI MSO_TO_OO__Application_get_EnableCancelKey(
        _Application* iface,
        LCID lcid,
        XlEnableCancelKey *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_EnableCancelKey(
        _Application* iface,
        LCID lcid,
        XlEnableCancelKey RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_EnableSound(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_EnableSound(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_EnableTipWizard(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_EnableTipWizard(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FileConverters(
        _Application* iface,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FileSearch(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FileFind(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application__FindFile(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FixedDecimal(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_FixedDecimal(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FixedDecimalPlaces(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_FixedDecimalPlaces(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_GetCustomListContents(
        _Application* iface,
        long ListNum,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_GetCustomListNum(
        _Application* iface,
        VARIANT ListArray,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_GetOpenFilename(
        _Application* iface,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        VARIANT MultiSelect,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_GetSaveAsFilename(
        _Application* iface,
        VARIANT InitialFilename,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Goto(
        _Application* iface,
        VARIANT Reference,
        VARIANT Scroll,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Height(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Height(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Help(
        _Application* iface,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_IgnoreRemoteRequests(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_IgnoreRemoteRequests(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_InchesToPoints(
        _Application* iface,
        double Inches,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_InputBox(
        _Application* iface,
        BSTR Prompt,
        VARIANT Title,
        VARIANT Default,
        VARIANT Left,
        VARIANT Top,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        VARIANT Type,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Interactive(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Interactive(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_International(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Iteration(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Iteration(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_LargeButtons(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_LargeButtons(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Left(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Left(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_LibraryPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_MacroOptions(
        _Application* iface,
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
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_MailLogoff(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_MailLogon(
        _Application* iface,
        VARIANT Name,
        VARIANT Password,
        VARIANT DownloadNewMail,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MailSession(
        _Application* iface,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MailSystem(
        _Application* iface,
        LCID lcid,
        XlMailSystem *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MathCoprocessorAvailable(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MaxChange(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_MaxChange(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MaxIterations(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_MaxIterations(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MemoryFree(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MemoryTotal(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MemoryUsed(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MouseAvailable(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MoveAfterReturn(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_MoveAfterReturn(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MoveAfterReturnDirection(
        _Application* iface,
        LCID lcid,
        XlDirection *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_MoveAfterReturnDirection(
        _Application* iface,
        LCID lcid,
        XlDirection RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_RecentFiles(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Name(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_NextLetter(
        _Application* iface,
        LCID lcid,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_NetworkTemplatesPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ODBCErrors(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ODBCTimeout(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ODBCTimeout(
        _Application* iface,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OnCalculate(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_OnCalculate(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OnData(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_OnData(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OnDoubleClick(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_OnDoubleClick(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OnEntry(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_OnEntry(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_OnKey(
        _Application* iface,
        BSTR Key,
        VARIANT Procedure,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_OnRepeat(
        _Application* iface,
        BSTR Text,
        BSTR Procedure,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OnSheetActivate(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_OnSheetActivate(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OnSheetDeactivate(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_OnSheetDeactivate(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_OnTime(
        _Application* iface,
        VARIANT EarliestTime,
        BSTR Procedure,
        VARIANT LatestTime,
        VARIANT Schedule,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_OnUndo(
        _Application* iface,
        BSTR Text,
        BSTR Procedure,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OnWindow(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_OnWindow(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OperatingSystem(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OrganizationName(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Path(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_PathSeparator(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_PreviousSelections(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_PivotTableSelection(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_PivotTableSelection(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_PromptForSummaryInfo(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_PromptForSummaryInfo(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_RecordMacro(
        _Application* iface,
        VARIANT BasicCode,
        VARIANT XlmCode,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_RecordRelative(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ReferenceStyle(
        _Application* iface,
        LCID lcid,
        XlReferenceStyle *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ReferenceStyle(
        _Application* iface,
        LCID lcid,
        XlReferenceStyle RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_RegisteredFunctions(
        _Application* iface,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_RegisterXLL(
        _Application* iface,
        BSTR Filename,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Repeat(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_ResetTipWizard(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_RollZoom(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_RollZoom(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Save(
        _Application* iface,
        VARIANT Filename,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_SaveWorkspace(
        _Application* iface,
        VARIANT Filename,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_SetDefaultChart(
        _Application* iface,
        VARIANT FormatName,
        VARIANT Gallery)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_SheetsInNewWorkbook(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE("\n");
    *RHS = This->sheetsinnewworkbook;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_SheetsInNewWorkbook(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE("\n");
    This->sheetsinnewworkbook = RHS;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ShowChartTipNames(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ShowChartTipNames(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ShowChartTipValues(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ShowChartTipValues(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_StandardFont(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_StandardFont(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_StandardFontSize(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_StandardFontSize(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_StartupPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_StatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_StatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_TemplatesPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ShowToolTips(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ShowToolTips(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Top(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Top(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DefaultSaveFormat(
        _Application* iface,
        XlFileFormat *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DefaultSaveFormat(
        _Application* iface,
        XlFileFormat RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_TransitionMenuKey(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_TransitionMenuKey(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_TransitionMenuKeyAction(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_TransitionMenuKeyAction(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_TransitionNavigKeys(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_TransitionNavigKeys(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Undo(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UsableHeight(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UsableWidth(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UserName(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_UserName(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Value(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_VBE(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Volatile(
        _Application* iface,
        VARIANT Volatile,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application__Wait(
        _Application* iface,
        VARIANT Time,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Width(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_Width(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_WindowsForPens(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UILanguage(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_UILanguage(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DefaultSheetDirection(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DefaultSheetDirection(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CursorMovement(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_CursorMovement(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ControlCharacters(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ControlCharacters(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application__WSFunction(
        _Application* iface,
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
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayInfoWindow(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayInfoWindow(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Wait(
        _Application* iface,
        VARIANT Time,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ExtendList(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ExtendList(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_OLEDBErrors(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_GetPhonetic(
        _Application* iface,
        VARIANT Text,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_COMAddIns(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DefaultWebOptions(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ProductCode(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UserLibraryPath(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AutoPercentEntry(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_AutoPercentEntry(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_LanguageSettings(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Dummy101(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy12(
        _Application* iface,
        IDispatch *p1,
        IDispatch *p2)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AnswerWizard(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_CalculateFull(
        _Application* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_FindFile(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CalculationVersion(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ShowWindowsInTaskbar(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ShowWindowsInTaskbar(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FeatureInstall(
        _Application* iface,
        MsoFeatureInstall *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_FeatureInstall(
        _Application* iface,
        MsoFeatureInstall RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Ready(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy13(
        _Application* iface,
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
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FindFormat(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_putref_FindFormat(
        _Application* iface,
        IDispatch *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ReplaceFormat(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_putref_ReplaceFormat(
        _Application* iface,
        IDispatch *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UsedObjects(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CalculationState(
        _Application* iface,
        XlCalculationState *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_CalculationInterruptKey(
        _Application* iface,
        XlCalculationInterruptKey *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_CalculationInterruptKey(
        _Application* iface,
        XlCalculationInterruptKey RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Watches(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayFunctionToolTips(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayFunctionToolTips(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AutomationSecurity(
        _Application* iface,
        MsoAutomationSecurity *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_AutomationSecurity(
        _Application* iface,
        MsoAutomationSecurity RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_FileDialog(
        _Application* iface,
        MsoFileDialogType fileDialogType,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Dummy14(
        _Application* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_CalculateFullRebuild(
        _Application* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayPasteOptions(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayPasteOptions(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayInsertOptions(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
};

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayInsertOptions(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_GenerateGetPivotData(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_GenerateGetPivotData(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AutoRecover(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Hwnd(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Hinstance(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_CheckAbort(
        _Application* iface,
        VARIANT KeepAbort)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ErrorCheckingOptions(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_AutoFormatAsYouTypeReplaceHyperlinks(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_AutoFormatAsYouTypeReplaceHyperlinks(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_SmartTagRecognizers(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_NewWorkbook(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_SpellingOptions(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_Speech(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_MapPaperSize(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_MapPaperSize(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ShowStartupDialog(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ShowStartupDialog(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DecimalSeparator(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DecimalSeparator(
        _Application* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ThousandsSeparator(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_ThousandsSeparator(
        _Application* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_UseSystemSeparators(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_UseSystemSeparators(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ThisCell(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_RTD(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_DisplayDocumentActionTaskPane(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_put_DisplayDocumentActionTaskPane(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_DisplayXMLSourcePane(
        _Application* iface,
        VARIANT XmlMap)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_get_ArbitraryXMLSupportAvailable(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO__Application_Support(
        _Application* iface,
        IDispatch *Object,
        long ID,
        VARIANT arg,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

/*
IDispatch
*/
static HRESULT WINAPI MSO_TO_OO__Application_GetTypeInfoCount(
        _Application* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO__Application_GetTypeInfo(
        _Application* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_application(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_GetIDsOfNames(
        _Application* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_application(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO__Application_Invoke(
        _Application* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    HRESULT hres;
    BSTR pVersion;
    IDispatch *pdisp;
    IDispatch *pretdisp;
    long tmp;
    VARIANT vNull;
    VARIANT vRet,vtmp,cell1,cell2;
    VARIANT_BOOL vbin;
    ITypeInfo *typeinfo;

    VariantInit(&vNull);
    VariantInit(&vtmp);
    VariantInit(&cell1);
    VariantInit(&cell2);
    VariantInit(&vRet);

    TRACE(" %i\n", dispIdMember);

    if (This == NULL) return E_POINTER;

    /*special operation*/
    if ((dispIdMember == dispid_application_range) && (wFlags == DISPATCH_PROPERTYPUT)) {
            switch (pDispParams->cArgs) {
                case 2:
                    hres = MSO_TO_OO__Application_get_Range(iface,pDispParams->rgvarg[1], vNull, &pretdisp);
                    if (FAILED(hres)) {
                        pExcepInfo->bstrDescription=SysAllocString(str_error);
                        TRACE("(case 2) ERROR get_range hres = %08x\n",hres);
                        return hres;
                    }
                    I_Range_put_Value((I_Range*)pretdisp, vNull, 0, pDispParams->rgvarg[0]);
                    IDispatch_Release(pretdisp);
                    return S_OK;
                case 3:
                    hres = MSO_TO_OO__Application_get_Range(iface,pDispParams->rgvarg[2], pDispParams->rgvarg[1], &pretdisp);
                    if (FAILED(hres)) {
                        pExcepInfo->bstrDescription=SysAllocString(str_error);
                        TRACE("(case 2) ERROR get_range hres = %08x\n",hres);
                        return hres;
                    }
                    I_Range_put_Value((I_Range*)pretdisp, vNull, 0, pDispParams->rgvarg[0]);
                    IDispatch_Release(pretdisp);
                return S_OK;
            }
    }

    switch (dispIdMember)
    {
    case dispid_application_displayalerts:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE(" (case 2) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO__Application_put_DisplayAlerts(iface, 0, vbin);
        } else {
            hres = MSO_TO_OO__Application_get_DisplayAlerts(iface, 0, &vbin);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = vbin;
            }
            return S_OK;
        }
    case dispid_application_windowstate:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            /*преобразовываем любой тип к I4*/
            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (case 3) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            tmp = V_I4(&vtmp);
            MSO_TO_OO__Application_put_WindowState(iface, 0, tmp);
            return S_OK;
        } else {

            return E_NOTIMPL;
        }
    case dispid_application_visible:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE(" (case 4) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO__Application_put_Visible(iface, 0, vbin);
        } else {
            hres = MSO_TO_OO__Application_get_Visible(iface, 0, &vbin);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = vbin;
            }
            return S_OK;
        }
    case dispid_application_cells:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            switch (pDispParams->cArgs) {
            case 3:
                hres = MSO_TO_OO__Application_get_Cells(iface,&pdisp);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
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
            hres = MSO_TO_OO__Application_get_Cells(iface,&pdisp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
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
            return hres;
        }
    case dispid_application_version:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO__Application_get_Version(iface,0,&pVersion);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_BSTR;
                V_BSTR(pVarResult)=pVersion;
            }
            return hres;
        }
    case dispid_application_convertformula:
        /*MSO_TO_OO__Application_ConvertFormula*/
        if (pDispParams->cArgs<3) return E_FAIL;

        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[pDispParams->cArgs-2]), 0, 0, VT_I4);
        if (FAILED(hres)) {
            TRACE(" (case 11) ERROR when VariantChangeTypeEx\n");
            return E_FAIL;
        }
        tmp = V_I4(&vtmp);

        hres = MSO_TO_OO__Application_ConvertFormula(iface, pDispParams->rgvarg[pDispParams->cArgs-1], tmp, pDispParams->rgvarg[pDispParams->cArgs-3], vNull, vNull, tmp, &vRet);
        if (FAILED(hres)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hres;
        }
        if (pVarResult!=NULL){
            V_VT(pVarResult)=VT_BSTR;
            V_BSTR(pVarResult)=V_BSTR(&vRet);
        }
        return S_OK;
    case dispid_application_enableevents:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE(" (case 15) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
           return MSO_TO_OO__Application_put_EnableEvents(iface, vbin);
        } else {
            hres = MSO_TO_OO__Application_get_EnableEvents(iface,&vbin);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
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
            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE(" (case 16) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO__Application_put_ScreenUpdating(iface, 0, vbin);
        } else {
            hres = MSO_TO_OO__Application_get_ScreenUpdating(iface, 0, &vbin);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = vbin;
            }
            return S_OK;
        }
    case dispid_application_sheetsinnewworkbook:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            /*преобразовываем любой тип к I4*/
            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("sheetsinnewworkbook ERROR when VariantChangeTypeEx\n");
                return E_FAIL;
            }
            tmp = V_I4(&vtmp);
            return MSO_TO_OO__Application_put_SheetsInNewWorkbook(iface, 0, tmp);
        } else {
            hres = MSO_TO_OO__Application_get_SheetsInNewWorkbook(iface, 0, &tmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return E_FAIL;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = tmp;
            }
            return S_OK;
        }
    default:
        hres = get_typeinfo_application(&typeinfo);
        if (FAILED(hres))
            return hres;

        hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
        if (FAILED(hres)) {
            TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
        }

        return hres;
    }

    return E_NOTIMPL;
}

#undef APPEXCEL_THIS


const _ApplicationVtbl MSO_TO_OO__Application_Vtbl =
{
    MSO_TO_OO__Application_QueryInterface,
    MSO_TO_OO__Application_AddRef,
    MSO_TO_OO__Application_Release,
    MSO_TO_OO__Application_GetTypeInfoCount,
    MSO_TO_OO__Application_GetTypeInfo,
    MSO_TO_OO__Application_GetIDsOfNames,
    MSO_TO_OO__Application_Invoke,
    MSO_TO_OO__Application_get_Application,
    MSO_TO_OO__Application_get_Creator,
    MSO_TO_OO__Application_get_Parent,
    MSO_TO_OO__Application_get_ActiveCell,
    MSO_TO_OO__Application_get_ActiveChart,
    MSO_TO_OO__Application_get_ActiveDialog,
    MSO_TO_OO__Application_get_ActiveMenuBar,
    MSO_TO_OO__Application_get_ActivePrinter,
    MSO_TO_OO__Application_put_ActivePrinter,
    MSO_TO_OO__Application_get_ActiveSheet,
    MSO_TO_OO__Application_get_ActiveWindow,
    MSO_TO_OO__Application_get_ActiveWorkbook,
    MSO_TO_OO__Application_get_AddIns,
    MSO_TO_OO__Application_get_Assistant,
    MSO_TO_OO__Application_Calculate,
    MSO_TO_OO__Application_get_Cells,
    MSO_TO_OO__Application_get_Charts,
    MSO_TO_OO__Application_get_Columns,
    MSO_TO_OO__Application_get_CommandBars,
    MSO_TO_OO__Application_get_DDEAppReturnCode,
    MSO_TO_OO__Application_DDEExecute,
    MSO_TO_OO__Application_DDEInitiate,
    MSO_TO_OO__Application_DDEPoke,
    MSO_TO_OO__Application_DDERequest,
    MSO_TO_OO__Application_DDETerminate,
    MSO_TO_OO__Application_get_DialogSheets,
    MSO_TO_OO__Application_Evaluate,
    MSO_TO_OO__Application__Evaluate,
    MSO_TO_OO__Application_ExecuteExcel4Macro,
    MSO_TO_OO__Application_Intersect,
    MSO_TO_OO__Application_get_MenuBars,
    MSO_TO_OO__Application_get_Modules,
    MSO_TO_OO__Application_get_Names,
    MSO_TO_OO__Application_get_Range,
    MSO_TO_OO__Application_get_Rows,
    MSO_TO_OO__Application_Run,
    MSO_TO_OO__Application__Run2,
    MSO_TO_OO__Application_get_Selection,
    MSO_TO_OO__Application_SendKeys,
    MSO_TO_OO__Application_get_Sheets,
    MSO_TO_OO__Application_get_ShortcutMenus,
    MSO_TO_OO__Application_get_ThisWorkbook,
    MSO_TO_OO__Application_get_Toolbars,
    MSO_TO_OO__Application_Union,
    MSO_TO_OO__Application_get_Windows,
    MSO_TO_OO__Application_get_Workbooks,
    MSO_TO_OO__Application_get_WorksheetFunction,
    MSO_TO_OO__Application_get_Worksheets,
    MSO_TO_OO__Application_get_Excel4IntlMacroSheets,
    MSO_TO_OO__Application_get_Excel4MacroSheets,
    MSO_TO_OO__Application_ActivateMicrosoftApp,
    MSO_TO_OO__Application_AddChartAutoFormat,
    MSO_TO_OO__Application_AddCustomList,
    MSO_TO_OO__Application_get_AlertBeforeOverwriting,
    MSO_TO_OO__Application_put_AlertBeforeOverwriting,
    MSO_TO_OO__Application_get_AltStartupPath,
    MSO_TO_OO__Application_put_AltStartupPath,
    MSO_TO_OO__Application_get_AskToUpdateLinks,
    MSO_TO_OO__Application_put_AskToUpdateLinks,
    MSO_TO_OO__Application_get_EnableAnimations,
    MSO_TO_OO__Application_put_EnableAnimations,
    MSO_TO_OO__Application_get_AutoCorrect,
    MSO_TO_OO__Application_get_Build,
    MSO_TO_OO__Application_get_CalculateBeforeSave,
    MSO_TO_OO__Application_put_CalculateBeforeSave,
    MSO_TO_OO__Application_get_Calculation,
    MSO_TO_OO__Application_put_Calculation,
    MSO_TO_OO__Application_get_Caller,
    MSO_TO_OO__Application_get_CanPlaySounds,
    MSO_TO_OO__Application_get_CanRecordSounds,
    MSO_TO_OO__Application_get_Caption,
    MSO_TO_OO__Application_put_Caption,
    MSO_TO_OO__Application_get_CellDragAndDrop,
    MSO_TO_OO__Application_put_CellDragAndDrop,
    MSO_TO_OO__Application_CentimetersToPoints,
    MSO_TO_OO__Application_CheckSpelling,
    MSO_TO_OO__Application_get_ClipboardFormats,
    MSO_TO_OO__Application_get_DisplayClipboardWindow,
    MSO_TO_OO__Application_put_DisplayClipboardWindow,
    MSO_TO_OO__Application_get_ColorButtons,
    MSO_TO_OO__Application_put_ColorButtons,
    MSO_TO_OO__Application_get_CommandUnderlines,
    MSO_TO_OO__Application_put_CommandUnderlines,
    MSO_TO_OO__Application_get_ConstrainNumeric,
    MSO_TO_OO__Application_put_ConstrainNumeric,
    MSO_TO_OO__Application_ConvertFormula,
    MSO_TO_OO__Application_get_CopyObjectsWithCells,
    MSO_TO_OO__Application_put_CopyObjectsWithCells,
    MSO_TO_OO__Application_get_Cursor,
    MSO_TO_OO__Application_put_Cursor,
    MSO_TO_OO__Application_get_CustomListCount,
    MSO_TO_OO__Application_get_CutCopyMode,
    MSO_TO_OO__Application_put_CutCopyMode,
    MSO_TO_OO__Application_get_DataEntryMode,
    MSO_TO_OO__Application_put_DataEntryMode,
    MSO_TO_OO__Application_Dummy1,
    MSO_TO_OO__Application_Dummy2,
    MSO_TO_OO__Application_Dummy3,
    MSO_TO_OO__Application_Dummy4,
    MSO_TO_OO__Application_Dummy5,
    MSO_TO_OO__Application_Dummy6,
    MSO_TO_OO__Application_Dummy7,
    MSO_TO_OO__Application_Dummy8,
    MSO_TO_OO__Application_Dummy9,
    MSO_TO_OO__Application_Dummy10,
    MSO_TO_OO__Application_Dummy11,
    MSO_TO_OO__Application_get__Default,
    MSO_TO_OO__Application_get_DefaultFilePath,
    MSO_TO_OO__Application_put_DefaultFilePath,
    MSO_TO_OO__Application_DeleteChartAutoFormat,
    MSO_TO_OO__Application_DeleteCustomList,
    MSO_TO_OO__Application_get_Dialogs,
    MSO_TO_OO__Application_put_DisplayAlerts,
    MSO_TO_OO__Application_get_DisplayAlerts,
    MSO_TO_OO__Application_get_DisplayFormulaBar,
    MSO_TO_OO__Application_put_DisplayFormulaBar,
    MSO_TO_OO__Application_get_DisplayFullScreen,
    MSO_TO_OO__Application_put_DisplayFullScreen,
    MSO_TO_OO__Application_get_DisplayNoteIndicator,
    MSO_TO_OO__Application_put_DisplayNoteIndicator,
    MSO_TO_OO__Application_get_DisplayCommentIndicator,
    MSO_TO_OO__Application_put_DisplayCommentIndicator,
    MSO_TO_OO__Application_get_DisplayExcel4Menus,
    MSO_TO_OO__Application_put_DisplayExcel4Menus,
    MSO_TO_OO__Application_get_DisplayRecentFiles,
    MSO_TO_OO__Application_put_DisplayRecentFiles,
    MSO_TO_OO__Application_get_DisplayScrollBars,
    MSO_TO_OO__Application_put_DisplayScrollBars,
    MSO_TO_OO__Application_get_DisplayStatusBar,
    MSO_TO_OO__Application_put_DisplayStatusBar,
    MSO_TO_OO__Application_DoubleClick,
    MSO_TO_OO__Application_get_EditDirectlyInCell,
    MSO_TO_OO__Application_put_EditDirectlyInCell,
    MSO_TO_OO__Application_get_EnableAutoComplete,
    MSO_TO_OO__Application_put_EnableAutoComplete,
    MSO_TO_OO__Application_get_EnableCancelKey,
    MSO_TO_OO__Application_put_EnableCancelKey,
    MSO_TO_OO__Application_get_EnableSound,
    MSO_TO_OO__Application_put_EnableSound,
    MSO_TO_OO__Application_get_EnableTipWizard,
    MSO_TO_OO__Application_put_EnableTipWizard,
    MSO_TO_OO__Application_get_FileConverters,
    MSO_TO_OO__Application_get_FileSearch,
    MSO_TO_OO__Application_get_FileFind,
    MSO_TO_OO__Application__FindFile,
    MSO_TO_OO__Application_get_FixedDecimal,
    MSO_TO_OO__Application_put_FixedDecimal,
    MSO_TO_OO__Application_get_FixedDecimalPlaces,
    MSO_TO_OO__Application_put_FixedDecimalPlaces,
    MSO_TO_OO__Application_GetCustomListContents,
    MSO_TO_OO__Application_GetCustomListNum,
    MSO_TO_OO__Application_GetOpenFilename,
    MSO_TO_OO__Application_GetSaveAsFilename,
    MSO_TO_OO__Application_Goto,
    MSO_TO_OO__Application_get_Height,
    MSO_TO_OO__Application_put_Height,
    MSO_TO_OO__Application_Help,
    MSO_TO_OO__Application_get_IgnoreRemoteRequests,
    MSO_TO_OO__Application_put_IgnoreRemoteRequests,
    MSO_TO_OO__Application_InchesToPoints,
    MSO_TO_OO__Application_InputBox,
    MSO_TO_OO__Application_get_Interactive,
    MSO_TO_OO__Application_put_Interactive,
    MSO_TO_OO__Application_get_International,
    MSO_TO_OO__Application_get_Iteration,
    MSO_TO_OO__Application_put_Iteration,
    MSO_TO_OO__Application_get_LargeButtons,
    MSO_TO_OO__Application_put_LargeButtons,
    MSO_TO_OO__Application_get_Left,
    MSO_TO_OO__Application_put_Left,
    MSO_TO_OO__Application_get_LibraryPath,
    MSO_TO_OO__Application_MacroOptions,
    MSO_TO_OO__Application_MailLogoff,
    MSO_TO_OO__Application_MailLogon,
    MSO_TO_OO__Application_get_MailSession,
    MSO_TO_OO__Application_get_MailSystem,
    MSO_TO_OO__Application_get_MathCoprocessorAvailable,
    MSO_TO_OO__Application_get_MaxChange,
    MSO_TO_OO__Application_put_MaxChange,
    MSO_TO_OO__Application_get_MaxIterations,
    MSO_TO_OO__Application_put_MaxIterations,
    MSO_TO_OO__Application_get_MemoryFree,
    MSO_TO_OO__Application_get_MemoryTotal,
    MSO_TO_OO__Application_get_MemoryUsed,
    MSO_TO_OO__Application_get_MouseAvailable,
    MSO_TO_OO__Application_get_MoveAfterReturn,
    MSO_TO_OO__Application_put_MoveAfterReturn,
    MSO_TO_OO__Application_get_MoveAfterReturnDirection,
    MSO_TO_OO__Application_put_MoveAfterReturnDirection,
    MSO_TO_OO__Application_get_RecentFiles,
    MSO_TO_OO__Application_get_Name,
    MSO_TO_OO__Application_NextLetter,
    MSO_TO_OO__Application_get_NetworkTemplatesPath,
    MSO_TO_OO__Application_get_ODBCErrors,
    MSO_TO_OO__Application_get_ODBCTimeout,
    MSO_TO_OO__Application_put_ODBCTimeout,
    MSO_TO_OO__Application_get_OnCalculate,
    MSO_TO_OO__Application_put_OnCalculate,
    MSO_TO_OO__Application_get_OnData,
    MSO_TO_OO__Application_put_OnData,
    MSO_TO_OO__Application_get_OnDoubleClick,
    MSO_TO_OO__Application_put_OnDoubleClick,
    MSO_TO_OO__Application_get_OnEntry,
    MSO_TO_OO__Application_put_OnEntry,
    MSO_TO_OO__Application_OnKey,
    MSO_TO_OO__Application_OnRepeat,
    MSO_TO_OO__Application_get_OnSheetActivate,
    MSO_TO_OO__Application_put_OnSheetActivate,
    MSO_TO_OO__Application_get_OnSheetDeactivate,
    MSO_TO_OO__Application_put_OnSheetDeactivate,
    MSO_TO_OO__Application_OnTime,
    MSO_TO_OO__Application_OnUndo,
    MSO_TO_OO__Application_get_OnWindow,
    MSO_TO_OO__Application_put_OnWindow,
    MSO_TO_OO__Application_get_OperatingSystem,
    MSO_TO_OO__Application_get_OrganizationName,
    MSO_TO_OO__Application_get_Path,
    MSO_TO_OO__Application_get_PathSeparator,
    MSO_TO_OO__Application_get_PreviousSelections,
    MSO_TO_OO__Application_get_PivotTableSelection,
    MSO_TO_OO__Application_put_PivotTableSelection,
    MSO_TO_OO__Application_get_PromptForSummaryInfo,
    MSO_TO_OO__Application_put_PromptForSummaryInfo,
    MSO_TO_OO__Application_Quit,
    MSO_TO_OO__Application_RecordMacro,
    MSO_TO_OO__Application_get_RecordRelative,
    MSO_TO_OO__Application_get_ReferenceStyle,
    MSO_TO_OO__Application_put_ReferenceStyle,
    MSO_TO_OO__Application_get_RegisteredFunctions,
    MSO_TO_OO__Application_RegisterXLL,
    MSO_TO_OO__Application_Repeat,
    MSO_TO_OO__Application_ResetTipWizard,
    MSO_TO_OO__Application_get_RollZoom,
    MSO_TO_OO__Application_put_RollZoom,
    MSO_TO_OO__Application_Save,
    MSO_TO_OO__Application_SaveWorkspace,
    MSO_TO_OO__Application_get_ScreenUpdating,
    MSO_TO_OO__Application_put_ScreenUpdating,
    MSO_TO_OO__Application_SetDefaultChart,
    MSO_TO_OO__Application_get_SheetsInNewWorkbook,
    MSO_TO_OO__Application_put_SheetsInNewWorkbook,
    MSO_TO_OO__Application_get_ShowChartTipNames,
    MSO_TO_OO__Application_put_ShowChartTipNames,
    MSO_TO_OO__Application_get_ShowChartTipValues,
    MSO_TO_OO__Application_put_ShowChartTipValues,
    MSO_TO_OO__Application_get_StandardFont,
    MSO_TO_OO__Application_put_StandardFont,
    MSO_TO_OO__Application_get_StandardFontSize,
    MSO_TO_OO__Application_put_StandardFontSize,
    MSO_TO_OO__Application_get_StartupPath,
    MSO_TO_OO__Application_get_StatusBar,
    MSO_TO_OO__Application_put_StatusBar,
    MSO_TO_OO__Application_get_TemplatesPath,
    MSO_TO_OO__Application_get_ShowToolTips,
    MSO_TO_OO__Application_put_ShowToolTips,
    MSO_TO_OO__Application_get_Top,
    MSO_TO_OO__Application_put_Top,
    MSO_TO_OO__Application_get_DefaultSaveFormat,
    MSO_TO_OO__Application_put_DefaultSaveFormat,
    MSO_TO_OO__Application_get_TransitionMenuKey,
    MSO_TO_OO__Application_put_TransitionMenuKey,
    MSO_TO_OO__Application_get_TransitionMenuKeyAction,
    MSO_TO_OO__Application_put_TransitionMenuKeyAction,
    MSO_TO_OO__Application_get_TransitionNavigKeys,
    MSO_TO_OO__Application_put_TransitionNavigKeys,
    MSO_TO_OO__Application_Undo,
    MSO_TO_OO__Application_get_UsableHeight,
    MSO_TO_OO__Application_get_UsableWidth,
    MSO_TO_OO__Application_put_UserControl,
    MSO_TO_OO__Application_get_UserControl,
    MSO_TO_OO__Application_get_UserName,
    MSO_TO_OO__Application_put_UserName,
    MSO_TO_OO__Application_get_Value,
    MSO_TO_OO__Application_get_VBE,
    MSO_TO_OO__Application_get_Version,
    MSO_TO_OO__Application_get_Visible,
    MSO_TO_OO__Application_put_Visible,
    MSO_TO_OO__Application_Volatile,
    MSO_TO_OO__Application__Wait,
    MSO_TO_OO__Application_get_Width,
    MSO_TO_OO__Application_put_Width,
    MSO_TO_OO__Application_get_WindowsForPens,
    MSO_TO_OO__Application_put_WindowState,
    MSO_TO_OO__Application_get_WindowState,
    MSO_TO_OO__Application_get_UILanguage,
    MSO_TO_OO__Application_put_UILanguage,
    MSO_TO_OO__Application_get_DefaultSheetDirection,
    MSO_TO_OO__Application_put_DefaultSheetDirection,
    MSO_TO_OO__Application_get_CursorMovement,
    MSO_TO_OO__Application_put_CursorMovement,
    MSO_TO_OO__Application_get_ControlCharacters,
    MSO_TO_OO__Application_put_ControlCharacters,
    MSO_TO_OO__Application__WSFunction,
    MSO_TO_OO__Application_get_EnableEvents,
    MSO_TO_OO__Application_put_EnableEvents,
    MSO_TO_OO__Application_get_DisplayInfoWindow,
    MSO_TO_OO__Application_put_DisplayInfoWindow,
    MSO_TO_OO__Application_Wait,
    MSO_TO_OO__Application_get_ExtendList,
    MSO_TO_OO__Application_put_ExtendList,
    MSO_TO_OO__Application_get_OLEDBErrors,
    MSO_TO_OO__Application_GetPhonetic,
    MSO_TO_OO__Application_get_COMAddIns,
    MSO_TO_OO__Application_get_DefaultWebOptions,
    MSO_TO_OO__Application_get_ProductCode,
    MSO_TO_OO__Application_get_UserLibraryPath,
    MSO_TO_OO__Application_get_AutoPercentEntry,
    MSO_TO_OO__Application_put_AutoPercentEntry,
    MSO_TO_OO__Application_get_LanguageSettings,
    MSO_TO_OO__Application_get_Dummy101,
    MSO_TO_OO__Application_Dummy12,
    MSO_TO_OO__Application_get_AnswerWizard,
    MSO_TO_OO__Application_CalculateFull,
    MSO_TO_OO__Application_FindFile,
    MSO_TO_OO__Application_get_CalculationVersion,
    MSO_TO_OO__Application_get_ShowWindowsInTaskbar,
    MSO_TO_OO__Application_put_ShowWindowsInTaskbar,
    MSO_TO_OO__Application_get_FeatureInstall,
    MSO_TO_OO__Application_put_FeatureInstall,
    MSO_TO_OO__Application_get_Ready,
    MSO_TO_OO__Application_Dummy13,
    MSO_TO_OO__Application_get_FindFormat,
    MSO_TO_OO__Application_putref_FindFormat,
    MSO_TO_OO__Application_get_ReplaceFormat,
    MSO_TO_OO__Application_putref_ReplaceFormat,
    MSO_TO_OO__Application_get_UsedObjects,
    MSO_TO_OO__Application_get_CalculationState,
    MSO_TO_OO__Application_get_CalculationInterruptKey,
    MSO_TO_OO__Application_put_CalculationInterruptKey,
    MSO_TO_OO__Application_get_Watches,
    MSO_TO_OO__Application_get_DisplayFunctionToolTips,
    MSO_TO_OO__Application_put_DisplayFunctionToolTips,
    MSO_TO_OO__Application_get_AutomationSecurity,
    MSO_TO_OO__Application_put_AutomationSecurity,
    MSO_TO_OO__Application_get_FileDialog,
    MSO_TO_OO__Application_Dummy14,
    MSO_TO_OO__Application_CalculateFullRebuild,
    MSO_TO_OO__Application_get_DisplayPasteOptions,
    MSO_TO_OO__Application_put_DisplayPasteOptions,
    MSO_TO_OO__Application_get_DisplayInsertOptions,
    MSO_TO_OO__Application_put_DisplayInsertOptions,
    MSO_TO_OO__Application_get_GenerateGetPivotData,
    MSO_TO_OO__Application_put_GenerateGetPivotData,
    MSO_TO_OO__Application_get_AutoRecover,
    MSO_TO_OO__Application_get_Hwnd,
    MSO_TO_OO__Application_get_Hinstance,
    MSO_TO_OO__Application_CheckAbort,
    MSO_TO_OO__Application_get_ErrorCheckingOptions,
    MSO_TO_OO__Application_get_AutoFormatAsYouTypeReplaceHyperlinks,
    MSO_TO_OO__Application_put_AutoFormatAsYouTypeReplaceHyperlinks,
    MSO_TO_OO__Application_get_SmartTagRecognizers,
    MSO_TO_OO__Application_get_NewWorkbook,
    MSO_TO_OO__Application_get_SpellingOptions,
    MSO_TO_OO__Application_get_Speech,
    MSO_TO_OO__Application_get_MapPaperSize,
    MSO_TO_OO__Application_put_MapPaperSize,
    MSO_TO_OO__Application_get_ShowStartupDialog,
    MSO_TO_OO__Application_put_ShowStartupDialog,
    MSO_TO_OO__Application_get_DecimalSeparator,
    MSO_TO_OO__Application_put_DecimalSeparator,
    MSO_TO_OO__Application_get_ThousandsSeparator,
    MSO_TO_OO__Application_put_ThousandsSeparator,
    MSO_TO_OO__Application_get_UseSystemSeparators,
    MSO_TO_OO__Application_put_UseSystemSeparators,
    MSO_TO_OO__Application_get_ThisCell,
    MSO_TO_OO__Application_get_RTD,
    MSO_TO_OO__Application_get_DisplayDocumentActionTaskPane,
    MSO_TO_OO__Application_put_DisplayDocumentActionTaskPane,
    MSO_TO_OO__Application_DisplayXMLSourcePane,
    MSO_TO_OO__Application_get_ArbitraryXMLSupportAvailable,
    MSO_TO_OO__Application_Support,
};

HRESULT _ApplicationConstructor(LPVOID *ppObj)
{
    _ApplicationImpl *_applicationexcell;
    CLSID clsid;
    HRESULT hres;
    VARIANT result;
    VARIANT param1;
    IUnknown *punk = NULL;
    TRACE_IN;
    TRACE("(%p) \n", ppObj);

    _applicationexcell = HeapAlloc(GetProcessHeap(), 0, sizeof(*_applicationexcell));
    if (!_applicationexcell) {
        return E_OUTOFMEMORY;
    }

    _applicationexcell->pApplicationVtbl = &MSO_TO_OO__Application_Vtbl;
    _applicationexcell->pConnectionPointContainerVtbl = &MSO_TO_OO_ConnectionPointContainerVtbl;
    _applicationexcell->pConnectionPointVtbl = &MSO_TO_OO_ConnectionPointVtbl;
    _applicationexcell->ref = 0;
    _applicationexcell->pdOOApp = NULL;
    _applicationexcell->pdOODesktop = NULL;
    _applicationexcell->pdWorkbooks = NULL;
    _applicationexcell->screenupdating = VARIANT_TRUE;
    _applicationexcell->visible = VARIANT_FALSE;
    _applicationexcell->sheetsinnewworkbook = 1;
    _applicationexcell->displayalerts = VARIANT_TRUE;
    /*Создание указателей на объекты openOfffice 
    Create OpenOffice Service Manager */
    hres = CLSIDFromProgID(L"com.sun.star.ServiceManager", &clsid);
    if (FAILED(hres)) {
        TRACE("ERROR when  CLSIDFromProgID  com.sun.star.ServiceManager \n");
        return E_NOINTERFACE;
    }

    /* Start server and get IDispatch...*/
    hres = CoCreateInstance(&clsid, NULL, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER, &IID_IDispatch, (void**) &(_applicationexcell->pdOOApp));
    if (FAILED(hres)) {
        TRACE("ERROR when CoCreateInstance \n");
        return E_NOINTERFACE;
    }

    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"com.sun.star.frame.Desktop");
    /* Get Desktop and its assoc. IDispatch...*/
    hres = AutoWrap(DISPATCH_METHOD, &result, _applicationexcell->pdOOApp, L"CreateInstance", 1, param1);

    if (FAILED(hres)) {
        TRACE("ERROR when CreateInstance \n");
        return E_NOINTERFACE;
    }

    _applicationexcell->pdOODesktop = result.pdispVal;
    hres = IDispatch_AddRef(_applicationexcell->pdOODesktop);

    hres = _I_WorkbooksConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Workbooks_QueryInterface(punk, &IID_I_Workbooks, (void**) &(_applicationexcell->pdWorkbooks));
    if (FAILED(hres)) return E_NOINTERFACE;
/*    I_Workbooks_Release(punk);*/

    MSO_TO_OO_I_Workbooks_Initialize((I_Workbooks*)(_applicationexcell->pdWorkbooks), (_Application*)_applicationexcell);

    *ppObj = APPEXCEL(_applicationexcell);
/*Пытаемся получить номер версии*/
    VARIANT p1,p2, param2, conf_prov, access_prov, version, res;
    IDispatch *dpv;
    long index = 0;
    SAFEARRAY FAR* pPropVals;

    VariantInit(&p1);
    VariantInit(&p2);
    VariantInit(&param2);
    VariantInit(&conf_prov);
    VariantInit(&access_prov);
    VariantInit(&version);
    VariantInit(&res);

    VariantClear(&param1);
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"com.sun.star.configuration.ConfigurationProvider");

    hres = AutoWrap(DISPATCH_METHOD, &conf_prov, _applicationexcell->pdOOApp, L"CreateInstance", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when CreateInstance ------- ConfigurationProvider\n");
        return S_OK;
    }

    MSO_TO_OO_GetDispatchPropertyValue(APPEXCEL(_applicationexcell), &dpv);
    if (dpv == NULL) {
        TRACE("ERROR when GetDispatchPropertyValue\n");
        return S_OK;
    }

    V_VT(&p1) = VT_BSTR;
    V_BSTR(&p1) = SysAllocString(L"nodepath");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Name", 1, p1);
    V_VT(&p2) = VT_BSTR;
    V_BSTR(&p2) = SysAllocString(L"/org.openoffice.Setup/Product");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Value", 1, p2);

    pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 1 );
    hres = SafeArrayPutElement( pPropVals, &index, dpv );
    V_VT(&param2) = VT_ARRAY | VT_DISPATCH;
    V_ARRAY(&param2) = pPropVals;

    VariantClear(&param1);
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"com.sun.star.configuration.ConfigurationAccess");
    hres = AutoWrap(DISPATCH_METHOD, &access_prov, V_DISPATCH(&conf_prov), L"createInstanceWithArguments", 2, param2, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when CreateInstance --- ConfigurationAccess \n");
        return S_OK;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"ooSetupVersion");
    hres = AutoWrap(DISPATCH_METHOD, &version, V_DISPATCH(&access_prov), L"getByName", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when getByName \n");
        return S_OK;
    }

    if (*(V_BSTR(&version))==L'2') OOVersion = VER_2;
    else if (*(V_BSTR(&version))==L'3') OOVersion = VER_3;

    VariantClear(&p1);
    VariantClear(&p2);
    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&conf_prov);
    VariantClear(&access_prov);
    VariantClear(&version);
    VariantClear(&res);

    /*освобождаем память выделенную под строку*/
    SysFreeString(V_BSTR(&param1));
    VariantClear(&result);

    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}


