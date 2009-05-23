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
static HRESULT WINAPI IMPConnectionPoint_QueryInterface(
        IConnectionPoint* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationImpl *This = CONPOINT_THIS(iface);
    return _Application_QueryInterface(APPEXCEL(This), riid, ppvObject);
}

static ULONG WINAPI IMPConnectionPoint_AddRef(
        IConnectionPoint* iface)
{
    _ApplicationImpl *This = CONPOINT_THIS(iface);
    return _Application_AddRef(APPEXCEL(This));
}

static ULONG WINAPI IMPConnectionPoint_Release(
        IConnectionPoint* iface)
{
    _ApplicationImpl *This = CONPOINT_THIS(iface);
    return _Application_Release(APPEXCEL(This));
}

    /*** IConnectionPoint methods ***/
static HRESULT WINAPI IMPConnectionPoint_GetConnectionInterface(
        IConnectionPoint* iface,
        IID *pIID)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI IMPConnectionPoint_GetConnectionPointContainer(
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

static HRESULT WINAPI IMPConnectionPoint_Advise(
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

static HRESULT WINAPI IMPConnectionPoint_Unadvise(
        IConnectionPoint* iface,
        DWORD dwCookie)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI IMPConnectionPoint_EnumConnections(
        IConnectionPoint* iface,
        IEnumConnections **ppEnum)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

const IConnectionPointVtbl IMPConnectionPointVtbl = 
{
    IMPConnectionPoint_QueryInterface,
    IMPConnectionPoint_AddRef,
    IMPConnectionPoint_Release,
    IMPConnectionPoint_GetConnectionInterface,
    IMPConnectionPoint_GetConnectionPointContainer,
    IMPConnectionPoint_Advise,
    IMPConnectionPoint_Unadvise,
    IMPConnectionPoint_EnumConnections
};

#undef CONPOINT_THIS

/*IConnectionPointContainer interface*/

#define CONPOINTCONT_THIS(iface) DEFINE_THIS(_ApplicationImpl,ConnectionPointContainer,iface);

    /*** IUnknown methods ***/
static HRESULT WINAPI IMPConnectionPointContainer_QueryInterface(
        IConnectionPointContainer* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationImpl *This = CONPOINTCONT_THIS(iface);
    return _Application_QueryInterface(APPEXCEL(This), riid, ppvObject);
}

static ULONG WINAPI IMPConnectionPointContainer_AddRef(
        IConnectionPointContainer* iface)
{
    _ApplicationImpl *This = CONPOINTCONT_THIS(iface);
    return _Application_AddRef(APPEXCEL(This));
}

static ULONG WINAPI IMPConnectionPointContainer_Release(
        IConnectionPointContainer* iface)
{
    _ApplicationImpl *This = CONPOINTCONT_THIS(iface);
    return _Application_Release(APPEXCEL(This));
}

    /*** IConnectionPointContainer methods ***/
static HRESULT WINAPI IMPConnectionPointContainer_EnumConnectionPoints(
        IConnectionPointContainer* iface,
        IEnumConnectionPoints **ppEnum)
{
    TRACE("Not implemented \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI IMPConnectionPointContainer_FindConnectionPoint(
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

const IConnectionPointContainerVtbl IMPConnectionPointContainerVtbl = 
{
    IMPConnectionPointContainer_QueryInterface,
    IMPConnectionPointContainer_AddRef,
    IMPConnectionPointContainer_Release,
    IMPConnectionPointContainer_EnumConnectionPoints,
    IMPConnectionPointContainer_FindConnectionPoint
};

#undef CONPOINTCONT_THIS

/*_Application interface*/

/*
IUnknown
*/

#define APPEXCEL_THIS(iface) DEFINE_THIS(_ApplicationImpl, Application, iface);

static ULONG WINAPI IMP_Application_AddRef(
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

static HRESULT WINAPI IMP_Application_QueryInterface(
        _Application* iface,
        REFIID riid,
        void **ppvObject)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    WCHAR str_clsid[39];

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    *ppvObject = NULL;

    if (IsEqualGUID(riid, &IID_IUnknown)) {
        TRACE("IUnknown \n");
        *ppvObject = DISPAPPEXCEL(This);
    }
    
    if (IsEqualGUID(riid, &IID_IDispatch)) {
        TRACE("IDispatch \n");
        *ppvObject = APPEXCEL(This);
    }
    
    if ( IsEqualGUID(riid, &IID__Application)) {
        TRACE("_Application \n");
        *ppvObject = DISPAPPEXCEL(This);
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

static ULONG WINAPI IMP_Application_Release(
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
static HRESULT WINAPI IMP_Application_put_UserControl(
        _Application* iface,
        VARIANT_BOOL vbUserControl)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_get_UserControl(
        _Application* iface,
        VARIANT_BOOL *vbUserControl)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_put_DisplayAlerts(
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

static HRESULT WINAPI IMP_Application_get_DisplayAlerts(
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

static HRESULT WINAPI IMP_Application_put_WindowState(
        _Application* iface,
        LCID lcid,
        XlWindowState State)
{
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_get_WindowState(
        _Application* iface,
        LCID lcid,
        XlWindowState *State)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Visible(
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

static HRESULT WINAPI IMP_Application_get_Visible(
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

static HRESULT WINAPI IMP_Application_get_Workbooks(
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

static HRESULT WINAPI IMP_Application_get_Sheets(
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

static HRESULT WINAPI IMP_Application_get_Worksheets(
        _Application* iface,
        IDispatch **ppSheets)
{
   /*Используем Sheets - они выполняют одинаковые функции*/
   TRACE(" ----> get_Sheets");
   return IMP_Application_get_Sheets(iface, ppSheets);
}

static HRESULT WINAPI IMP_Application_get_Cells(
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

static HRESULT WINAPI IMP_Application_get_ActiveSheet(
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

static HRESULT WINAPI IMP_Application_get_Version(
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

static HRESULT WINAPI IMP_Application_ConvertFormula(
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

static HRESULT WINAPI IMP_Application_Quit(
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

static HRESULT WINAPI IMP_Application_get_ActiveCell(
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

static HRESULT WINAPI IMP_Application_get_Application(
        _Application* iface,
        IDispatch **value)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE_IN;

    if (iface!=NULL) {
        *value = (IDispatch*)APPEXCEL(This);
        IMP_Application_AddRef((_Application*)*value);
    } else {
        return E_FAIL;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_get_EnableEvents(
        _Application* iface,
        VARIANT_BOOL *pvbee)
{
    TRACE_IN;
    /*Always return TRUE*/
    *pvbee = VARIANT_TRUE;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_put_EnableEvents(
        _Application* iface,
        VARIANT_BOOL vbee)
{
    TRACE_NOTIMPL;
    /*Always return S_OK*/
    return S_OK;
}

static HRESULT WINAPI IMP_Application_put_ScreenUpdating(
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

static HRESULT WINAPI IMP_Application_get_ScreenUpdating(
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

static HRESULT WINAPI IMP_Application_get_Caption(
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

static HRESULT WINAPI IMP_Application_put_Caption(
        _Application* iface,
        VARIANT vName)
{
    TRACE_NOTIMPL;
    MSO_TO_OO_CorrectArg(vName, &vName);
    return S_OK;
}

static HRESULT WINAPI IMP_Application_get_ActiveWorkbook(
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

static HRESULT WINAPI IMP_Application_get_Range(
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

    hres = IMP_Application_get_ActiveSheet(iface, (IDispatch**) &wsh);

    hres = I_Worksheet_get_Range(wsh,Cell1, Cell2, ppRange);

    I_Worksheet_Release(wsh);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI IMP_Application_get_Columns(
        _Application* iface,
        VARIANT param,
        IDispatch **ppRange)
{
    HRESULT hres;
    IDispatch *active_sheet;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    hres = IMP_Application_get_ActiveSheet(iface, &active_sheet);

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

static HRESULT WINAPI IMP_Application_get_Rows(
        _Application* iface,
        VARIANT param,
        IDispatch **ppRange)
{
    HRESULT hres;
    IDispatch *active_sheet;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    hres = IMP_Application_get_ActiveSheet(iface, &active_sheet);

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

static HRESULT WINAPI IMP_Application_get_Selection(
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

    hres = IMP_Application_get_ActiveWorkbook(iface, (IDispatch**)&awb);
    if (FAILED(hres)) {
        TRACE("ERROR when get_ActiveWorkbook\n");
        return E_FAIL;
    }

    hres = IMP_Application_get_ActiveSheet(iface, &asheet);
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

static HRESULT WINAPI IMP_Application_get_Creator(
        _Application* iface,
        XlCreator *RHS)
{
    TRACE_IN;
    *RHS = xlCreatorCode;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_get_Parent(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ActiveChart(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ActiveDialog(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ActiveMenuBar(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ActivePrinter(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ActivePrinter(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ActiveWindow(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AddIns(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Assistant(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Calculate(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Charts(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CommandBars(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DDEAppReturnCode(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DDEExecute(
        _Application* iface,
        long Channel,
        BSTR String,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DDEInitiate(
        _Application* iface,
        BSTR App,
        BSTR Topic,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DDEPoke(
        _Application* iface,
        long Channel,
        VARIANT Item,
        VARIANT Data,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DDERequest(
        _Application* iface,
        long Channel,
        BSTR Item,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DDETerminate(
        _Application* iface,
        long Channel,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DialogSheets(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Evaluate(
        _Application* iface,
        VARIANT Name,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application__Evaluate(
        _Application* iface,
        VARIANT Name,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_ExecuteExcel4Macro(
        _Application* iface,
        BSTR String,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Intersect(
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

static HRESULT WINAPI IMP_Application_get_MenuBars(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Modules(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Names(
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

static HRESULT WINAPI IMP_Application_Run(
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

static HRESULT WINAPI IMP_Application__Run2(
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

static HRESULT WINAPI IMP_Application_SendKeys(
        _Application* iface,
        VARIANT Keys,
        VARIANT Wait,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ShortcutMenus(
        _Application* iface,
        long Index,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ThisWorkbook(
        _Application* iface,
        LCID lcid,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Toolbars(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Union(
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

static HRESULT WINAPI IMP_Application_get_Windows(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_WorksheetFunction(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Worksheets(
        _Application* iface,
        IDispatch **ppSheets);

static HRESULT WINAPI IMP_Application_get_Excel4IntlMacroSheets(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Excel4MacroSheets(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_ActivateMicrosoftApp(
        _Application* iface,
        XlMSApplication Index,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_AddChartAutoFormat(
        _Application* iface,
        VARIANT Chart,
        BSTR Name,
        VARIANT Description,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_AddCustomList(
        _Application* iface,
        VARIANT ListArray,
        VARIANT ByRow,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AlertBeforeOverwriting(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_AlertBeforeOverwriting(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AltStartupPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_AltStartupPath(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AskToUpdateLinks(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_AskToUpdateLinks(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_EnableAnimations(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_EnableAnimations(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AutoCorrect(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Build(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CalculateBeforeSave(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_CalculateBeforeSave(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Calculation(
        _Application* iface,
        LCID lcid,
        XlCalculation *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Calculation(
        _Application* iface,
        LCID lcid,
        XlCalculation RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Caller(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CanPlaySounds(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CanRecordSounds(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CellDragAndDrop(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_CellDragAndDrop(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_CentimetersToPoints(
        _Application* iface,
        double Centimeters,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_CheckSpelling(
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

static HRESULT WINAPI IMP_Application_get_ClipboardFormats(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayClipboardWindow(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayClipboardWindow(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ColorButtons(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ColorButtons(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CommandUnderlines(
        _Application* iface,
        LCID lcid,
        XlCommandUnderlines *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_CommandUnderlines(
        _Application* iface,
        LCID lcid,
        XlCommandUnderlines RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ConstrainNumeric(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ConstrainNumeric(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CopyObjectsWithCells(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_CopyObjectsWithCells(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Cursor(
        _Application* iface,
        LCID lcid,
        XlMousePointer *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Cursor(
        _Application* iface,
        LCID lcid,
        XlMousePointer RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CustomListCount(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CutCopyMode(
        _Application* iface,
        LCID lcid,
        XlCutCopyMode *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_CutCopyMode(
        _Application* iface,
        LCID lcid,
        XlCutCopyMode RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}
static HRESULT WINAPI IMP_Application_get_DataEntryMode(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DataEntryMode(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy1(
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

static HRESULT WINAPI IMP_Application_Dummy2(
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

static HRESULT WINAPI IMP_Application_Dummy3(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy4(
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

static HRESULT WINAPI IMP_Application_Dummy5(
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

static HRESULT WINAPI IMP_Application_Dummy6(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy7(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy8(
        _Application* iface,
        VARIANT Arg1,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy9(
        _Application* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy10(
        _Application* iface,
        VARIANT arg,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy11(
        _Application* This)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get__Default(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DefaultFilePath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DefaultFilePath(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DeleteChartAutoFormat(
        _Application* iface,
        BSTR Name,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DeleteCustomList(
        _Application* iface,
        long ListNum,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Dialogs(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayFormulaBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayFormulaBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayFullScreen(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayFullScreen(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayNoteIndicator(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayNoteIndicator(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayCommentIndicator(
        _Application* iface,
        XlCommentDisplayMode *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayCommentIndicator(
        _Application* iface,
        XlCommentDisplayMode RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayExcel4Menus(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayExcel4Menus(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayRecentFiles(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayRecentFiles(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayScrollBars(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayScrollBars(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayStatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayStatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DoubleClick(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_EditDirectlyInCell(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_EditDirectlyInCell(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_EnableAutoComplete(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_EnableAutoComplete(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}
static HRESULT WINAPI IMP_Application_get_EnableCancelKey(
        _Application* iface,
        LCID lcid,
        XlEnableCancelKey *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_EnableCancelKey(
        _Application* iface,
        LCID lcid,
        XlEnableCancelKey RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_EnableSound(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_EnableSound(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_EnableTipWizard(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_EnableTipWizard(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_FileConverters(
        _Application* iface,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_FileSearch(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_FileFind(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application__FindFile(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_FixedDecimal(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_FixedDecimal(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_FixedDecimalPlaces(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_FixedDecimalPlaces(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_GetCustomListContents(
        _Application* iface,
        long ListNum,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_GetCustomListNum(
        _Application* iface,
        VARIANT ListArray,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_GetOpenFilename(
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

static HRESULT WINAPI IMP_Application_GetSaveAsFilename(
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

static HRESULT WINAPI IMP_Application_Goto(
        _Application* iface,
        VARIANT Reference,
        VARIANT Scroll,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Height(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Height(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Help(
        _Application* iface,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_IgnoreRemoteRequests(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_IgnoreRemoteRequests(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_InchesToPoints(
        _Application* iface,
        double Inches,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_InputBox(
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

static HRESULT WINAPI IMP_Application_get_Interactive(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Interactive(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_International(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Iteration(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Iteration(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_LargeButtons(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_LargeButtons(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Left(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Left(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_LibraryPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_MacroOptions(
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

static HRESULT WINAPI IMP_Application_MailLogoff(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_MailLogon(
        _Application* iface,
        VARIANT Name,
        VARIANT Password,
        VARIANT DownloadNewMail,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MailSession(
        _Application* iface,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MailSystem(
        _Application* iface,
        LCID lcid,
        XlMailSystem *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MathCoprocessorAvailable(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MaxChange(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_MaxChange(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MaxIterations(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_MaxIterations(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MemoryFree(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MemoryTotal(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MemoryUsed(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MouseAvailable(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MoveAfterReturn(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_MoveAfterReturn(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MoveAfterReturnDirection(
        _Application* iface,
        LCID lcid,
        XlDirection *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_MoveAfterReturnDirection(
        _Application* iface,
        LCID lcid,
        XlDirection RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_RecentFiles(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Name(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_NextLetter(
        _Application* iface,
        LCID lcid,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_NetworkTemplatesPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ODBCErrors(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ODBCTimeout(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ODBCTimeout(
        _Application* iface,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OnCalculate(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_OnCalculate(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OnData(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_OnData(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OnDoubleClick(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_OnDoubleClick(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OnEntry(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_OnEntry(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_OnKey(
        _Application* iface,
        BSTR Key,
        VARIANT Procedure,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_OnRepeat(
        _Application* iface,
        BSTR Text,
        BSTR Procedure,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OnSheetActivate(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_OnSheetActivate(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OnSheetDeactivate(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_OnSheetDeactivate(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_OnTime(
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

static HRESULT WINAPI IMP_Application_OnUndo(
        _Application* iface,
        BSTR Text,
        BSTR Procedure,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OnWindow(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_OnWindow(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OperatingSystem(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OrganizationName(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Path(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_PathSeparator(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_PreviousSelections(
        _Application* iface,
        VARIANT Index,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_PivotTableSelection(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_PivotTableSelection(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_PromptForSummaryInfo(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_PromptForSummaryInfo(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_RecordMacro(
        _Application* iface,
        VARIANT BasicCode,
        VARIANT XlmCode,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_RecordRelative(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ReferenceStyle(
        _Application* iface,
        LCID lcid,
        XlReferenceStyle *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ReferenceStyle(
        _Application* iface,
        LCID lcid,
        XlReferenceStyle RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_RegisteredFunctions(
        _Application* iface,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_RegisterXLL(
        _Application* iface,
        BSTR Filename,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Repeat(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_ResetTipWizard(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_RollZoom(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_RollZoom(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Save(
        _Application* iface,
        VARIANT Filename,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_SaveWorkspace(
        _Application* iface,
        VARIANT Filename,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_SetDefaultChart(
        _Application* iface,
        VARIANT FormatName,
        VARIANT Gallery)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_SheetsInNewWorkbook(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE("\n");
    *RHS = This->sheetsinnewworkbook;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_put_SheetsInNewWorkbook(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    _ApplicationImpl *This = APPEXCEL_THIS(iface);
    TRACE("\n");
    This->sheetsinnewworkbook = RHS;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_get_ShowChartTipNames(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ShowChartTipNames(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ShowChartTipValues(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ShowChartTipValues(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_StandardFont(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_StandardFont(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_StandardFontSize(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_StandardFontSize(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_StartupPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_StatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_StatusBar(
        _Application* iface,
        LCID lcid,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_TemplatesPath(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ShowToolTips(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ShowToolTips(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Top(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Top(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DefaultSaveFormat(
        _Application* iface,
        XlFileFormat *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DefaultSaveFormat(
        _Application* iface,
        XlFileFormat RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_TransitionMenuKey(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_TransitionMenuKey(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_TransitionMenuKeyAction(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_TransitionMenuKeyAction(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_TransitionNavigKeys(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_TransitionNavigKeys(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Undo(
        _Application* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_UsableHeight(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_UsableWidth(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_UserName(
        _Application* iface,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_UserName(
        _Application* iface,
        LCID lcid,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Value(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_VBE(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Volatile(
        _Application* iface,
        VARIANT Volatile,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application__Wait(
        _Application* iface,
        VARIANT Time,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Width(
        _Application* iface,
        LCID lcid,
        double *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_Width(
        _Application* iface,
        LCID lcid,
        double RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_WindowsForPens(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_UILanguage(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_UILanguage(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DefaultSheetDirection(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DefaultSheetDirection(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CursorMovement(
        _Application* iface,
        LCID lcid,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_CursorMovement(
        _Application* iface,
        LCID lcid,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ControlCharacters(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ControlCharacters(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application__WSFunction(
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

static HRESULT WINAPI IMP_Application_get_DisplayInfoWindow(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayInfoWindow(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Wait(
        _Application* iface,
        VARIANT Time,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ExtendList(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ExtendList(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_OLEDBErrors(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_GetPhonetic(
        _Application* iface,
        VARIANT Text,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_COMAddIns(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DefaultWebOptions(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ProductCode(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_UserLibraryPath(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AutoPercentEntry(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_AutoPercentEntry(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_LanguageSettings(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Dummy101(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy12(
        _Application* iface,
        IDispatch *p1,
        IDispatch *p2)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AnswerWizard(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_CalculateFull(
        _Application* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_FindFile(
        _Application* iface,
        LCID lcid,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CalculationVersion(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ShowWindowsInTaskbar(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ShowWindowsInTaskbar(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_FeatureInstall(
        _Application* iface,
        MsoFeatureInstall *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_FeatureInstall(
        _Application* iface,
        MsoFeatureInstall RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Ready(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy13(
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

static HRESULT WINAPI IMP_Application_get_FindFormat(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_putref_FindFormat(
        _Application* iface,
        IDispatch *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ReplaceFormat(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_putref_ReplaceFormat(
        _Application* iface,
        IDispatch *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_UsedObjects(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CalculationState(
        _Application* iface,
        XlCalculationState *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_CalculationInterruptKey(
        _Application* iface,
        XlCalculationInterruptKey *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_CalculationInterruptKey(
        _Application* iface,
        XlCalculationInterruptKey RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Watches(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayFunctionToolTips(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayFunctionToolTips(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AutomationSecurity(
        _Application* iface,
        MsoAutomationSecurity *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_AutomationSecurity(
        _Application* iface,
        MsoAutomationSecurity RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_FileDialog(
        _Application* iface,
        MsoFileDialogType fileDialogType,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Dummy14(
        _Application* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_CalculateFullRebuild(
        _Application* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayPasteOptions(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayPasteOptions(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayInsertOptions(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
};

static HRESULT WINAPI IMP_Application_put_DisplayInsertOptions(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_GenerateGetPivotData(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_GenerateGetPivotData(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AutoRecover(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Hwnd(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Hinstance(
        _Application* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_CheckAbort(
        _Application* iface,
        VARIANT KeepAbort)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ErrorCheckingOptions(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_AutoFormatAsYouTypeReplaceHyperlinks(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_AutoFormatAsYouTypeReplaceHyperlinks(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_SmartTagRecognizers(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_NewWorkbook(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_SpellingOptions(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_Speech(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_MapPaperSize(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_MapPaperSize(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ShowStartupDialog(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ShowStartupDialog(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DecimalSeparator(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DecimalSeparator(
        _Application* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ThousandsSeparator(
        _Application* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_ThousandsSeparator(
        _Application* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_UseSystemSeparators(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_UseSystemSeparators(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ThisCell(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_RTD(
        _Application* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_DisplayDocumentActionTaskPane(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_put_DisplayDocumentActionTaskPane(
        _Application* iface,
        VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_DisplayXMLSourcePane(
        _Application* iface,
        VARIANT XmlMap)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_get_ArbitraryXMLSupportAvailable(
        _Application* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI IMP_Application_Support(
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
static HRESULT WINAPI IMP_Application_GetTypeInfoCount(
        _Application* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI IMP_Application_GetTypeInfo(
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

static HRESULT WINAPI IMP_Application_GetIDsOfNames(
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

static HRESULT WINAPI IMP_Application_Invoke(
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
                    hres = IMP_Application_get_Range(iface,pDispParams->rgvarg[1], vNull, &pretdisp);
                    if (FAILED(hres)) {
                        pExcepInfo->bstrDescription=SysAllocString(str_error);
                        TRACE("(case 2) ERROR get_range hres = %08x\n",hres);
                        return hres;
                    }
                    I_Range_put_Value((I_Range*)pretdisp, vNull, 0, pDispParams->rgvarg[0]);
                    IDispatch_Release(pretdisp);
                    return S_OK;
                case 3:
                    hres = IMP_Application_get_Range(iface,pDispParams->rgvarg[2], pDispParams->rgvarg[1], &pretdisp);
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
            return IMP_Application_put_DisplayAlerts(iface, 0, vbin);
        } else {
            hres = IMP_Application_get_DisplayAlerts(iface, 0, &vbin);
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
            IMP_Application_put_WindowState(iface, 0, tmp);
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
            return IMP_Application_put_Visible(iface, 0, vbin);
        } else {
            hres = IMP_Application_get_Visible(iface, 0, &vbin);
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
                hres = IMP_Application_get_Cells(iface,&pdisp);
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
            hres = IMP_Application_get_Cells(iface,&pdisp);
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
            hres = IMP_Application_get_Version(iface,0,&pVersion);
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
        /*IMP_Application_ConvertFormula*/
        if (pDispParams->cArgs<3) return E_FAIL;

        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[pDispParams->cArgs-2]), 0, 0, VT_I4);
        if (FAILED(hres)) {
            TRACE(" (case 11) ERROR when VariantChangeTypeEx\n");
            return E_FAIL;
        }
        tmp = V_I4(&vtmp);

        hres = IMP_Application_ConvertFormula(iface, pDispParams->rgvarg[pDispParams->cArgs-1], tmp, pDispParams->rgvarg[pDispParams->cArgs-3], vNull, vNull, tmp, &vRet);
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
           return IMP_Application_put_EnableEvents(iface, vbin);
        } else {
            hres = IMP_Application_get_EnableEvents(iface,&vbin);
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
            return IMP_Application_put_ScreenUpdating(iface, 0, vbin);
        } else {
            hres = IMP_Application_get_ScreenUpdating(iface, 0, &vbin);
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
            return IMP_Application_put_SheetsInNewWorkbook(iface, 0, tmp);
        } else {
            hres = IMP_Application_get_SheetsInNewWorkbook(iface, 0, &tmp);
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


const _ApplicationVtbl IMP_Application_Vtbl =
{
    IMP_Application_QueryInterface,
    IMP_Application_AddRef,
    IMP_Application_Release,
    IMP_Application_GetTypeInfoCount,
    IMP_Application_GetTypeInfo,
    IMP_Application_GetIDsOfNames,
    IMP_Application_Invoke,
    IMP_Application_get_Application,
    IMP_Application_get_Creator,
    IMP_Application_get_Parent,
    IMP_Application_get_ActiveCell,
    IMP_Application_get_ActiveChart,
    IMP_Application_get_ActiveDialog,
    IMP_Application_get_ActiveMenuBar,
    IMP_Application_get_ActivePrinter,
    IMP_Application_put_ActivePrinter,
    IMP_Application_get_ActiveSheet,
    IMP_Application_get_ActiveWindow,
    IMP_Application_get_ActiveWorkbook,
    IMP_Application_get_AddIns,
    IMP_Application_get_Assistant,
    IMP_Application_Calculate,
    IMP_Application_get_Cells,
    IMP_Application_get_Charts,
    IMP_Application_get_Columns,
    IMP_Application_get_CommandBars,
    IMP_Application_get_DDEAppReturnCode,
    IMP_Application_DDEExecute,
    IMP_Application_DDEInitiate,
    IMP_Application_DDEPoke,
    IMP_Application_DDERequest,
    IMP_Application_DDETerminate,
    IMP_Application_get_DialogSheets,
    IMP_Application_Evaluate,
    IMP_Application__Evaluate,
    IMP_Application_ExecuteExcel4Macro,
    IMP_Application_Intersect,
    IMP_Application_get_MenuBars,
    IMP_Application_get_Modules,
    IMP_Application_get_Names,
    IMP_Application_get_Range,
    IMP_Application_get_Rows,
    IMP_Application_Run,
    IMP_Application__Run2,
    IMP_Application_get_Selection,
    IMP_Application_SendKeys,
    IMP_Application_get_Sheets,
    IMP_Application_get_ShortcutMenus,
    IMP_Application_get_ThisWorkbook,
    IMP_Application_get_Toolbars,
    IMP_Application_Union,
    IMP_Application_get_Windows,
    IMP_Application_get_Workbooks,
    IMP_Application_get_WorksheetFunction,
    IMP_Application_get_Worksheets,
    IMP_Application_get_Excel4IntlMacroSheets,
    IMP_Application_get_Excel4MacroSheets,
    IMP_Application_ActivateMicrosoftApp,
    IMP_Application_AddChartAutoFormat,
    IMP_Application_AddCustomList,
    IMP_Application_get_AlertBeforeOverwriting,
    IMP_Application_put_AlertBeforeOverwriting,
    IMP_Application_get_AltStartupPath,
    IMP_Application_put_AltStartupPath,
    IMP_Application_get_AskToUpdateLinks,
    IMP_Application_put_AskToUpdateLinks,
    IMP_Application_get_EnableAnimations,
    IMP_Application_put_EnableAnimations,
    IMP_Application_get_AutoCorrect,
    IMP_Application_get_Build,
    IMP_Application_get_CalculateBeforeSave,
    IMP_Application_put_CalculateBeforeSave,
    IMP_Application_get_Calculation,
    IMP_Application_put_Calculation,
    IMP_Application_get_Caller,
    IMP_Application_get_CanPlaySounds,
    IMP_Application_get_CanRecordSounds,
    IMP_Application_get_Caption,
    IMP_Application_put_Caption,
    IMP_Application_get_CellDragAndDrop,
    IMP_Application_put_CellDragAndDrop,
    IMP_Application_CentimetersToPoints,
    IMP_Application_CheckSpelling,
    IMP_Application_get_ClipboardFormats,
    IMP_Application_get_DisplayClipboardWindow,
    IMP_Application_put_DisplayClipboardWindow,
    IMP_Application_get_ColorButtons,
    IMP_Application_put_ColorButtons,
    IMP_Application_get_CommandUnderlines,
    IMP_Application_put_CommandUnderlines,
    IMP_Application_get_ConstrainNumeric,
    IMP_Application_put_ConstrainNumeric,
    IMP_Application_ConvertFormula,
    IMP_Application_get_CopyObjectsWithCells,
    IMP_Application_put_CopyObjectsWithCells,
    IMP_Application_get_Cursor,
    IMP_Application_put_Cursor,
    IMP_Application_get_CustomListCount,
    IMP_Application_get_CutCopyMode,
    IMP_Application_put_CutCopyMode,
    IMP_Application_get_DataEntryMode,
    IMP_Application_put_DataEntryMode,
    IMP_Application_Dummy1,
    IMP_Application_Dummy2,
    IMP_Application_Dummy3,
    IMP_Application_Dummy4,
    IMP_Application_Dummy5,
    IMP_Application_Dummy6,
    IMP_Application_Dummy7,
    IMP_Application_Dummy8,
    IMP_Application_Dummy9,
    IMP_Application_Dummy10,
    IMP_Application_Dummy11,
    IMP_Application_get__Default,
    IMP_Application_get_DefaultFilePath,
    IMP_Application_put_DefaultFilePath,
    IMP_Application_DeleteChartAutoFormat,
    IMP_Application_DeleteCustomList,
    IMP_Application_get_Dialogs,
    IMP_Application_put_DisplayAlerts,
    IMP_Application_get_DisplayAlerts,
    IMP_Application_get_DisplayFormulaBar,
    IMP_Application_put_DisplayFormulaBar,
    IMP_Application_get_DisplayFullScreen,
    IMP_Application_put_DisplayFullScreen,
    IMP_Application_get_DisplayNoteIndicator,
    IMP_Application_put_DisplayNoteIndicator,
    IMP_Application_get_DisplayCommentIndicator,
    IMP_Application_put_DisplayCommentIndicator,
    IMP_Application_get_DisplayExcel4Menus,
    IMP_Application_put_DisplayExcel4Menus,
    IMP_Application_get_DisplayRecentFiles,
    IMP_Application_put_DisplayRecentFiles,
    IMP_Application_get_DisplayScrollBars,
    IMP_Application_put_DisplayScrollBars,
    IMP_Application_get_DisplayStatusBar,
    IMP_Application_put_DisplayStatusBar,
    IMP_Application_DoubleClick,
    IMP_Application_get_EditDirectlyInCell,
    IMP_Application_put_EditDirectlyInCell,
    IMP_Application_get_EnableAutoComplete,
    IMP_Application_put_EnableAutoComplete,
    IMP_Application_get_EnableCancelKey,
    IMP_Application_put_EnableCancelKey,
    IMP_Application_get_EnableSound,
    IMP_Application_put_EnableSound,
    IMP_Application_get_EnableTipWizard,
    IMP_Application_put_EnableTipWizard,
    IMP_Application_get_FileConverters,
    IMP_Application_get_FileSearch,
    IMP_Application_get_FileFind,
    IMP_Application__FindFile,
    IMP_Application_get_FixedDecimal,
    IMP_Application_put_FixedDecimal,
    IMP_Application_get_FixedDecimalPlaces,
    IMP_Application_put_FixedDecimalPlaces,
    IMP_Application_GetCustomListContents,
    IMP_Application_GetCustomListNum,
    IMP_Application_GetOpenFilename,
    IMP_Application_GetSaveAsFilename,
    IMP_Application_Goto,
    IMP_Application_get_Height,
    IMP_Application_put_Height,
    IMP_Application_Help,
    IMP_Application_get_IgnoreRemoteRequests,
    IMP_Application_put_IgnoreRemoteRequests,
    IMP_Application_InchesToPoints,
    IMP_Application_InputBox,
    IMP_Application_get_Interactive,
    IMP_Application_put_Interactive,
    IMP_Application_get_International,
    IMP_Application_get_Iteration,
    IMP_Application_put_Iteration,
    IMP_Application_get_LargeButtons,
    IMP_Application_put_LargeButtons,
    IMP_Application_get_Left,
    IMP_Application_put_Left,
    IMP_Application_get_LibraryPath,
    IMP_Application_MacroOptions,
    IMP_Application_MailLogoff,
    IMP_Application_MailLogon,
    IMP_Application_get_MailSession,
    IMP_Application_get_MailSystem,
    IMP_Application_get_MathCoprocessorAvailable,
    IMP_Application_get_MaxChange,
    IMP_Application_put_MaxChange,
    IMP_Application_get_MaxIterations,
    IMP_Application_put_MaxIterations,
    IMP_Application_get_MemoryFree,
    IMP_Application_get_MemoryTotal,
    IMP_Application_get_MemoryUsed,
    IMP_Application_get_MouseAvailable,
    IMP_Application_get_MoveAfterReturn,
    IMP_Application_put_MoveAfterReturn,
    IMP_Application_get_MoveAfterReturnDirection,
    IMP_Application_put_MoveAfterReturnDirection,
    IMP_Application_get_RecentFiles,
    IMP_Application_get_Name,
    IMP_Application_NextLetter,
    IMP_Application_get_NetworkTemplatesPath,
    IMP_Application_get_ODBCErrors,
    IMP_Application_get_ODBCTimeout,
    IMP_Application_put_ODBCTimeout,
    IMP_Application_get_OnCalculate,
    IMP_Application_put_OnCalculate,
    IMP_Application_get_OnData,
    IMP_Application_put_OnData,
    IMP_Application_get_OnDoubleClick,
    IMP_Application_put_OnDoubleClick,
    IMP_Application_get_OnEntry,
    IMP_Application_put_OnEntry,
    IMP_Application_OnKey,
    IMP_Application_OnRepeat,
    IMP_Application_get_OnSheetActivate,
    IMP_Application_put_OnSheetActivate,
    IMP_Application_get_OnSheetDeactivate,
    IMP_Application_put_OnSheetDeactivate,
    IMP_Application_OnTime,
    IMP_Application_OnUndo,
    IMP_Application_get_OnWindow,
    IMP_Application_put_OnWindow,
    IMP_Application_get_OperatingSystem,
    IMP_Application_get_OrganizationName,
    IMP_Application_get_Path,
    IMP_Application_get_PathSeparator,
    IMP_Application_get_PreviousSelections,
    IMP_Application_get_PivotTableSelection,
    IMP_Application_put_PivotTableSelection,
    IMP_Application_get_PromptForSummaryInfo,
    IMP_Application_put_PromptForSummaryInfo,
    IMP_Application_Quit,
    IMP_Application_RecordMacro,
    IMP_Application_get_RecordRelative,
    IMP_Application_get_ReferenceStyle,
    IMP_Application_put_ReferenceStyle,
    IMP_Application_get_RegisteredFunctions,
    IMP_Application_RegisterXLL,
    IMP_Application_Repeat,
    IMP_Application_ResetTipWizard,
    IMP_Application_get_RollZoom,
    IMP_Application_put_RollZoom,
    IMP_Application_Save,
    IMP_Application_SaveWorkspace,
    IMP_Application_get_ScreenUpdating,
    IMP_Application_put_ScreenUpdating,
    IMP_Application_SetDefaultChart,
    IMP_Application_get_SheetsInNewWorkbook,
    IMP_Application_put_SheetsInNewWorkbook,
    IMP_Application_get_ShowChartTipNames,
    IMP_Application_put_ShowChartTipNames,
    IMP_Application_get_ShowChartTipValues,
    IMP_Application_put_ShowChartTipValues,
    IMP_Application_get_StandardFont,
    IMP_Application_put_StandardFont,
    IMP_Application_get_StandardFontSize,
    IMP_Application_put_StandardFontSize,
    IMP_Application_get_StartupPath,
    IMP_Application_get_StatusBar,
    IMP_Application_put_StatusBar,
    IMP_Application_get_TemplatesPath,
    IMP_Application_get_ShowToolTips,
    IMP_Application_put_ShowToolTips,
    IMP_Application_get_Top,
    IMP_Application_put_Top,
    IMP_Application_get_DefaultSaveFormat,
    IMP_Application_put_DefaultSaveFormat,
    IMP_Application_get_TransitionMenuKey,
    IMP_Application_put_TransitionMenuKey,
    IMP_Application_get_TransitionMenuKeyAction,
    IMP_Application_put_TransitionMenuKeyAction,
    IMP_Application_get_TransitionNavigKeys,
    IMP_Application_put_TransitionNavigKeys,
    IMP_Application_Undo,
    IMP_Application_get_UsableHeight,
    IMP_Application_get_UsableWidth,
    IMP_Application_put_UserControl,
    IMP_Application_get_UserControl,
    IMP_Application_get_UserName,
    IMP_Application_put_UserName,
    IMP_Application_get_Value,
    IMP_Application_get_VBE,
    IMP_Application_get_Version,
    IMP_Application_get_Visible,
    IMP_Application_put_Visible,
    IMP_Application_Volatile,
    IMP_Application__Wait,
    IMP_Application_get_Width,
    IMP_Application_put_Width,
    IMP_Application_get_WindowsForPens,
    IMP_Application_put_WindowState,
    IMP_Application_get_WindowState,
    IMP_Application_get_UILanguage,
    IMP_Application_put_UILanguage,
    IMP_Application_get_DefaultSheetDirection,
    IMP_Application_put_DefaultSheetDirection,
    IMP_Application_get_CursorMovement,
    IMP_Application_put_CursorMovement,
    IMP_Application_get_ControlCharacters,
    IMP_Application_put_ControlCharacters,
    IMP_Application__WSFunction,
    IMP_Application_get_EnableEvents,
    IMP_Application_put_EnableEvents,
    IMP_Application_get_DisplayInfoWindow,
    IMP_Application_put_DisplayInfoWindow,
    IMP_Application_Wait,
    IMP_Application_get_ExtendList,
    IMP_Application_put_ExtendList,
    IMP_Application_get_OLEDBErrors,
    IMP_Application_GetPhonetic,
    IMP_Application_get_COMAddIns,
    IMP_Application_get_DefaultWebOptions,
    IMP_Application_get_ProductCode,
    IMP_Application_get_UserLibraryPath,
    IMP_Application_get_AutoPercentEntry,
    IMP_Application_put_AutoPercentEntry,
    IMP_Application_get_LanguageSettings,
    IMP_Application_get_Dummy101,
    IMP_Application_Dummy12,
    IMP_Application_get_AnswerWizard,
    IMP_Application_CalculateFull,
    IMP_Application_FindFile,
    IMP_Application_get_CalculationVersion,
    IMP_Application_get_ShowWindowsInTaskbar,
    IMP_Application_put_ShowWindowsInTaskbar,
    IMP_Application_get_FeatureInstall,
    IMP_Application_put_FeatureInstall,
    IMP_Application_get_Ready,
    IMP_Application_Dummy13,
    IMP_Application_get_FindFormat,
    IMP_Application_putref_FindFormat,
    IMP_Application_get_ReplaceFormat,
    IMP_Application_putref_ReplaceFormat,
    IMP_Application_get_UsedObjects,
    IMP_Application_get_CalculationState,
    IMP_Application_get_CalculationInterruptKey,
    IMP_Application_put_CalculationInterruptKey,
    IMP_Application_get_Watches,
    IMP_Application_get_DisplayFunctionToolTips,
    IMP_Application_put_DisplayFunctionToolTips,
    IMP_Application_get_AutomationSecurity,
    IMP_Application_put_AutomationSecurity,
    IMP_Application_get_FileDialog,
    IMP_Application_Dummy14,
    IMP_Application_CalculateFullRebuild,
    IMP_Application_get_DisplayPasteOptions,
    IMP_Application_put_DisplayPasteOptions,
    IMP_Application_get_DisplayInsertOptions,
    IMP_Application_put_DisplayInsertOptions,
    IMP_Application_get_GenerateGetPivotData,
    IMP_Application_put_GenerateGetPivotData,
    IMP_Application_get_AutoRecover,
    IMP_Application_get_Hwnd,
    IMP_Application_get_Hinstance,
    IMP_Application_CheckAbort,
    IMP_Application_get_ErrorCheckingOptions,
    IMP_Application_get_AutoFormatAsYouTypeReplaceHyperlinks,
    IMP_Application_put_AutoFormatAsYouTypeReplaceHyperlinks,
    IMP_Application_get_SmartTagRecognizers,
    IMP_Application_get_NewWorkbook,
    IMP_Application_get_SpellingOptions,
    IMP_Application_get_Speech,
    IMP_Application_get_MapPaperSize,
    IMP_Application_put_MapPaperSize,
    IMP_Application_get_ShowStartupDialog,
    IMP_Application_put_ShowStartupDialog,
    IMP_Application_get_DecimalSeparator,
    IMP_Application_put_DecimalSeparator,
    IMP_Application_get_ThousandsSeparator,
    IMP_Application_put_ThousandsSeparator,
    IMP_Application_get_UseSystemSeparators,
    IMP_Application_put_UseSystemSeparators,
    IMP_Application_get_ThisCell,
    IMP_Application_get_RTD,
    IMP_Application_get_DisplayDocumentActionTaskPane,
    IMP_Application_put_DisplayDocumentActionTaskPane,
    IMP_Application_DisplayXMLSourcePane,
    IMP_Application_get_ArbitraryXMLSupportAvailable,
    IMP_Application_Support,
};

/*
**  Dispinterface
*/

#define DISPAPPEXCEL_THIS(iface) DEFINE_THIS(_ApplicationImpl, DispApplication, iface);

/*** IUnknown methods ***/
static HRESULT WINAPI DISP_IMP_Application_QueryInterface(
        Disp_Application* iface,
        REFIID riid,
        void **ppvObject)
{
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    
    return IMP_Application_QueryInterface( APPEXCEL(This), riid, ppvObject );    
}

static ULONG WINAPI DISP_IMP_Application_AddRef(
        Disp_Application* iface)
{
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    
    return IMP_Application_AddRef( APPEXCEL(This) );            
}

static ULONG WINAPI DISP_IMP_Application_Release(
        Disp_Application* iface)
{
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    
    return IMP_Application_Release( APPEXCEL(This) );              
}

/*** IDispatch methods ***/
static HRESULT WINAPI DISP_IMP_Application_GetTypeInfoCount(
        Disp_Application* iface,
        UINT *pctinfo)
{
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    
    return IMP_Application_GetTypeInfoCount( APPEXCEL(This), pctinfo );                 
}

static HRESULT WINAPI DISP_IMP_Application_GetTypeInfo(
        Disp_Application* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    
    return IMP_Application_GetTypeInfo( APPEXCEL(This), iTInfo, lcid, ppTInfo );                 
}

static HRESULT WINAPI DISP_IMP_Application_GetIDsOfNames(
        Disp_Application* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    
    return IMP_Application_GetIDsOfNames( APPEXCEL(This), riid, rgszNames, cNames, lcid, rgDispId );                 
}

static HRESULT WINAPI DISP_IMP_Application_Invoke(
        Disp_Application* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    
    return IMP_Application_Invoke( APPEXCEL(This), dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);                 
}

/*** _Application methods ***/
static IDispatch* WINAPI DISP_IMP_Application_get_Application(
        Disp_Application* iface)
{
             
}

static XlCreator WINAPI DISP_IMP_Application_get_Creator(
        Disp_Application* iface) 
{    
    TRACE_IN;  
    _ApplicationImpl *This = DISPAPPEXCEL_THIS( iface );
    HRESULT hres;
    XlCreator tmp_creator = xlCreatorCode;
    
    hres = IMP_Application_get_Creator( APPEXCEL(This), &tmp_creator );
    if (FAILED(hres))
    {
        ERR("IMP_Application_get_Creator failed");             
    }
    
    TRACE_OUT;
    return tmp_creator;
}

static IDispatch* WINAPI DISP_IMP_Application_get_Parent(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_ActiveCell(
        Disp_Application* iface,
        IDispatch **RHS) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_ActiveChart(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_ActiveDialog(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_ActiveMenuBar(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get_ActivePrinter(
        Disp_Application* iface,
        LCID lcid) {         }

static BSTR WINAPI DISP_IMP_Application_put_ActivePrinter(
        Disp_Application* iface,
        LCID lcid) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_ActiveSheet(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_ActiveWindow(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_ActiveWorkbook(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_AddIns(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_Assistant(
        Disp_Application* iface,
        IDispatch **RHS) {         }

static void WINAPI DISP_IMP_Application_Calculate(
        Disp_Application* iface,
        LCID lcid) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_Cells(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_Charts(
        Disp_Application* iface) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_Columns(
        Disp_Application* iface,
        VARIANT param) {         }

static IDispatch* WINAPI DISP_IMP_Application_get_CommandBars(
        Disp_Application* iface) {         }

static long WINAPI DISP_IMP_Application_get_DDEAppReturnCode(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_DDEExecute(
        Disp_Application* iface,
        long Channel,
        BSTR String,
        LCID lcid) {         }

static long WINAPI DISP_IMP_Application_DDEInitiate(
        Disp_Application* iface,
        BSTR App,
        BSTR Topic,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_DDEPoke(
        Disp_Application* iface,
        long Channel,
        VARIANT Item,
        VARIANT Data,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_DDERequest(
        Disp_Application* iface,
        long Channel,
        BSTR Item,
        LCID lcid) {         }

static HRESULT WINAPI DISP_IMP_Application_DDETerminate(
        Disp_Application* iface,
        long Channel,
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_DialogSheets(
        Disp_Application* iface) {         }

static VARIANT WINAPI DISP_IMP_Application_Evaluate(
        Disp_Application* iface,
        VARIANT Name,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application__Evaluate(
        Disp_Application* iface,
        VARIANT Name,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_ExecuteExcel4Macro(
        Disp_Application* iface,
        BSTR String,
        LCID lcid) {         }

static IDispatch* WINAPI DISP_IMP_Application_Intersect(
        Disp_Application* iface,
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
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_MenuBars(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Modules(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Names(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Range(
        Disp_Application* iface,
        VARIANT Cell1,
        VARIANT Cell2) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Rows(
        Disp_Application* iface,
        VARIANT param) {         }

static VARIANT WINAPI DISP_IMP_Application_Run(
        Disp_Application* iface,
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
        VARIANT Arg30) {         }

static VARIANT WINAPI DISP_IMP_Application__Run2(
        Disp_Application* iface,
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
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Selection(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_SendKeys(
        Disp_Application* iface,
        VARIANT Keys,
        VARIANT Wait,
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Sheets(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_ShortcutMenus(
        Disp_Application* iface,
        long Index) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_ThisWorkbook(
        Disp_Application* iface,
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Toolbars(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_Union(
        Disp_Application* iface,
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
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Windows(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Workbooks(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_WorksheetFunction(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Worksheets(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Excel4IntlMacroSheets(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Excel4MacroSheets(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_ActivateMicrosoftApp(
        Disp_Application* iface,
        XlMSApplication Index,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_AddChartAutoFormat(
        Disp_Application* iface,
        VARIANT Chart,
        BSTR Name,
        VARIANT Description,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_AddCustomList(
        Disp_Application* iface,
        VARIANT ListArray,
        VARIANT ByRow,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_AlertBeforeOverwriting(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_AlertBeforeOverwriting(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_AltStartupPath(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_AltStartupPath(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_AskToUpdateLinks(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_AskToUpdateLinks(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_EnableAnimations(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_EnableAnimations(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_AutoCorrect(
        Disp_Application* iface) {         }

static long WINAPI DISP_IMP_Application_get_Build(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_CalculateBeforeSave(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_CalculateBeforeSave(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static XlCalculation WINAPI DISP_IMP_Application_get_Calculation(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Calculation(
        Disp_Application* iface,
        LCID lcid,
        XlCalculation RHS) {         }

static VARIANT WINAPI DISP_IMP_Application_get_Caller(
        Disp_Application* iface,
        VARIANT Index,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_CanPlaySounds(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_CanRecordSounds(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_get_Caption(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_Caption(
        Disp_Application* iface,
        VARIANT vName) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_CellDragAndDrop(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_CellDragAndDrop(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static double WINAPI DISP_IMP_Application_CentimetersToPoints(
        Disp_Application* iface,
        double Centimeters,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_CheckSpelling(
        Disp_Application* iface,
        BSTR Word,
        VARIANT CustomDictionary,
        VARIANT IgnoreUppercase,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_get_ClipboardFormats(
        Disp_Application* iface,
        VARIANT Index,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayClipboardWindow(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DisplayClipboardWindow(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ColorButtons(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ColorButtons(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static XlCommandUnderlines WINAPI DISP_IMP_Application_get_CommandUnderlines(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_CommandUnderlines(
        Disp_Application* iface,
        LCID lcid,
        XlCommandUnderlines RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ConstrainNumeric(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_ConstrainNumeric(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT WINAPI DISP_IMP_Application_ConvertFormula(
        Disp_Application* iface,
        VARIANT Formula,
        XlReferenceStyle FromReferenceStyle,
        VARIANT ToReferenceStyle,
        VARIANT ToAbsolute,
        VARIANT RelativeTo,
        long Lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_CopyObjectsWithCells(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_CopyObjectsWithCells(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static XlMousePointer WINAPI DISP_IMP_Application_get_Cursor(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Cursor(
        Disp_Application* iface,
        LCID lcid,
        XlMousePointer RHS) {         }

static long WINAPI DISP_IMP_Application_get_CustomListCount(
        Disp_Application* iface,
        LCID lcid) {         }

static XlCutCopyMode WINAPI DISP_IMP_Application_get_CutCopyMode(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_CutCopyMode(
        Disp_Application* iface,
        LCID lcid,
        XlCutCopyMode RHS) {         }

static long WINAPI DISP_IMP_Application_get_DataEntryMode(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DataEntryMode(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy1(
        Disp_Application* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy2(
        Disp_Application* iface,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy3(
        Disp_Application* iface) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy4(
        Disp_Application* iface,
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
        VARIANT Arg15) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy5(
        Disp_Application* iface,
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
        VARIANT Arg13) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy6(
        Disp_Application* iface) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy7(
        Disp_Application* iface) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy8(
        Disp_Application* iface,
        VARIANT Arg1) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy9(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_Dummy10(
        Disp_Application* iface,
        VARIANT arg) {         }

static void WINAPI DISP_IMP_Application_Dummy11(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get__Default(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get_DefaultFilePath(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DefaultFilePath(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static void WINAPI DISP_IMP_Application_DeleteChartAutoFormat(
        Disp_Application* iface,
        BSTR Name,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_DeleteCustomList(
        Disp_Application* iface,
        long ListNum,
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Dialogs(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayAlerts(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL vbDisplayAlerts) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayAlerts(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayFormulaBar(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DisplayFormulaBar(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayFullScreen(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DisplayFullScreen(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayNoteIndicator(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayNoteIndicator(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static XlCommentDisplayMode WINAPI DISP_IMP_Application_get_DisplayCommentIndicator(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayCommentIndicator(
        Disp_Application* iface,
        XlCommentDisplayMode RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayExcel4Menus(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DisplayExcel4Menus(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayRecentFiles(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayRecentFiles(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayScrollBars(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DisplayScrollBars(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayStatusBar(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DisplayStatusBar(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static void WINAPI DISP_IMP_Application_DoubleClick(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_EditDirectlyInCell(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_EditDirectlyInCell(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_EnableAutoComplete(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_EnableAutoComplete(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static XlEnableCancelKey WINAPI DISP_IMP_Application_get_EnableCancelKey(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_EnableCancelKey(
        Disp_Application* iface,
        LCID lcid,
        XlEnableCancelKey RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_EnableSound(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_EnableSound(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_EnableTipWizard(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_EnableTipWizard(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT WINAPI DISP_IMP_Application_get_FileConverters(
        Disp_Application* iface,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_FileSearch(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_FileFind(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application__FindFile(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_FixedDecimal(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_FixedDecimal(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static long WINAPI DISP_IMP_Application_get_FixedDecimalPlaces(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_FixedDecimalPlaces(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static VARIANT WINAPI DISP_IMP_Application_GetCustomListContents(
        Disp_Application* iface,
        long ListNum,
        LCID lcid) {         }

static long WINAPI DISP_IMP_Application_GetCustomListNum(
        Disp_Application* iface,
        VARIANT ListArray,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_GetOpenFilename(
        Disp_Application* iface,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        VARIANT MultiSelect,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_GetSaveAsFilename(
        Disp_Application* iface,
        VARIANT InitialFilename,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_Goto(
        Disp_Application* iface,
        VARIANT Reference,
        VARIANT Scroll,
        LCID lcid) {         }

static double WINAPI DISP_IMP_Application_get_Height(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Height(
        Disp_Application* iface,
        LCID lcid,
        double RHS) {         }

static void WINAPI DISP_IMP_Application_Help(
        Disp_Application* iface,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_IgnoreRemoteRequests(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_IgnoreRemoteRequests(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static double WINAPI DISP_IMP_Application_InchesToPoints(
        Disp_Application* iface,
        double Inches,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_InputBox(
        Disp_Application* iface,
        BSTR Prompt,
        VARIANT Title,
        VARIANT Default,
        VARIANT Left,
        VARIANT Top,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        VARIANT Type,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_Interactive(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Interactive(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT WINAPI DISP_IMP_Application_get_International(
        Disp_Application* iface,
        VARIANT Index,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_Iteration(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Iteration(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_LargeButtons(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_LargeButtons(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static double WINAPI DISP_IMP_Application_get_Left(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Left(
        Disp_Application* iface,
        LCID lcid,
        double RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_LibraryPath(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_MacroOptions(
        Disp_Application* iface,
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
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_MailLogoff(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_MailLogon(
        Disp_Application* iface,
        VARIANT Name,
        VARIANT Password,
        VARIANT DownloadNewMail,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_get_MailSession(
        Disp_Application* iface,
        LCID lcid) {         }

static XlMailSystem WINAPI DISP_IMP_Application_get_MailSystem(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_MathCoprocessorAvailable(
        Disp_Application* iface,
        LCID lcid) {         }

static double WINAPI DISP_IMP_Application_get_MaxChange(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_MaxChange(
        Disp_Application* iface,
        LCID lcid,
        double RHS) {         }

static long WINAPI DISP_IMP_Application_get_MaxIterations(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_MaxIterations(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static long WINAPI DISP_IMP_Application_get_MemoryFree(
        Disp_Application* iface,
        LCID lcid) {         }

static long WINAPI DISP_IMP_Application_get_MemoryTotal(
        Disp_Application* iface,
        LCID lcid) {         }

static long WINAPI DISP_IMP_Application_get_MemoryUsed(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_MouseAvailable(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_MoveAfterReturn(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_MoveAfterReturn(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static XlDirection WINAPI DISP_IMP_Application_get_MoveAfterReturnDirection(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_MoveAfterReturnDirection(
        Disp_Application* iface,
        LCID lcid,
        XlDirection RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_RecentFiles(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get_Name(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_NextLetter(
        Disp_Application* iface,
        LCID lcid) {         }

static BSTR WINAPI DISP_IMP_Application_get_NetworkTemplatesPath(
        Disp_Application* iface,
        LCID lcid) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_ODBCErrors(
        Disp_Application* iface) {         }

static long WINAPI DISP_IMP_Application_get_ODBCTimeout(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ODBCTimeout(
        Disp_Application* iface,
        long RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_OnCalculate(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_OnCalculate(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_OnData(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_OnData(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_OnDoubleClick(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_OnDoubleClick(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_OnEntry(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_OnEntry(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static void WINAPI DISP_IMP_Application_OnKey(
        Disp_Application* iface,
        BSTR Key,
        VARIANT Procedure,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_OnRepeat(
        Disp_Application* iface,
        BSTR Text,
        BSTR Procedure,
        LCID lcid) {         }

static BSTR WINAPI DISP_IMP_Application_get_OnSheetActivate(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_OnSheetActivate(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_OnSheetDeactivate(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_OnSheetDeactivate(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static void WINAPI DISP_IMP_Application_OnTime(
        Disp_Application* iface,
        VARIANT EarliestTime,
        BSTR Procedure,
        VARIANT LatestTime,
        VARIANT Schedule,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_OnUndo(
        Disp_Application* iface,
        BSTR Text,
        BSTR Procedure,
        LCID lcid) {         }

static BSTR WINAPI DISP_IMP_Application_get_OnWindow(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_OnWindow(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_OperatingSystem(
        Disp_Application* iface,
        LCID lcid) {         }

static BSTR WINAPI DISP_IMP_Application_get_OrganizationName(
        Disp_Application* iface,
        LCID lcid) {         }

static BSTR WINAPI DISP_IMP_Application_get_Path(
        Disp_Application* iface,
        LCID lcid) {         }

static BSTR WINAPI DISP_IMP_Application_get_PathSeparator(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_get_PreviousSelections(
        Disp_Application* iface,
        VARIANT Index,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_PivotTableSelection(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_PivotTableSelection(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_PromptForSummaryInfo(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_PromptForSummaryInfo(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static void WINAPI DISP_IMP_Application_Quit(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_RecordMacro(
        Disp_Application* iface,
        VARIANT BasicCode,
        VARIANT XlmCode,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_RecordRelative(
        Disp_Application* iface,
        LCID lcid) {         }

static XlReferenceStyle WINAPI DISP_IMP_Application_get_ReferenceStyle(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_ReferenceStyle(
        Disp_Application* iface,
        LCID lcid,
        XlReferenceStyle RHS) {         }

static VARIANT WINAPI DISP_IMP_Application_get_RegisteredFunctions(
        Disp_Application* iface,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_RegisterXLL(
        Disp_Application* iface,
        BSTR Filename,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_Repeat(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_ResetTipWizard(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_RollZoom(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_RollZoom(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static void WINAPI DISP_IMP_Application_Save(
        Disp_Application* iface,
        VARIANT Filename,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_SaveWorkspace(
        Disp_Application* iface,
        VARIANT Filename,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ScreenUpdating(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_ScreenUpdating(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static void WINAPI DISP_IMP_Application_SetDefaultChart(
        Disp_Application* iface,
        VARIANT FormatName,
        VARIANT Gallery) {         }

static long WINAPI DISP_IMP_Application_get_SheetsInNewWorkbook(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_SheetsInNewWorkbook(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ShowChartTipNames(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ShowChartTipNames(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ShowChartTipValues(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ShowChartTipValues(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_StandardFont(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_StandardFont(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static double WINAPI DISP_IMP_Application_get_StandardFontSize(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_StandardFontSize(
        Disp_Application* iface,
        LCID lcid,
        double RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_StartupPath(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT WINAPI DISP_IMP_Application_get_StatusBar(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_StatusBar(
        Disp_Application* iface,
        LCID lcid,
        VARIANT RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_TemplatesPath(
        Disp_Application* iface,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ShowToolTips(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ShowToolTips(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static double WINAPI DISP_IMP_Application_get_Top(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Top(
        Disp_Application* iface,
        LCID lcid,
        double RHS) {         }

static XlFileFormat WINAPI DISP_IMP_Application_get_DefaultSaveFormat(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DefaultSaveFormat(
        Disp_Application* iface,
        XlFileFormat RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_TransitionMenuKey(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_TransitionMenuKey(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static long WINAPI DISP_IMP_Application_get_TransitionMenuKeyAction(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_TransitionMenuKeyAction(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_TransitionNavigKeys(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_TransitionNavigKeys(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static void WINAPI DISP_IMP_Application_Undo(
        Disp_Application* iface,
        LCID lcid) {         }

static double WINAPI DISP_IMP_Application_get_UsableHeight(
        Disp_Application* iface,
        LCID lcid) {         }

static double WINAPI DISP_IMP_Application_get_UsableWidth(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_UserControl(
        Disp_Application* iface,
        VARIANT_BOOL vbUserControl) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_UserControl(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get_UserName(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_UserName(
        Disp_Application* iface,
        LCID lcid,
        BSTR RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_Value(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_VBE(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get_Version(
        Disp_Application* iface,
        long Lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_Visible(
        Disp_Application* iface,
        LCID Lcid) {         }

static void WINAPI DISP_IMP_Application_put_Visible(
        Disp_Application* iface,
        LCID Lcid,
        VARIANT_BOOL vbVisible) {         }

static void WINAPI DISP_IMP_Application_Volatile(
        Disp_Application* iface,
        VARIANT Volatile,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application__Wait(
        Disp_Application* iface,
        VARIANT Time,
        LCID lcid) {         }

static double WINAPI DISP_IMP_Application_get_Width(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_Width(
        Disp_Application* iface,
        LCID lcid,
        double RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_WindowsForPens(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_WindowState(
        Disp_Application* iface,
        LCID lcid,
        XlWindowState State) {         }

static XlWindowState WINAPI DISP_IMP_Application_get_WindowState(
        Disp_Application* iface,
        LCID lcid) {         }

static long WINAPI DISP_IMP_Application_get_UILanguage(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_UILanguage(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static long WINAPI DISP_IMP_Application_get_DefaultSheetDirection(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_DefaultSheetDirection(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static long WINAPI DISP_IMP_Application_get_CursorMovement(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_CursorMovement(
        Disp_Application* iface,
        LCID lcid,
        long RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ControlCharacters(
        Disp_Application* iface,
        LCID lcid) {         }

static void WINAPI DISP_IMP_Application_put_ControlCharacters(
        Disp_Application* iface,
        LCID lcid,
        VARIANT_BOOL RHS) {         }

static VARIANT WINAPI DISP_IMP_Application__WSFunction(
        Disp_Application* iface,
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
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_EnableEvents(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_EnableEvents(
        Disp_Application* iface,
        VARIANT_BOOL vbee) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayInfoWindow(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayInfoWindow(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_Wait(
        Disp_Application* iface,
        VARIANT Time,
        LCID lcid) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ExtendList(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ExtendList(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_OLEDBErrors(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_GetPhonetic(
        Disp_Application* iface,
        VARIANT Text) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_COMAddIns(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_DefaultWebOptions(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get_ProductCode(
        Disp_Application* iface) {         }

static BSTR WINAPI DISP_IMP_Application_get_UserLibraryPath(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_AutoPercentEntry(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_AutoPercentEntry(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_LanguageSettings(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Dummy101(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_Dummy12(
        Disp_Application* iface,
        IDispatch *p1,
        IDispatch *p2) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_AnswerWizard(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_CalculateFull(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_FindFile(
        Disp_Application* iface,
        LCID lcid) {         }

static long WINAPI DISP_IMP_Application_get_CalculationVersion(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ShowWindowsInTaskbar(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ShowWindowsInTaskbar(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static MsoFeatureInstall WINAPI DISP_IMP_Application_get_FeatureInstall(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_FeatureInstall(
        Disp_Application* iface,
        MsoFeatureInstall RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_Ready(
        Disp_Application* iface) {         }

static VARIANT WINAPI DISP_IMP_Application_Dummy13(
        Disp_Application* iface,
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
        VARIANT Arg30) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_FindFormat(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_putref_FindFormat(
        Disp_Application* iface,
        IDispatch *RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_ReplaceFormat(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_putref_ReplaceFormat(
        Disp_Application* iface,
        IDispatch *RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_UsedObjects(
        Disp_Application* iface) {         }

static XlCalculationState WINAPI DISP_IMP_Application_get_CalculationState(
        Disp_Application* iface) {         }

static XlCalculationInterruptKey WINAPI DISP_IMP_Application_get_CalculationInterruptKey(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_CalculationInterruptKey(
        Disp_Application* iface,
        XlCalculationInterruptKey RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Watches(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayFunctionToolTips(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayFunctionToolTips(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static MsoAutomationSecurity WINAPI DISP_IMP_Application_get_AutomationSecurity(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_AutomationSecurity(
        Disp_Application* iface,
        MsoAutomationSecurity RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_FileDialog(
        Disp_Application* iface,
        MsoFileDialogType fileDialogType) {         }

static void WINAPI DISP_IMP_Application_Dummy14(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_CalculateFullRebuild(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayPasteOptions(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayPasteOptions(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayInsertOptions(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayInsertOptions(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_GenerateGetPivotData(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_GenerateGetPivotData(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_AutoRecover(
        Disp_Application* iface) {         }

static long WINAPI DISP_IMP_Application_get_Hwnd(
        Disp_Application* iface) {         }

static long WINAPI DISP_IMP_Application_get_Hinstance(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_CheckAbort(
        Disp_Application* iface,
        VARIANT KeepAbort) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_ErrorCheckingOptions(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_AutoFormatAsYouTypeReplaceHyperlinks(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_AutoFormatAsYouTypeReplaceHyperlinks(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_SmartTagRecognizers(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_NewWorkbook(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_SpellingOptions(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_Speech(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_MapPaperSize(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_MapPaperSize(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ShowStartupDialog(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ShowStartupDialog(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_DecimalSeparator(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DecimalSeparator(
        Disp_Application* iface,
        BSTR RHS) {         }

static BSTR WINAPI DISP_IMP_Application_get_ThousandsSeparator(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_ThousandsSeparator(
        Disp_Application* iface,
        BSTR RHS) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_UseSystemSeparators(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_UseSystemSeparators(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_ThisCell(
        Disp_Application* iface) {         }

static IDispatch * WINAPI DISP_IMP_Application_get_RTD(
        Disp_Application* iface) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_DisplayDocumentActionTaskPane(
        Disp_Application* iface) {         }

static void WINAPI DISP_IMP_Application_put_DisplayDocumentActionTaskPane(
        Disp_Application* iface,
        VARIANT_BOOL RHS) {         }

static void WINAPI DISP_IMP_Application_DisplayXMLSourcePane(
        Disp_Application* iface,
        VARIANT XmlMap) {         }

static VARIANT_BOOL WINAPI DISP_IMP_Application_get_ArbitraryXMLSupportAvailable(
        Disp_Application* iface) {         }

static VARIANT WINAPI DISP_IMP_Application_Support(
        Disp_Application* iface,
        IDispatch *Object,
        long ID,
        VARIANT arg) {         }



#undef DISPAPPEXCEL_THIS

const Disp_ApplicationVtbl DISP_IMP_Application_Vtbl =
{
    DISP_IMP_Application_QueryInterface,
    DISP_IMP_Application_AddRef,
    DISP_IMP_Application_Release,
    DISP_IMP_Application_GetTypeInfoCount,
    DISP_IMP_Application_GetTypeInfo,
    DISP_IMP_Application_GetIDsOfNames,
    DISP_IMP_Application_Invoke,
    DISP_IMP_Application_get_Application,
    DISP_IMP_Application_get_Creator,
    DISP_IMP_Application_get_Parent,
    DISP_IMP_Application_get_ActiveCell,
    DISP_IMP_Application_get_ActiveChart,
    DISP_IMP_Application_get_ActiveDialog,
    DISP_IMP_Application_get_ActiveMenuBar,
    DISP_IMP_Application_get_ActivePrinter,
    DISP_IMP_Application_put_ActivePrinter,
    DISP_IMP_Application_get_ActiveSheet,
    DISP_IMP_Application_get_ActiveWindow,
    DISP_IMP_Application_get_ActiveWorkbook,
    DISP_IMP_Application_get_AddIns,
    DISP_IMP_Application_get_Assistant,
    DISP_IMP_Application_Calculate,
    DISP_IMP_Application_get_Cells,
    DISP_IMP_Application_get_Charts,
    DISP_IMP_Application_get_Columns,
    DISP_IMP_Application_get_CommandBars,
    DISP_IMP_Application_get_DDEAppReturnCode,
    DISP_IMP_Application_DDEExecute,
    DISP_IMP_Application_DDEInitiate,
    DISP_IMP_Application_DDEPoke,
    DISP_IMP_Application_DDERequest,
    DISP_IMP_Application_DDETerminate,
    DISP_IMP_Application_get_DialogSheets,
    DISP_IMP_Application_Evaluate,
    DISP_IMP_Application__Evaluate,
    DISP_IMP_Application_ExecuteExcel4Macro,
    DISP_IMP_Application_Intersect,
    DISP_IMP_Application_get_MenuBars,
    DISP_IMP_Application_get_Modules,
    DISP_IMP_Application_get_Names,
    DISP_IMP_Application_get_Range,
    DISP_IMP_Application_get_Rows,
    DISP_IMP_Application_Run,
    DISP_IMP_Application__Run2,
    DISP_IMP_Application_get_Selection,
    DISP_IMP_Application_SendKeys,
    DISP_IMP_Application_get_Sheets,
    DISP_IMP_Application_get_ShortcutMenus,
    DISP_IMP_Application_get_ThisWorkbook,
    DISP_IMP_Application_get_Toolbars,
    DISP_IMP_Application_Union,
    DISP_IMP_Application_get_Windows,
    DISP_IMP_Application_get_Workbooks,
    DISP_IMP_Application_get_WorksheetFunction,
    DISP_IMP_Application_get_Worksheets,
    DISP_IMP_Application_get_Excel4IntlMacroSheets,
    DISP_IMP_Application_get_Excel4MacroSheets,
    DISP_IMP_Application_ActivateMicrosoftApp,
    DISP_IMP_Application_AddChartAutoFormat,
    DISP_IMP_Application_AddCustomList,
    DISP_IMP_Application_get_AlertBeforeOverwriting,
    DISP_IMP_Application_put_AlertBeforeOverwriting,
    DISP_IMP_Application_get_AltStartupPath,
    DISP_IMP_Application_put_AltStartupPath,
    DISP_IMP_Application_get_AskToUpdateLinks,
    DISP_IMP_Application_put_AskToUpdateLinks,
    DISP_IMP_Application_get_EnableAnimations,
    DISP_IMP_Application_put_EnableAnimations,
    DISP_IMP_Application_get_AutoCorrect,
    DISP_IMP_Application_get_Build,
    DISP_IMP_Application_get_CalculateBeforeSave,
    DISP_IMP_Application_put_CalculateBeforeSave,
    DISP_IMP_Application_get_Calculation,
    DISP_IMP_Application_put_Calculation,
    DISP_IMP_Application_get_Caller,
    DISP_IMP_Application_get_CanPlaySounds,
    DISP_IMP_Application_get_CanRecordSounds,
    DISP_IMP_Application_get_Caption,
    DISP_IMP_Application_put_Caption,
    DISP_IMP_Application_get_CellDragAndDrop,
    DISP_IMP_Application_put_CellDragAndDrop,
    DISP_IMP_Application_CentimetersToPoints,
    DISP_IMP_Application_CheckSpelling,
    DISP_IMP_Application_get_ClipboardFormats,
    DISP_IMP_Application_get_DisplayClipboardWindow,
    DISP_IMP_Application_put_DisplayClipboardWindow,
    DISP_IMP_Application_get_ColorButtons,
    DISP_IMP_Application_put_ColorButtons,
    DISP_IMP_Application_get_CommandUnderlines,
    DISP_IMP_Application_put_CommandUnderlines,
    DISP_IMP_Application_get_ConstrainNumeric,
    DISP_IMP_Application_put_ConstrainNumeric,
    DISP_IMP_Application_ConvertFormula,
    DISP_IMP_Application_get_CopyObjectsWithCells,
    DISP_IMP_Application_put_CopyObjectsWithCells,
    DISP_IMP_Application_get_Cursor,
    DISP_IMP_Application_put_Cursor,
    DISP_IMP_Application_get_CustomListCount,
    DISP_IMP_Application_get_CutCopyMode,
    DISP_IMP_Application_put_CutCopyMode,
    DISP_IMP_Application_get_DataEntryMode,
    DISP_IMP_Application_put_DataEntryMode,
    DISP_IMP_Application_Dummy1,
    DISP_IMP_Application_Dummy2,
    DISP_IMP_Application_Dummy3,
    DISP_IMP_Application_Dummy4,
    DISP_IMP_Application_Dummy5,
    DISP_IMP_Application_Dummy6,
    DISP_IMP_Application_Dummy7,
    DISP_IMP_Application_Dummy8,
    DISP_IMP_Application_Dummy9,
    DISP_IMP_Application_Dummy10,
    DISP_IMP_Application_Dummy11,
    DISP_IMP_Application_get__Default,
    DISP_IMP_Application_get_DefaultFilePath,
    DISP_IMP_Application_put_DefaultFilePath,
    DISP_IMP_Application_DeleteChartAutoFormat,
    DISP_IMP_Application_DeleteCustomList,
    DISP_IMP_Application_get_Dialogs,
    DISP_IMP_Application_put_DisplayAlerts,
    DISP_IMP_Application_get_DisplayAlerts,
    DISP_IMP_Application_get_DisplayFormulaBar,
    DISP_IMP_Application_put_DisplayFormulaBar,
    DISP_IMP_Application_get_DisplayFullScreen,
    DISP_IMP_Application_put_DisplayFullScreen,
    DISP_IMP_Application_get_DisplayNoteIndicator,
    DISP_IMP_Application_put_DisplayNoteIndicator,
    DISP_IMP_Application_get_DisplayCommentIndicator,
    DISP_IMP_Application_put_DisplayCommentIndicator,
    DISP_IMP_Application_get_DisplayExcel4Menus,
    DISP_IMP_Application_put_DisplayExcel4Menus,
    DISP_IMP_Application_get_DisplayRecentFiles,
    DISP_IMP_Application_put_DisplayRecentFiles,
    DISP_IMP_Application_get_DisplayScrollBars,
    DISP_IMP_Application_put_DisplayScrollBars,
    DISP_IMP_Application_get_DisplayStatusBar,
    DISP_IMP_Application_put_DisplayStatusBar,
    DISP_IMP_Application_DoubleClick,
    DISP_IMP_Application_get_EditDirectlyInCell,
    DISP_IMP_Application_put_EditDirectlyInCell,
    DISP_IMP_Application_get_EnableAutoComplete,
    DISP_IMP_Application_put_EnableAutoComplete,
    DISP_IMP_Application_get_EnableCancelKey,
    DISP_IMP_Application_put_EnableCancelKey,
    DISP_IMP_Application_get_EnableSound,
    DISP_IMP_Application_put_EnableSound,
    DISP_IMP_Application_get_EnableTipWizard,
    DISP_IMP_Application_put_EnableTipWizard,
    DISP_IMP_Application_get_FileConverters,
    DISP_IMP_Application_get_FileSearch,
    DISP_IMP_Application_get_FileFind,
    DISP_IMP_Application__FindFile,
    DISP_IMP_Application_get_FixedDecimal,
    DISP_IMP_Application_put_FixedDecimal,
    DISP_IMP_Application_get_FixedDecimalPlaces,
    DISP_IMP_Application_put_FixedDecimalPlaces,
    DISP_IMP_Application_GetCustomListContents,
    DISP_IMP_Application_GetCustomListNum,
    DISP_IMP_Application_GetOpenFilename,
    DISP_IMP_Application_GetSaveAsFilename,
    DISP_IMP_Application_Goto,
    DISP_IMP_Application_get_Height,
    DISP_IMP_Application_put_Height,
    DISP_IMP_Application_Help,
    DISP_IMP_Application_get_IgnoreRemoteRequests,
    DISP_IMP_Application_put_IgnoreRemoteRequests,
    DISP_IMP_Application_InchesToPoints,
    DISP_IMP_Application_InputBox,
    DISP_IMP_Application_get_Interactive,
    DISP_IMP_Application_put_Interactive,
    DISP_IMP_Application_get_International,
    DISP_IMP_Application_get_Iteration,
    DISP_IMP_Application_put_Iteration,
    DISP_IMP_Application_get_LargeButtons,
    DISP_IMP_Application_put_LargeButtons,
    DISP_IMP_Application_get_Left,
    DISP_IMP_Application_put_Left,
    DISP_IMP_Application_get_LibraryPath,
    DISP_IMP_Application_MacroOptions,
    DISP_IMP_Application_MailLogoff,
    DISP_IMP_Application_MailLogon,
    DISP_IMP_Application_get_MailSession,
    DISP_IMP_Application_get_MailSystem,
    DISP_IMP_Application_get_MathCoprocessorAvailable,
    DISP_IMP_Application_get_MaxChange,
    DISP_IMP_Application_put_MaxChange,
    DISP_IMP_Application_get_MaxIterations,
    DISP_IMP_Application_put_MaxIterations,
    DISP_IMP_Application_get_MemoryFree,
    DISP_IMP_Application_get_MemoryTotal,
    DISP_IMP_Application_get_MemoryUsed,
    DISP_IMP_Application_get_MouseAvailable,
    DISP_IMP_Application_get_MoveAfterReturn,
    DISP_IMP_Application_put_MoveAfterReturn,
    DISP_IMP_Application_get_MoveAfterReturnDirection,
    DISP_IMP_Application_put_MoveAfterReturnDirection,
    DISP_IMP_Application_get_RecentFiles,
    DISP_IMP_Application_get_Name,
    DISP_IMP_Application_NextLetter,
    DISP_IMP_Application_get_NetworkTemplatesPath,
    DISP_IMP_Application_get_ODBCErrors,
    DISP_IMP_Application_get_ODBCTimeout,
    DISP_IMP_Application_put_ODBCTimeout,
    DISP_IMP_Application_get_OnCalculate,
    DISP_IMP_Application_put_OnCalculate,
    DISP_IMP_Application_get_OnData,
    DISP_IMP_Application_put_OnData,
    DISP_IMP_Application_get_OnDoubleClick,
    DISP_IMP_Application_put_OnDoubleClick,
    DISP_IMP_Application_get_OnEntry,
    DISP_IMP_Application_put_OnEntry,
    DISP_IMP_Application_OnKey,
    DISP_IMP_Application_OnRepeat,
    DISP_IMP_Application_get_OnSheetActivate,
    DISP_IMP_Application_put_OnSheetActivate,
    DISP_IMP_Application_get_OnSheetDeactivate,
    DISP_IMP_Application_put_OnSheetDeactivate,
    DISP_IMP_Application_OnTime,
    DISP_IMP_Application_OnUndo,
    DISP_IMP_Application_get_OnWindow,
    DISP_IMP_Application_put_OnWindow,
    DISP_IMP_Application_get_OperatingSystem,
    DISP_IMP_Application_get_OrganizationName,
    DISP_IMP_Application_get_Path,
    DISP_IMP_Application_get_PathSeparator,
    DISP_IMP_Application_get_PreviousSelections,
    DISP_IMP_Application_get_PivotTableSelection,
    DISP_IMP_Application_put_PivotTableSelection,
    DISP_IMP_Application_get_PromptForSummaryInfo,
    DISP_IMP_Application_put_PromptForSummaryInfo,
    DISP_IMP_Application_Quit,
    DISP_IMP_Application_RecordMacro,
    DISP_IMP_Application_get_RecordRelative,
    DISP_IMP_Application_get_ReferenceStyle,
    DISP_IMP_Application_put_ReferenceStyle,
    DISP_IMP_Application_get_RegisteredFunctions,
    DISP_IMP_Application_RegisterXLL,
    DISP_IMP_Application_Repeat,
    DISP_IMP_Application_ResetTipWizard,
    DISP_IMP_Application_get_RollZoom,
    DISP_IMP_Application_put_RollZoom,
    DISP_IMP_Application_Save,
    DISP_IMP_Application_SaveWorkspace,
    DISP_IMP_Application_get_ScreenUpdating,
    DISP_IMP_Application_put_ScreenUpdating,
    DISP_IMP_Application_SetDefaultChart,
    DISP_IMP_Application_get_SheetsInNewWorkbook,
    DISP_IMP_Application_put_SheetsInNewWorkbook,
    DISP_IMP_Application_get_ShowChartTipNames,
    DISP_IMP_Application_put_ShowChartTipNames,
    DISP_IMP_Application_get_ShowChartTipValues,
    DISP_IMP_Application_put_ShowChartTipValues,
    DISP_IMP_Application_get_StandardFont,
    DISP_IMP_Application_put_StandardFont,
    DISP_IMP_Application_get_StandardFontSize,
    DISP_IMP_Application_put_StandardFontSize,
    DISP_IMP_Application_get_StartupPath,
    DISP_IMP_Application_get_StatusBar,
    DISP_IMP_Application_put_StatusBar,
    DISP_IMP_Application_get_TemplatesPath,
    DISP_IMP_Application_get_ShowToolTips,
    DISP_IMP_Application_put_ShowToolTips,
    DISP_IMP_Application_get_Top,
    DISP_IMP_Application_put_Top,
    DISP_IMP_Application_get_DefaultSaveFormat,
    DISP_IMP_Application_put_DefaultSaveFormat,
    DISP_IMP_Application_get_TransitionMenuKey,
    DISP_IMP_Application_put_TransitionMenuKey,
    DISP_IMP_Application_get_TransitionMenuKeyAction,
    DISP_IMP_Application_put_TransitionMenuKeyAction,
    DISP_IMP_Application_get_TransitionNavigKeys,
    DISP_IMP_Application_put_TransitionNavigKeys,
    DISP_IMP_Application_Undo,
    DISP_IMP_Application_get_UsableHeight,
    DISP_IMP_Application_get_UsableWidth,
    DISP_IMP_Application_put_UserControl,
    DISP_IMP_Application_get_UserControl,
    DISP_IMP_Application_get_UserName,
    DISP_IMP_Application_put_UserName,
    DISP_IMP_Application_get_Value,
    DISP_IMP_Application_get_VBE,
    DISP_IMP_Application_get_Version,
    DISP_IMP_Application_get_Visible,
    DISP_IMP_Application_put_Visible,
    DISP_IMP_Application_Volatile,
    DISP_IMP_Application__Wait,
    DISP_IMP_Application_get_Width,
    DISP_IMP_Application_put_Width,
    DISP_IMP_Application_get_WindowsForPens,
    DISP_IMP_Application_put_WindowState,
    DISP_IMP_Application_get_WindowState,
    DISP_IMP_Application_get_UILanguage,
    DISP_IMP_Application_put_UILanguage,
    DISP_IMP_Application_get_DefaultSheetDirection,
    DISP_IMP_Application_put_DefaultSheetDirection,
    DISP_IMP_Application_get_CursorMovement,
    DISP_IMP_Application_put_CursorMovement,
    DISP_IMP_Application_get_ControlCharacters,
    DISP_IMP_Application_put_ControlCharacters,
    DISP_IMP_Application__WSFunction,
    DISP_IMP_Application_get_EnableEvents,
    DISP_IMP_Application_put_EnableEvents,
    DISP_IMP_Application_get_DisplayInfoWindow,
    DISP_IMP_Application_put_DisplayInfoWindow,
    DISP_IMP_Application_Wait,
    DISP_IMP_Application_get_ExtendList,
    DISP_IMP_Application_put_ExtendList,
    DISP_IMP_Application_get_OLEDBErrors,
    DISP_IMP_Application_GetPhonetic,
    DISP_IMP_Application_get_COMAddIns,
    DISP_IMP_Application_get_DefaultWebOptions,
    DISP_IMP_Application_get_ProductCode,
    DISP_IMP_Application_get_UserLibraryPath,
    DISP_IMP_Application_get_AutoPercentEntry,
    DISP_IMP_Application_put_AutoPercentEntry,
    DISP_IMP_Application_get_LanguageSettings,
    DISP_IMP_Application_get_Dummy101,
    DISP_IMP_Application_Dummy12,
    DISP_IMP_Application_get_AnswerWizard,
    DISP_IMP_Application_CalculateFull,
    DISP_IMP_Application_FindFile,
    DISP_IMP_Application_get_CalculationVersion,
    DISP_IMP_Application_get_ShowWindowsInTaskbar,
    DISP_IMP_Application_put_ShowWindowsInTaskbar,
    DISP_IMP_Application_get_FeatureInstall,
    DISP_IMP_Application_put_FeatureInstall,
    DISP_IMP_Application_get_Ready,
    DISP_IMP_Application_Dummy13,
    DISP_IMP_Application_get_FindFormat,
    DISP_IMP_Application_putref_FindFormat,
    DISP_IMP_Application_get_ReplaceFormat,
    DISP_IMP_Application_putref_ReplaceFormat,
    DISP_IMP_Application_get_UsedObjects,
    DISP_IMP_Application_get_CalculationState,
    DISP_IMP_Application_get_CalculationInterruptKey,
    DISP_IMP_Application_put_CalculationInterruptKey,
    DISP_IMP_Application_get_Watches,
    DISP_IMP_Application_get_DisplayFunctionToolTips,
    DISP_IMP_Application_put_DisplayFunctionToolTips,
    DISP_IMP_Application_get_AutomationSecurity,
    DISP_IMP_Application_put_AutomationSecurity,
    DISP_IMP_Application_get_FileDialog,
    DISP_IMP_Application_Dummy14,
    DISP_IMP_Application_CalculateFullRebuild,
    DISP_IMP_Application_get_DisplayPasteOptions,
    DISP_IMP_Application_put_DisplayPasteOptions,
    DISP_IMP_Application_get_DisplayInsertOptions,
    DISP_IMP_Application_put_DisplayInsertOptions,
    DISP_IMP_Application_get_GenerateGetPivotData,
    DISP_IMP_Application_put_GenerateGetPivotData,
    DISP_IMP_Application_get_AutoRecover,
    DISP_IMP_Application_get_Hwnd,
    DISP_IMP_Application_get_Hinstance,
    DISP_IMP_Application_CheckAbort,
    DISP_IMP_Application_get_ErrorCheckingOptions,
    DISP_IMP_Application_get_AutoFormatAsYouTypeReplaceHyperlinks,
    DISP_IMP_Application_put_AutoFormatAsYouTypeReplaceHyperlinks,
    DISP_IMP_Application_get_SmartTagRecognizers,
    DISP_IMP_Application_get_NewWorkbook,
    DISP_IMP_Application_get_SpellingOptions,
    DISP_IMP_Application_get_Speech,
    DISP_IMP_Application_get_MapPaperSize,
    DISP_IMP_Application_put_MapPaperSize,
    DISP_IMP_Application_get_ShowStartupDialog,
    DISP_IMP_Application_put_ShowStartupDialog,
    DISP_IMP_Application_get_DecimalSeparator,
    DISP_IMP_Application_put_DecimalSeparator,
    DISP_IMP_Application_get_ThousandsSeparator,
    DISP_IMP_Application_put_ThousandsSeparator,
    DISP_IMP_Application_get_UseSystemSeparators,
    DISP_IMP_Application_put_UseSystemSeparators,
    DISP_IMP_Application_get_ThisCell,
    DISP_IMP_Application_get_RTD,
    DISP_IMP_Application_get_DisplayDocumentActionTaskPane,
    DISP_IMP_Application_put_DisplayDocumentActionTaskPane,
    DISP_IMP_Application_DisplayXMLSourcePane,
    DISP_IMP_Application_get_ArbitraryXMLSupportAvailable,
    DISP_IMP_Application_Support,
};


/*
**
*/


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

    _applicationexcell->pApplicationVtbl = &IMP_Application_Vtbl;
    _applicationexcell->pDispApplicationVtbl = &DISP_IMP_Application_Vtbl;
    _applicationexcell->pConnectionPointContainerVtbl = &IMPConnectionPointContainerVtbl;
    _applicationexcell->pConnectionPointVtbl = &IMPConnectionPointVtbl;
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


