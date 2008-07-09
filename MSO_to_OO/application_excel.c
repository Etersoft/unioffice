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

/*
IUnknown
*/
static ULONG WINAPI MSO_TO_OO_I_ApplicationExcel_AddRef(
        I_ApplicationExcel* iface)
{
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
    ULONG ref;

    if (This == NULL) return E_POINTER;

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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
    WCHAR str_clsid[39];

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_ApplicationExcel)) {
        *ppvObject = &This->_applicationexcellVtbl;
        MSO_TO_OO_I_ApplicationExcel_AddRef(iface);
        return S_OK;
    }
    StringFromGUID2(riid, str_clsid, 39);
    TRACE("Interface not supported\n");
    WTRACE(L" (%s) \n", str_clsid);
    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_ApplicationExcel_Release(
        I_ApplicationExcel* iface)
{
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
    ULONG ref;

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);

    TRACE("REF = %i \n", This->ref);

    if (ref == 0) {
        InterlockedDecrement(&dll_ref);
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
        VARIANT_BOOL vbDisplayAlerts)
{
   TRACE("\n");
   /*Возвращаем успех*/
   return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_DisplayAlerts(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *vbDisplayAlerts)
{
   TRACE("\n");
   /*Возвращаем успех*/
   *vbDisplayAlerts = VARIANT_FALSE;
   return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_WindowState(
        I_ApplicationExcel* iface,
        XlWindowState State)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_WindowState(
        I_ApplicationExcel* iface,
        XlWindowState *State)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_put_Visible(
        I_ApplicationExcel* iface,
        VARIANT_BOOL vbVisible)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Visible(
        I_ApplicationExcel* iface,
        VARIANT_BOOL *vbVisible)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_Workbooks(
        I_ApplicationExcel* iface,
        IDispatch **ppWorkbooks)
{
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;

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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
    I_Workbook *pwb;
    HRESULT hres;

    TRACE("\n");

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) return E_FAIL;

    hres = I_Workbook_get_Sheets(pwb, ppSheets);
    if (FAILED(hres)) return E_FAIL;

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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
    I_Workbook *pwb;
    I_Sheets *pSheets;
    I_Worksheet *pworksheet;
    HRESULT hres;

    TRACE("\n");

    hres = MSO_TO_OO_GetActiveWorkbook((I_Workbooks*)(This->pdWorkbooks), &pwb);
    if (FAILED(hres)) return E_FAIL;

    hres = I_Workbook_get_Sheets(pwb, (IDispatch**) &pSheets);
    if (FAILED(hres)) {
        *ppRange = NULL;
        return E_FAIL;
    }

    hres = MSO_TO_OO_GetActiveSheet(pSheets, &pworksheet);
    if (FAILED(hres)) {
        I_Sheets_Release(pSheets);
        *ppRange = NULL;
        return hres;
    }

    hres = I_Worksheet_get_Cells(pworksheet, ppRange);
    if (FAILED(hres)) {
        *ppRange = NULL;
    }

    I_Sheets_Release(pSheets);
    I_Worksheet_Release(pworksheet);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(
        I_ApplicationExcel* iface,
        IDispatch **RHS)
{
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;

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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;

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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;

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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;

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
        VARIANT_BOOL vbscup)
{
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_ApplicationExcel_get_ScreenUpdating(
        I_ApplicationExcel* iface,
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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
    HRESULT hres; 

    TRACE("\n");
    I_Worksheet *wsh;

    hres = MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet(iface,(IDispatch**) &wsh);

    hres = I_Worksheet_get_Range(wsh,Cell1, Cell2, ppRange);

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
        TRACE("ERROR when get_ActiveSheet\n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vRes, awb->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vRes), L"getSelection",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getSelectionr \n");
        return hres;
    }

    hres = _I_RangeConstructor((void**)&pobj);
    if (FAILED(hres)) {
        TRACE("ERROR when _I_RangeConstructor\n");
        VariantClear(&vRes);
        VariantClear(&vRet);
        I_Workbook_Release((I_Workbook*)awb);
        return E_FAIL;
    }

    hres = I_Range_QueryInterface(pobj, &IID_I_Range, (void**)ppRange);
    if (FAILED(hres)) {
        TRACE("ERROR when _I_RangeConstructor\n");
        VariantClear(&vRes);
        VariantClear(&vRet);
        I_Workbook_Release((I_Workbook*)awb);
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Range_Initialize3((I_Range*)*ppRange, V_DISPATCH(&vRet), asheet, (IDispatch*)iface);
    if (FAILED(hres)) {
        TRACE("ERROR when MSO_TO_OO_I_Range_Initialize2\n");
        VariantClear(&vRes);
        VariantClear(&vRet);
        I_Workbook_Release((I_Workbook*)awb);
        return E_FAIL;
    }

    VariantClear(&vRes);
    VariantClear(&vRet);
    I_Worksheet_Release((I_Worksheet*)asheet);
    I_Workbook_Release((I_Workbook*)awb);
    return S_OK;
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
        *rgDispId = 1;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_displayalerts)) {
        *rgDispId = 2;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_windowstate)) {
        *rgDispId = 3;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_visible)) {
        *rgDispId = 4;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_workbooks)) {
        *rgDispId = 5;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_sheets)) {
        *rgDispId = 6;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_worksheets)) {
        *rgDispId = 7;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_cells)) {
        *rgDispId = 8;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_activesheet)) {
        *rgDispId = 9;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_version)) {
        *rgDispId = 10;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_convertformula)) {
        *rgDispId = 11;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_quit)) {
        *rgDispId = 12;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_activecell)) {
        *rgDispId = 13;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = 14;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_enableevents)) {
        *rgDispId = 15;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_screenupdating)) {
        *rgDispId = 16;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_caption)) {
        *rgDispId = 17;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_activeworkbook)) {
        *rgDispId = 18;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_range)) {
        *rgDispId = 19;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_columns)) {
        *rgDispId = 20;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_rows)) {
        *rgDispId = 21;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_selection)) {
        *rgDispId = 22;
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
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)iface;
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
    case 1:
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
    case 2:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 2) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_ApplicationExcel_put_DisplayAlerts(iface, vbin);
        } else {
            return E_NOTIMPL;
        }
    case 3:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            /*преобразовываем любой тип к I4*/
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hr)) {
                TRACE(" (case 3) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            tmp = V_I4(&vtmp);
            MSO_TO_OO_I_ApplicationExcel_put_WindowState(iface, tmp);
            return S_OK;
        } else {
            return E_NOTIMPL;
        }
    case 4:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 4) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_ApplicationExcel_put_Visible(iface, vbin);
        } else {
            return E_NOTIMPL;
        }
    case 5:
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
    case 6:
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
    case 7:
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
    case 8:
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
                I_Range_put_Value((I_Range*)pretdisp, vNull, vtmp);
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
    case 9:
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
    case 10:
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
    case 11:
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
    case 12:
        return MSO_TO_OO_I_ApplicationExcel_Quit(iface);
    case 13:
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
    case 14:
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
    case 15:
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
    case 16:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hr = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 16) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_ApplicationExcel_put_ScreenUpdating(iface,vbin);
        } else {
            hr = MSO_TO_OO_I_ApplicationExcel_get_ScreenUpdating(iface,&vbin);
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
    case 17:
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
    case 18:
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
    case 19:
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
    case 20:
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
    case 21:
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
    case 22://Selection
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



const I_ApplicationExcelVtbl MSO_TO_OO_I_ApplicationExcel_Vtbl =
{
    MSO_TO_OO_I_ApplicationExcel_QueryInterface,
    MSO_TO_OO_I_ApplicationExcel_AddRef,
    MSO_TO_OO_I_ApplicationExcel_Release,
    MSO_TO_OO_I_ApplicationExcel_GetTypeInfoCount,
    MSO_TO_OO_I_ApplicationExcel_GetTypeInfo,
    MSO_TO_OO_I_ApplicationExcel_GetIDsOfNames,
    MSO_TO_OO_I_ApplicationExcel_Invoke,
    MSO_TO_OO_I_ApplicationExcel_put_UserControl,
    MSO_TO_OO_I_ApplicationExcel_get_UserControl,
    MSO_TO_OO_I_ApplicationExcel_put_DisplayAlerts,
    MSO_TO_OO_I_ApplicationExcel_get_DisplayAlerts,
    MSO_TO_OO_I_ApplicationExcel_put_WindowState,
    MSO_TO_OO_I_ApplicationExcel_get_WindowState,
    MSO_TO_OO_I_ApplicationExcel_put_Visible,
    MSO_TO_OO_I_ApplicationExcel_get_Visible,
    MSO_TO_OO_I_ApplicationExcel_get_Workbooks,
    MSO_TO_OO_I_ApplicationExcel_get_Sheets,
    MSO_TO_OO_I_ApplicationExcel_get_Worksheets,
    MSO_TO_OO_I_ApplicationExcel_get_Cells,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveSheet,
    MSO_TO_OO_I_ApplicationExcel_get_Version,
    MSO_TO_OO_I_ApplicationExcel_ConvertFormula,
    MSO_TO_OO_I_ApplicationExcel_Quit,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveCell,
    MSO_TO_OO_I_ApplicationExcel_get_Application,
    MSO_TO_OO_I_ApplicationExcel_get_EnableEvents,
    MSO_TO_OO_I_ApplicationExcel_put_EnableEvents,
    MSO_TO_OO_I_ApplicationExcel_get_ScreenUpdating,
    MSO_TO_OO_I_ApplicationExcel_put_ScreenUpdating,
    MSO_TO_OO_I_ApplicationExcel_get_Caption,
    MSO_TO_OO_I_ApplicationExcel_put_Caption,
    MSO_TO_OO_I_ApplicationExcel_get_ActiveWorkbook,
    MSO_TO_OO_I_ApplicationExcel_get_Range,
    MSO_TO_OO_I_ApplicationExcel_get_Columns,
    MSO_TO_OO_I_ApplicationExcel_get_Rows,
    MSO_TO_OO_I_ApplicationExcel_get_Selection
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

    _applicationexcell->_applicationexcellVtbl = &MSO_TO_OO_I_ApplicationExcel_Vtbl;
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

    *ppObj = &_applicationexcell->_applicationexcellVtbl;

    /*освобождаем память выделенную под строку*/
    SysFreeString(V_BSTR(&param1));
    VariantClear(&result);
    return S_OK;
}


