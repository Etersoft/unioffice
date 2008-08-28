/*
 * Special_functions
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

#include "special_functions.h"

HRESULT MSO_TO_OO_I_Workbooks_Initialize(
        I_Workbooks* iface,
        I_ApplicationExcel *app)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    TRACE("\n");

    This->pApplication = (IDispatch*)app;
/*    if (This->pApplication != NULL) I_ApplicationExcel_AddRef(This->pApplication);*/
    This->count_workbooks = 0;
    This->current_workbook = -1;
    This->pworkbook = NULL;

    return S_OK;
}


HRESULT MSO_TO_OO_I_Font_Initialize(
        I_Font* iface,
        I_Range *range)
{
    _FontImpl *This = (_FontImpl*)iface;

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->prange!=NULL) {
         I_Range_Release((I_Range*)(This->prange));
    }

    This->prange = (IDispatch*)range;
    if (This->prange != NULL) I_Range_AddRef((I_Range*)(This->prange));

    return S_OK;
}

HRESULT MSO_TO_OO_I_Interior_Initialize(
        I_Interior* iface,
        I_Range *range)
{
    InteriorImpl *This = (InteriorImpl*)iface;

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->prange!=NULL) {
        I_Range_Release((I_Range*)(This->prange));
    }

    This->prange = (IDispatch*)range;
    if (This->prange != NULL) I_Range_AddRef((I_Range*)(This->prange));

    return S_OK;
}

HRESULT MSO_TO_OO_I_Borders_Initialize(
        I_Borders* iface,
        I_Range *range)
{
    BordersImpl *This = (BordersImpl*)iface;

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->prange!=NULL) {
        I_Range_Release((I_Range*)(This->prange));
    }

    This->prange = (IDispatch*)range;
    if (This->prange != NULL) I_Range_AddRef((I_Range*)(This->prange));

    return S_OK;
}

HRESULT MSO_TO_OO_I_Border_Initialize(
        I_Border* iface,
        I_Borders *borders,
        XlBordersIndex key)
{
    BorderImpl *This = (BorderImpl*)iface;

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pborders!=NULL) {
        I_Range_Release((I_Range*)(This->pborders));
    }

    This->pborders = (IDispatch*)borders;
    if (This->pborders != NULL) I_Range_AddRef((I_Range*)(This->pborders));

    This->key = key;

    return S_OK;
}

HRESULT MSO_TO_OO_I_PageSetup_Initialize(
     I_PageSetup* pPageSetup,
     I_Worksheet* wsh)
{
    PageSetupImpl *this = (PageSetupImpl*) pPageSetup;
    WorksheetImpl *parent_wsh = (WorksheetImpl*) wsh;
    WorkbookImpl *wb = (WorkbookImpl*) parent_wsh->pwb;

    if (this->pwsheet!=NULL) {
        I_Worksheet_Release((I_Worksheet*)this->pwsheet);
    }
    this->pwsheet = (IDispatch*)wsh;
    if (this->pwsheet!=NULL) {
        I_Worksheet_AddRef((I_Worksheet*)this->pwsheet);
    }

    if (this->pApplication!=NULL) {
        I_ApplicationExcel_Release((I_ApplicationExcel*)this->pApplication);
    }
    this->pApplication = wb->pApplication;
    if (this->pApplication!=NULL) {
        I_ApplicationExcel_AddRef((I_ApplicationExcel*)this->pApplication);
    }


}

HRESULT MSO_TO_OO_I_Workbook_Initialize(
        I_Workbook* iface,
        I_ApplicationExcel *app)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    /* External AddWorkbook */
    VARIANT resultDoc;
    VARIANT param0,param1,param2,param3;
    VARIANT res;
    VARIANT varIndex, varName, vNull;
    BSTR bstrName;
    IDispatch *dpv = NULL, *wsh;
    HRESULT hres;
    IUnknown *punk = NULL;
    int count_list,delta_list;

    TRACE("\n");

    This->pApplication = (IDispatch*)app;
 /*   if (This->pApplication != NULL) I_ApplicationExcel_AddRef(This->pApplication);*/
    _ApplicationExcelImpl *Thisapp = (_ApplicationExcelImpl*)app;

    VariantInit(&param0);
    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&param3);
    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;
    V_VT(&param0) = VT_BSTR;
    V_BSTR(&param0) = SysAllocString(L"private:factory/scalc"); /* Type of created document */
/*    This->filename = SysAllocString(L""); */

    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"_blank");  /* Template */
    V_VT(&param2) = VT_I2;
    V_I2(&param2) = 0;  /* Another params count */

    long ix=0;
    MSO_TO_OO_GetDispatchPropertyValue(app, &dpv);
    if (dpv == NULL)
        return E_FAIL;
    VARIANT p1,p2;
    V_VT(&p1) = VT_BSTR;
    V_BSTR(&p1) = SysAllocString(L"Hidden");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Name", 1, p1);
    SysFreeString(V_BSTR(&p1));
    V_VT(&p2) = VT_BOOL; 
    if (Thisapp->visible==VARIANT_FALSE) V_BOOL(&p2) = VARIANT_TRUE;else V_BOOL(&p2) = VARIANT_FALSE;

    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Value", 1, p2);

    SAFEARRAY FAR* pPropVals;

    pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 1 );

    hres = SafeArrayPutElement( pPropVals, &ix, dpv );

    V_VT(&param3) = VT_ARRAY | VT_DISPATCH;
    V_ARRAY(&param3) = pPropVals;

    hres = AutoWrap(DISPATCH_METHOD, &resultDoc, Thisapp->pdOODesktop, L"loadComponentFromURL", 4, param3, param2, param1, param0);
    if (FAILED (hres)) {
    /*    I_ApplicationExcel_Release(This->pApplication);*/
        This->pApplication = NULL;
        return hres;
    }
    This->pDoc = V_DISPATCH(&resultDoc);
    IDispatch_AddRef(This->pDoc);

    /*надо создать pSheets*/
    hres = _I_SheetsConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Sheets_QueryInterface(punk, &IID_I_Sheets, (void**) &(This->pSheets));
/*    I_Sheets_Release(punk);*/
    if (FAILED(hres)) return E_NOINTERFACE;
    /*теперь инициализируем*/
    hres = MSO_TO_OO_I_Sheets_Initialize((I_Sheets*)(This->pSheets), iface);
    if (FAILED(hres)){
        TRACE("ERROR FAILED Sheets_Initialize \n");
    }
    /*теперь необходимо сделать указанное число листов*/
    I_Sheets_get_Count((I_Sheets*)(This->pSheets), &count_list);
    if (count_list>Thisapp->sheetsinnewworkbook) {
        /*Нужно удалить листы до требуемого кол-ва*/
        VariantInit(&varIndex);
        do {
            VariantClear(&varIndex);
            V_VT(&varIndex) = VT_I4;
            V_I4(&varIndex) = count_list;
            I_Sheets_get_Item((I_Sheets*)(This->pSheets), varIndex, &wsh);
            I_Worksheet_Delete((I_Worksheet*)wsh, 0);
            IDispatch_Release(wsh);
            wsh = NULL;
            I_Sheets_get_Count((I_Sheets*)(This->pSheets), &count_list);
        } while (count_list!=Thisapp->sheetsinnewworkbook);
        VariantClear(&varIndex);
    } else {
        if (count_list<Thisapp->sheetsinnewworkbook) {
            /*нужно добавить листов до требуемого кол-ва*/
            VariantInit(&varIndex);
            V_VT(&varIndex) = VT_I4;
            V_I4(&varIndex) = count_list;
            I_Sheets_get_Item((I_Sheets*)(This->pSheets), varIndex, &wsh);
            I_Worksheet_get_Name((I_Worksheet*)wsh, &bstrName);
            VariantClear(&varIndex);
            V_VT(&varIndex) = VT_I4;
            V_I4(&varIndex) = Thisapp->sheetsinnewworkbook - count_list;
            IDispatch_Release(wsh);
            wsh = NULL;
            V_VT(&varName) = VT_BSTR;
            V_BSTR(&varName) = SysAllocString(bstrName);
            I_Sheets_Add((I_Sheets*)(This->pSheets), vNull, varName, varIndex , vNull, &wsh);
            IDispatch_Release(wsh);
            wsh = NULL;
            SysFreeString(bstrName);
            VariantClear(&varIndex);
            VariantClear(&varName);
        }
    }


    /*освобождаем память выделенную под строки*/
    VariantClear(&param0);
    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&param3);
    VariantClear(&resultDoc);

    return hres;
}

HRESULT MSO_TO_OO_I_Sheets_Initialize(
        I_Sheets* iface,
        I_Workbook *wb)
{
    SheetsImpl *This = (SheetsImpl*)iface;
    VARIANT resultSheets;

    TRACE("\n");

    if (This==NULL) {
        TRACE("Object is NULL \n");
        return E_POINTER;
    }

    WorkbookImpl *Thiswb = (WorkbookImpl*)wb;
    This->pwb = (IDispatch*)wb;
/*    IDispatch_AddRef(This->pwb);*/

    VariantClear(&resultSheets);
    HRESULT hres = AutoWrap(DISPATCH_METHOD, &resultSheets, Thiswb->pDoc, L"getSheets", 0);

    This->pOOSheets = V_DISPATCH(&resultSheets);
    if (FAILED(hres)) {
        TRACE("ERROR when getSheets");
        return E_NOINTERFACE;
    }
    IDispatch_AddRef(This->pOOSheets);
    VariantClear(&resultSheets);
    return hres;
}

HRESULT MSO_TO_OO_GetActiveSheet(
        I_Sheets* iface,
        I_Worksheet **wsh)
{
    SheetsImpl *This = (SheetsImpl*)iface;

    VARIANT resultSheet;
    I_Worksheet *pworksheet = NULL;
    IUnknown *punk = NULL;
    WorkbookImpl *wb = (WorkbookImpl*)This->pwb;
    TRACE("\n");
    VARIANT res;
    HRESULT hres;
    VariantInit(&resultSheet);

/*TODO GET THE ACTIVE SHEET*/
    hres = AutoWrap(DISPATCH_METHOD, &res, wb->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &resultSheet, V_DISPATCH(&res), L"getActiveSheet",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getActiveSheet \n");
        return hres;
    }

    /*надо создать WorkSheet*/
    hres = _I_WorksheetConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Worksheet_QueryInterface(punk, &IID_I_Worksheet, (void**) &(pworksheet));

    if (FAILED(hres)) return E_NOINTERFACE;

    MSO_TO_OO_I_Worksheet_Initialize(pworksheet, (I_Workbook*)This->pwb, V_DISPATCH(&resultSheet));

    *wsh = pworksheet;
    I_Worksheet_AddRef(*wsh);
    I_Worksheet_Release(pworksheet);

    VariantClear(&res);
    VariantClear(&resultSheet);

    return S_OK;
}

HRESULT MSO_TO_OO_I_Worksheet_Initialize(
        I_Worksheet* iface,
        I_Workbook *wb,
        IDispatch *oosheet)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    IUnknown *punk = NULL;
    HRESULT hres;

    TRACE("\n");

    This->pwb = (IDispatch*)wb;
    IDispatch_AddRef(This->pwb);
    This->pOOSheet = oosheet;
    IDispatch_AddRef(This->pOOSheet);


    if (This->pAllRange != NULL) {
        IDispatch_Release(This->pAllRange);
        This->pAllRange = NULL;
    }
    /*если This->pAllRange = NULL его надо создать*/
    hres = _I_RangeConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &(This->pAllRange));

    if (FAILED(hres)) return E_NOINTERFACE;

    /*присваиваем указатель на worksheet*/
    RangeImpl *this_range = (RangeImpl*) ((I_Range*)This->pAllRange);
    this_range->pwsheet = (IDispatch*)iface;
    this_range->is_release = 0;
//    IDispatch_AddRef(this_range->pwsheet);

    WorkbookImpl *wbtemp = (WorkbookImpl*)This->pwb;
    /*присваиваем указатель на Application*/
    this_range->pApplication = wbtemp->pApplication;
    IDispatch_AddRef(this_range->pApplication);

    MSO_TO_OO_I_Range_Initialize2((I_Range*)This->pAllRange, This->pOOSheet);

    return S_OK;
}

HRESULT MSO_TO_OO_GetDispatchPropertyValue(
         I_ApplicationExcel *app,
         IDispatch** pIDispatch)
{
    /* there are many of the Open Office functions use "com.sun.star.beans.PropertyValue",
    using this method */
    HRESULT hres;
    VARIANT res;
    VARIANT objstr;

    TRACE("\n");

    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)app;

    V_VT(&objstr) = VT_BSTR;
    V_BSTR(&objstr) = SysAllocString(L"com.sun.star.beans.PropertyValue");
    hres = AutoWrap (DISPATCH_METHOD, &res, This->pdOOApp, L"Bridge_GetStruct", 1, objstr);
    if (hres == S_OK) {
        *pIDispatch = V_DISPATCH(&res);
        IDispatch_AddRef(*pIDispatch);
    } else {
        *pIDispatch = NULL;
    }
    VariantInit(&res);
    SysFreeString(V_BSTR(&objstr));
    return S_OK;
}

HRESULT MSO_TO_OO_GetDispatchHelper(
         I_ApplicationExcel *app,
         IDispatch** pIDispatch)
{
    /* there are many of the Open Office functions use "com.sun.star.frame.DispatchHelper",
    using this method */
    HRESULT hres;
    VARIANT res;
    VARIANT objstr;

    TRACE("\n");

    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)app;
    if (This==NULL) {
        return E_POINTER;
    }

    V_VT(&objstr) = VT_BSTR;
    V_BSTR(&objstr) = SysAllocString(L"com.sun.star.frame.DispatchHelper");
    hres = AutoWrap (DISPATCH_METHOD, &res, This->pdOOApp, L"CreateInstance", 1, objstr);
    if (hres == S_OK) {
        *pIDispatch = V_DISPATCH(&res);
        IDispatch_AddRef(*pIDispatch);
    } else {
        *pIDispatch = NULL;
    }
    VariantInit(&res);
    SysFreeString(V_BSTR(&objstr));
    return S_OK;
}

HRESULT MSO_TO_OO_ExecuteDispatchHelper_ActiveWorkBook(
        I_ApplicationExcel *app,
        BSTR ooCommand,
        VARIANT ooParams) /*должен быть массив*/
{
    _ApplicationExcelImpl *This = (_ApplicationExcelImpl*)app;
    HRESULT hres;
    VARIANT res;

    TRACE("\n");
    if (This==NULL) {
        return E_POINTER;
    }

    IDispatch *oodispatcher;
    hres = MSO_TO_OO_GetDispatchHelper(app, &oodispatcher);
    if (FAILED(hres)) {
        TRACE("ERROR when GetDispatchHelper\n");
        return E_FAIL;
    }

    VARIANT ooframe,param2,param3,param4;
    VariantInit(&param2);
    VariantInit(&param3);
    VariantInit(&param4);
    WorkbookImpl *wb;

    hres = I_ApplicationExcel_get_ActiveWorkbook(app,(IDispatch**) &wb);
    if (FAILED(hres)){
        TRACE("ERROR when get_ActiveWorkbook \n");
    }

    hres = AutoWrap(DISPATCH_METHOD, &res, wb->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &ooframe, V_DISPATCH(&res), L"getFrame",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getFrame \n");
        return hres;
    }

    V_VT(&param2) = VT_BSTR;
    V_BSTR(&param2) = SysAllocString(ooCommand);
    V_VT(&param3) = VT_BSTR;
    V_BSTR(&param3) = SysAllocString(L"");
    V_VT(&param4) = VT_I4;
    V_I4(&param4) = 0;
    VariantClear(&res);

    hres = AutoWrap (DISPATCH_METHOD, &res, oodispatcher, L"executeDispatch", 5, ooParams, param4, param3, param2, ooframe);
    if (FAILED(hres)) {
        TRACE("ERROR whe executeDispatch\n");
        return hres;
    }

    VariantClear(&param2);
    VariantClear(&param3);
    VariantClear(&param4);

    return S_OK;
}

HRESULT MSO_TO_OO_ExecuteDispatchHelper_WB(
        I_Workbook *wb,
        BSTR ooCommand,
        VARIANT ooParams)
{
    VARIANT ooframe,param2,param3,param4;
    VariantInit(&param2);
    VariantInit(&param3);
    VariantInit(&param4);
    WorkbookImpl *This_wb = (WorkbookImpl*)wb; 
    _ApplicationExcelImpl *This_app = (_ApplicationExcelImpl*)(This_wb->pApplication);
    HRESULT hres;
    VARIANT res;

    TRACE("\n");
    if (This_wb==NULL) {
        return E_POINTER;
    }
    if (This_app==NULL) {
        return E_POINTER;
    }

    IDispatch *oodispatcher;
    hres = MSO_TO_OO_GetDispatchHelper((I_ApplicationExcel*)This_app, &oodispatcher);
    if (FAILED(hres)) {
        TRACE("ERROR when GetDispatchHelper\n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_METHOD, &res, This_wb->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &ooframe, V_DISPATCH(&res), L"getFrame",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getFrame \n");
        return hres;
    }

    V_VT(&param2) = VT_BSTR;
    V_BSTR(&param2) = SysAllocString(ooCommand);
    V_VT(&param3) = VT_BSTR;
    V_BSTR(&param3) = SysAllocString(L"");
    V_VT(&param4) = VT_I4;
    V_I4(&param4) = 0;
    VariantClear(&res);

    hres = AutoWrap (DISPATCH_METHOD, &res, oodispatcher, L"executeDispatch", 5, ooParams, param4, param3, param2, ooframe);
    if (FAILED(hres)) {
        TRACE("ERROR whe executeDispatch\n");
        return hres;
    }

    VariantClear(&param2);
    VariantClear(&param3);
    VariantClear(&param4);

    return S_OK;
}

HRESULT MSO_TO_OO_CloseWorkbook(
         I_Workbook *wb,
         BSTR filename)
{
    WorkbookImpl *This = (WorkbookImpl*)wb;
    _ApplicationExcelImpl *this_app = (_ApplicationExcelImpl*)(This->pApplication);
    VARIANT res;
    SAFEARRAY FAR* pPropVals;
    long ix = 0;
    VARIANT p3,p2;
    HRESULT hres;
    V_VT(&p2) = VT_BOOL;
    V_BOOL(&p2) = VARIANT_TRUE;

    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object if NULL \n");
        return S_OK;
    }
    if (This->pDoc==NULL) {
        TRACE("ERROR Object if NULL \n");
        return S_OK;
    }
    if (!lstrcmpiW(filename, L"")) {
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"close", 1, p2);

        if (FAILED(hres)) TRACE("FAILED 1 CLOSE   \n"); else TRACE("SUCCESS 1 CLOSE   \n");
        IDispatch_Release(This->pDoc);
        This->pDoc = NULL;
        return S_OK;
    }

    /* Create PropertyValue with save-format-data */
    IDispatch *dpv;
    MSO_TO_OO_GetDispatchPropertyValue((I_ApplicationExcel*)(This->pApplication), &dpv);
    if (dpv == NULL)
        return E_FAIL;

    /* Set PropertyValue data */
    VARIANT p1;
    V_VT(&p1) = VT_BSTR;
    V_BSTR(&p1) = SysAllocString(L"FilterName");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Name", 1, p1);
    SysFreeString(V_BSTR(&p1));
    V_BSTR(&p1) = SysAllocString(L"MS Excel 97");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Value", 1, p1);
    SysFreeString(V_BSTR(&p1));
    /* Init params */
    pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 1 );

    hres = SafeArrayPutElement( pPropVals, &ix, dpv );

    VariantInit (&p3);
    V_VT(&p3) = VT_DISPATCH | VT_ARRAY;
    V_ARRAY(&p3) = pPropVals;

    V_BSTR(&p1) = SysAllocString(filename);

    /* Call StoreToURL for save document to file */
    AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"StoreAsURL", 2, p3, p1);

    hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"close", 1, p2);
    /* hres = AutoWrap(DISPATCH_METHOD, &res, This->pDoc, L"Dispose", 0); */
    if (FAILED(hres)) TRACE("FAILED CLOSE   \n"); else TRACE("SUCCESS CLOSE   \n");

    IDispatch_Release(This->pDoc);
    This->pDoc = NULL;
    IDispatch_Release(dpv);
    VariantClear(&res);
    return S_OK;
}

HRESULT MSO_TO_OO_I_Workbook_Initialize2(
        I_Workbook* iface,
        I_ApplicationExcel *app,
        BSTR Filename,
        VARIANT_BOOL astemplate)
{
    WorkbookImpl *This = (WorkbookImpl*)iface;
    VARIANT resultDoc;
    VARIANT param0,param1,param2,param3;
    VARIANT res;
    HRESULT hres;
    IUnknown *punk = NULL;
    IDispatch *dpv = NULL,*dpv2 = NULL;
    long ix=0;
    VARIANT p1,p2;

    TRACE("\n");

    This->pApplication = (IDispatch*)app;
/*    if (This->pApplication != NULL) I_ApplicationExcel_AddRef(This->pApplication);*/
    _ApplicationExcelImpl *Thisapp = (_ApplicationExcelImpl*)app;

    VariantInit(&param0);
    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&param3);
    V_VT(&param0) = VT_BSTR;
    V_BSTR(&param0) = SysAllocString(Filename); /* Name of document */
    /* This->filename = SysAllocString(Filename);  запоминаем имя файла */
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"_blank");  /* Template */
    /* При различных разширениях файлов надо задавать различные фильтры открытия*/


    /*Эти параметры используются по умолчанию*/
    V_VT(&param2) = VT_I2;
    V_I2(&param2) = 0;  // Another params count
    if (astemplate==VARIANT_FALSE) {
        MSO_TO_OO_GetDispatchPropertyValue(app, &dpv);
        if (dpv == NULL)
            return E_FAIL;
        V_VT(&p1) = VT_BSTR;
        V_BSTR(&p1) = SysAllocString(L"Hidden");
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Name", 1, p1);
        SysFreeString(V_BSTR(&p1));
        V_VT(&p2) = VT_BOOL;
        if (Thisapp->visible==VARIANT_FALSE) V_BOOL(&p2) = VARIANT_TRUE; else V_BOOL(&p2) = VARIANT_FALSE;
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Value", 1, p2);

        SAFEARRAY FAR* pPropVals;

        pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 1 );

        hres = SafeArrayPutElement( pPropVals, &ix, dpv );

        V_VT(&param3) = VT_ARRAY | VT_DISPATCH;
        V_ARRAY(&param3) = pPropVals;
    } else {
        /*формируем запрос на шаблон*/
        MSO_TO_OO_GetDispatchPropertyValue(app, &dpv);
        if (dpv == NULL)
            return E_FAIL;
        V_VT(&p1) = VT_BSTR;
        V_BSTR(&p1) = SysAllocString(L"AsTemplate");
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Name", 1, p1);
        SysFreeString(V_BSTR(&p1));
        V_VT(&p2) = VT_BOOL; 
        V_BOOL(&p2) = VARIANT_TRUE;
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Value", 1, p2);
        MSO_TO_OO_GetDispatchPropertyValue(app, &dpv2);
        if (dpv == NULL)
            return E_FAIL;
        VariantClear(&p1);
        VariantClear(&p2);
        V_VT(&p1) = VT_BSTR;
        V_BSTR(&p1) = SysAllocString(L"Hidden");
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv2, L"Name", 1, p1);
        SysFreeString(V_BSTR(&p1));
        V_VT(&p2) = VT_BOOL;
        if (Thisapp->visible==VARIANT_FALSE) V_BOOL(&p2) = VARIANT_TRUE; else V_BOOL(&p2) = VARIANT_FALSE;
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv2, L"Value", 1, p2);

        SAFEARRAY FAR* pPropVals;

        pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 2 );

        hres = SafeArrayPutElement( pPropVals, &ix, dpv );
        ix++;
        hres = SafeArrayPutElement( pPropVals, &ix, dpv2 );

        V_VT(&param3) = VT_ARRAY | VT_DISPATCH;
        V_ARRAY(&param3) = pPropVals;
        VariantClear(&p1);
        VariantClear(&p2);
    }

    hres = AutoWrap(DISPATCH_METHOD, &resultDoc, Thisapp->pdOODesktop, L"loadComponentFromURL", 4, param3, param2, param1, param0);
    if (FAILED (hres)) {
        TRACE("Ne udalos` zagruzit \n");
        WTRACE(L"Filename = %s \n", Filename);
/*        I_ApplicationExcel_Release(This->pApplication);*/
        This->pApplication = NULL;
        return hres;
    }
    This->pDoc = V_DISPATCH(&resultDoc);
    IDispatch_AddRef(This->pDoc);

    /*надо создать pSheets*/
    hres = _I_SheetsConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Sheets_QueryInterface(punk, &IID_I_Sheets, (void**) &(This->pSheets));

    if (FAILED(hres)) return E_NOINTERFACE;
    /*теперь инициализируем*/
    hres = MSO_TO_OO_I_Sheets_Initialize((I_Sheets*)(This->pSheets), iface);

    VariantClear(&param0);
    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&param3);
    VariantClear(&resultDoc);

    return hres;
}

HRESULT MSO_TO_OO_I_Range_Initialize(
         I_Range *iface,
         I_Range *pParentRange,
         struct CELL_COORD topLeft,
         struct CELL_COORD bottomRight)
{
    RangeImpl *This = (RangeImpl*)iface;
    RangeImpl *This_parent = (RangeImpl*)pParentRange;

    if (This_parent->pOORange == NULL) {
       return E_POINTER;
    }

    VARIANT resRange;
    VariantInit (&resRange);
    VARIANT vLeft, vRight, vTop, vBottom;
    VariantInit(&vLeft);
    V_VT(&vLeft) = VT_I4;
    V_I4(&vLeft) = topLeft.x - 1; /* вычитаем 1, т.к. нумерация с нуля */
    VariantInit(&vTop);
    V_VT(&vTop) = VT_I4;
    V_I4(&vTop) = topLeft.y - 1;
    VariantInit(&vRight);
    V_VT(&vRight) = VT_I4;
    V_I4(&vRight) = bottomRight.x - 1;
    VariantInit(&vBottom);
    V_VT(&vBottom) = VT_I4;
    V_I4(&vBottom) = bottomRight.y - 1;

    TRACE("\n");

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &resRange, This_parent->pOORange, L"getCellRangeByPosition", 4, vBottom, vRight, vTop, vLeft);
    if (FAILED(hres)) {
       This->pOORange = NULL;
       return hres;
    }
    /*Присваиваем указатель на worksheet*/
    This->pwsheet = This_parent->pwsheet;
    IDispatch_AddRef(This->pwsheet);
    /*Присваиваем указатель на Application*/
    This->pApplication = This_parent->pApplication;
    IDispatch_AddRef(This->pApplication);

    This->pOORange = V_DISPATCH(&resRange);
    IDispatch_AddRef(V_DISPATCH(&resRange));
    VariantClear(&resRange);
    return S_OK;
}

HRESULT MSO_TO_OO_I_Range_Initialize2(
        I_Range* iface,
        IDispatch *oosheet)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pOORange != NULL) {
        IDispatch_Release(This->pOORange);
        This->pOORange = NULL;
    }
    This->pOORange = oosheet;
    if (This->pOORange != NULL) {
        IDispatch_AddRef(This->pOORange);
        return S_OK;
    }
    return E_POINTER;
}

HRESULT MSO_TO_OO_I_Range_Initialize3(
        I_Range* iface,
        IDispatch *oosheet,
        IDispatch *psheet,
        IDispatch *pApp)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pOORange != NULL) {
        IDispatch_Release(This->pOORange);
        This->pOORange = NULL;
    }
    This->pOORange = oosheet;
    if (This->pOORange != NULL) {
        IDispatch_AddRef(This->pOORange);
        return S_OK;
    }

    This->pApplication = pApp;
    IDispatch_AddRef(pApp);
    This->pwsheet = psheet;
    IDispatch_AddRef(psheet);

    return E_POINTER;
}

HRESULT MSO_TO_OO_GetRangeAddress(
        I_Range* iface,
        long *lLeft,
        long *lTop,
        long *lRight,
        long *lBottom)
{
    RangeImpl *This = (RangeImpl*)iface;

    if (This==NULL) {
        TRACE("Error = Object is NULL \n");
        return E_FAIL;
    }

    IDispatch *pdRangeAddress = NULL;
    VARIANT vRes;
    HRESULT hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"getRangeAddress", 0);
    if (hres != S_OK) {
        return hres;
    }
    pdRangeAddress = V_DISPATCH(&vRes);
    IDispatch_AddRef(pdRangeAddress);
    VariantClear(&vRes);
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdRangeAddress, L"StartRow", 0);
    if (hres != S_OK) {
        IDispatch_Release(pdRangeAddress);
        return hres;
    }
    *lLeft = vRes.intVal;
    VariantClear(&vRes);
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdRangeAddress, L"StartColumn", 0);
    if (hres != S_OK) {
        IDispatch_Release(pdRangeAddress);
        return hres;
    }
    *lTop = vRes.intVal;
    VariantClear(&vRes);
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdRangeAddress, L"EndRow", 0);
    if (hres != S_OK) {
        IDispatch_Release(pdRangeAddress);
        return hres;
    }
    *lRight = vRes.intVal;
    VariantClear(&vRes);
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdRangeAddress, L"EndColumn", 0);
    if (hres != S_OK) {
        IDispatch_Release(pdRangeAddress);
        return hres;
    }
    *lBottom = vRes.intVal;

    IDispatch_Release(pdRangeAddress);
    return hres;
}

HRESULT MSO_TO_OO_GetActiveWorkbook(
        I_Workbooks* iface,
        I_Workbook **wb)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;

    *wb = (I_Workbook*)This->pworkbook[This->current_workbook];
    I_Workbook_AddRef(*wb);

    return S_OK;
}

HRESULT MSO_TO_OO_GetActiveCells(
        I_Workbooks* iface,
        I_Range **ppRange)
{
    IDispatch *pWorkbook;
    HRESULT hres;

    hres = MSO_TO_OO_GetActiveWorkbook(iface,(I_Workbook**) &pWorkbook);

    WorkbookImpl *wb = (WorkbookImpl*)pWorkbook;

    if ((pWorkbook == NULL) || (wb->pDoc == NULL)) {
        return E_FAIL;
    }

    VARIANT vRes;
    IDispatch *pdCurrentSelection;
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, wb->pDoc, L"CurrentSelection", 0);
    if (hres != S_OK) {
        TRACE("Error when getting CurrentSelection \n");
        return hres;
    }

    pdCurrentSelection = V_DISPATCH(&vRes);
    VARIANT vRow;
    V_VT(&vRow) = VT_I2;
    V_I2(&vRow) = 0;
    VARIANT vColumn;
    V_VT(&vColumn) = VT_I2;
    V_I2(&vColumn) = 0;
    IDispatch *pCurrentCell;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, pdCurrentSelection, L"getCellByPosition", 2, vRow, vColumn);

    IDispatch_Release(pdCurrentSelection);
    if (hres != S_OK) {
        TRACE("Error when getCellByPosition \n");
        return hres;
    }

    pCurrentCell = V_DISPATCH(&vRes);

    /*Создаем новый объект I_Range*/
    IDispatch *pRange;
    IUnknown *punk = NULL;

    hres = _I_RangeConstructor((LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);

    if (pRange == NULL) {
        return E_FAIL;
    }

    RangeImpl *this_range = (RangeImpl*) ((I_Range*)pRange);
    /*Присваиваем указатель на parent worksheet*/
    I_Sheets *pSheets;
    I_Worksheet *pworksheet;

    hres = I_Workbook_get_Sheets((I_Workbook*)pWorkbook, (IDispatch**) &pSheets);
    if (FAILED(hres)) {
        return E_FAIL;
    }

    hres = MSO_TO_OO_GetActiveSheet(pSheets, &pworksheet);
    if (FAILED(hres)) {
       I_Sheets_Release(pSheets);
       return hres;
    }


    this_range->pwsheet = (IDispatch*)pworksheet;
    IDispatch_AddRef(this_range->pwsheet);
    /*Присваиваем указатель на parent worksheet*/
    this_range->pApplication = wb->pApplication;
    IDispatch_AddRef(this_range->pApplication);

    hres = MSO_TO_OO_I_Range_Initialize2((I_Range*)pRange,pCurrentCell);

    *ppRange = (I_Range*)pRange;
    I_Range_AddRef(*ppRange);
    I_Range_Release((I_Range*)pRange);

    /*Освобождаем память*/
    I_Workbook_Release((I_Workbook*)pWorkbook);
    IDispatch_Release(pCurrentCell);
    I_Sheets_Release(pSheets);
    I_Worksheet_Release(pworksheet);
    return hres;
}

HRESULT MSO_TO_OO_I_Range_Initialize_ByName(
         I_Range *iface,
         I_Range *pParentRange,
         VARIANT rangename)
{
    RangeImpl *This = (RangeImpl*)iface;
    RangeImpl *This_parent = (RangeImpl*)pParentRange;

    if (This_parent->pOORange == NULL) {
       return E_POINTER;
    }

    VARIANT resRange;
    VariantInit (&resRange);

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &resRange, This_parent->pOORange, L"getCellRangeByName", 1, rangename);
    if (hres != S_OK) {
        This->pOORange = NULL;
        return hres;
    }

    /*Присваиваем указатель на worksheet*/
    This->pwsheet = This_parent->pwsheet;
    IDispatch_AddRef(This->pwsheet);
    /*Присваиваем указатель на worksheet*/
    This->pApplication = This_parent->pApplication;
    IDispatch_AddRef(This->pApplication);

    This->pOORange = V_DISPATCH(&resRange);
    IDispatch_AddRef(This->pOORange);
    VariantClear(&resRange);
    return S_OK;
}

HRESULT MSO_TO_OO_CorrectArg(
         VARIANT value,
         VARIANT *retval)
{
VariantInit(retval);
if (V_ISBYREF(&value)) {
    switch(V_VT(&value) - VT_BYREF) {
    case VT_EMPTY:{
        V_VT(retval) = VT_EMPTY;
        break;
        }
    case VT_NULL:{
        V_VT(retval) = VT_NULL;
        break;
        }
    case VT_I2:{
        V_VT(retval) = VT_I2;
        V_I2(retval) = *(V_I2REF(&value));
        break;
        }
    case VT_I4:{
        V_VT(retval) = VT_I4;
        V_I4(retval) = *(V_I4REF(&value));
        break;
        }
    case VT_I8:{
        V_VT(retval) = VT_I8;
        V_I8(retval) = *(V_I4REF(&value));
        break;
        }
    case VT_R4:{
        V_VT(retval) = VT_R4;
        V_R4(retval) = *(V_R4REF(&value));
        break;
        }
    case VT_R8:{
        V_VT(retval) = VT_R8;
        V_R8(retval) = *(V_R8REF(&value));
        break;
        }
    case VT_CY:{
        V_VT(retval) = VT_CY;
        V_CY(retval) = *(V_CYREF(&value));
        break;
        }
    case VT_DATE:{
        V_VT(retval) = VT_DATE;
        V_DATE(retval) = *(V_DATEREF(&value));
        break;
        }
    case VT_BSTR:{
        V_VT(retval) = VT_BSTR;
        V_BSTR(retval) = *(V_BSTRREF(&value));
        break;
        }
    case VT_DISPATCH:{
        V_VT(retval) = VT_DISPATCH;
        V_DISPATCH(retval) = *(V_DISPATCHREF(&value));
        break;
        }
    case VT_BOOL:{
        V_VT(retval) = VT_BOOL;
        V_BOOL(retval) = *(V_BOOLREF(&value));
        break;
        }
    case VT_VARIANT:{
/*        V_VT(retval) = V_VT(V_VARIANTREF(&value));*/
/*        V_DISPATCH(retval) = V_DISPATCH(V_VARIANTREF(&value));*/
        *retval = *(V_VARIANTREF(&value));
        break;
        }
    case VT_UNKNOWN:{
        VariantCopy((VARIANT*)V_UNKNOWNREF(&value),retval);
/*        V_VT(retval) = VT_UNKNOWN;
        V_UNKNOWN(retval) = *(V_UNKNOWNREF(&value));*/
        break;
        }
    case VT_UI1:{
        V_VT(retval) = VT_UI1;
        V_UI1(retval) = *(V_UI1REF(&value));
        break;
        }
    case VT_ERROR:{
        V_VT(retval) = VT_ERROR;
        break;
        }
    case VT_ARRAY:{
        V_VT(retval) = VT_ARRAY;
        V_ARRAY(retval) =*(V_ARRAYREF(&value));
        break;
        }
    }
} else {
    *retval = value; 
}
return S_OK;
}

WCHAR* insert(WCHAR* src,WCHAR* dst,unsigned int index)
{
    WCHAR* res;
    int len;
    unsigned int i,j,k;

    len = lstrlenW(src);
    len = len + lstrlenW(dst) + 1;
    res = (WCHAR*) malloc(sizeof(WCHAR)*len);
    j=0;
    for (i=0;i<lstrlenW(src);i++) {
        if (i==index) {
            for (k=0;k<lstrlenW(dst);k++) {
                *(res+j)=*(dst+k);
                j++;
            }
        }
        *(res+j)=*(src+i);
        j++;
    }
    *(res+len-1)=0;
    return res;
}

int strcmpnW(WCHAR *str1, WCHAR *str2)
{
    int i=0;
    TRACE("\n");
    while (*(str2+i)!=0) {
        if (*(str2+i)!=*(str1+i)) return 0;
        i++;
    }
    return 1;
}

HRESULT MSO_TO_OO_MakeURLFromFilename(
         BSTR value,
         BSTR *retval)
{
    int i;
    WCHAR *ptr;
    WCHAR *tmp1,tmp2[] = {'2','0',0};
    WCHAR file_str[] = {'f','i','l','e',':','/','/','l','o','c','a','l','h','o','s','t','/',0};
    WCHAR http[] = {'h','t','t','p',0};
    WCHAR https[] = {'h','t','t','p','s',0};
    WCHAR ftp[] = {'f','t','p',0};

    ptr = SysAllocString(value);
    TRACE("%i \n", strcmpnW(ptr, http));
    TRACE("%i \n", strcmpnW(ptr, https));
    TRACE("%i \n", strcmpnW(ptr, ftp));
    if ((strcmpnW(ptr, http)+strcmpnW(ptr, https)+strcmpnW(ptr, ftp))==0) {
        i=0;
        while (*(ptr+i)!=0) {
            if (*(ptr+i)==' ') {
                *(ptr+i)='%';
            tmp1=insert(ptr,tmp2,i+1);
            ptr = tmp1;
        }
        if (*(ptr+i) == '\\')
            *(ptr+i) = '/';
        i++;
        }
        tmp1=insert(ptr,file_str,0);
        ptr = tmp1;
        }

    *retval = ptr;

    return S_OK;
}


//возвращает индекс таблицы по ее имени.
//Если не находит возвращает -1
long MSO_TO_OO_FindIndexWorksheetByName(
        I_Sheets *iface,
        BSTR name)
{
    int i,count;
    SheetsImpl *This = (SheetsImpl*)iface;
    HRESULT hres;
    IDispatch *wsh;
    VARIANT par_tmp;
    BSTR tmp_name;

    TRACE("\n");
    if (This==NULL) {
        return E_POINTER;
    }

/*Необходимо заменять запятую на подчеркивание, т.к. OO не поддерживает запятые*/
    i=0;
    while (*(name+i)!=0) {
        if (*(name+i)==L',') *(name+i)=L'_';
        WTRACE(L"%c",*(name+i));
        i++;
    }
    TRACE("\n");

    hres = I_Sheets_get_Count(iface, &count);
    if (FAILED(hres)) {
        TRACE("ERROR when Sheets_get_Count\n");
        return -1;
    }
    i=1;
    while (i<=count) {
        VariantClear(&par_tmp);
        V_VT(&par_tmp) = VT_I4;
        V_I4(&par_tmp) = i;
        hres = I_Sheets_get__Default(iface, par_tmp, &wsh);
        if (!FAILED(hres)) {
            hres = I_Worksheet_get_Name((I_Worksheet*)wsh, &tmp_name);
            if (!FAILED(hres)) {
                if (!lstrcmpiW(tmp_name, name)) {
                    SysFreeString(tmp_name);
                    IDispatch_Release(wsh);
                    return i - 1;
                }
                SysFreeString(tmp_name);
            }
            IDispatch_Release(wsh);
            wsh = NULL;
        }
        i++;
    }
    return -1;
}

/*возвращает Workbook и индекс*/
long MSO_TO_OO_GlobalFindIndexWorksheetByName(
        I_ApplicationExcel *app,
        BSTR name,
        IDispatch **retval)
{
    _ApplicationExcelImpl *This_app = (_ApplicationExcelImpl*)app;
    int i,id;
    WorkbooksImpl *wbs = (WorkbooksImpl*)This_app->pdWorkbooks;
    SheetsImpl *wsheets;
    WorkbookImpl *wb;
    for (i=0;i<wbs->count_workbooks;i++){
        if (wbs->pworkbook[i]!=NULL) {
            wb = (WorkbookImpl*)(wbs->pworkbook[i]);
            wsheets = (SheetsImpl*)wb->pSheets;
            id = MSO_TO_OO_FindIndexWorksheetByName((I_Sheets*)wsheets, name);
            if (id>=0) {
               *retval = (IDispatch*)wb;
               return id;
            }
        }
    }
return -1;
}

HRESULT MSO_TO_OO_I_Shapes_Initialize(
        I_Shapes* iface,
        I_Worksheet *iwsh)
{
    ShapesImpl *This = (ShapesImpl*)iface;
    WorksheetImpl *wsh = (WorksheetImpl*)iwsh;
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);
    HRESULT hres;
    VARIANT vframe, param1, vRet;

    TRACE("\n");

    VariantInit(&vframe);
    VariantInit(&vRet);
    VariantInit(&param1);

    if (This->pwsheet!=NULL) {
        I_Worksheet_Release((I_Worksheet*)This->pwsheet);
    }
    This->pwsheet = (IDispatch*)wsh;
    I_Worksheet_AddRef((I_Worksheet*)This->pwsheet);

    if (This->pApplication!=NULL) {
        I_ApplicationExcel_Release((I_ApplicationExcel*)This->pApplication);
    }
    This->pApplication = (IDispatch*)wb->pApplication;
    I_ApplicationExcel_AddRef((I_ApplicationExcel*)This->pApplication);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vframe, wb->pDoc, L"DrawPages",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when get DrawPages \n");
        return E_FAIL;
    }
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = 1;
    hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vframe), L"insertNewByIndex",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when insertNewByIndex \n");
        return E_FAIL;
    }

    V_VT(&param1) = VT_I4;
    V_I4(&param1) = 0;
    hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vframe), L"getByIndex",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when get getByIndex \n");
        return E_FAIL;
    }

    if (This->pOOPage!=NULL) {
        IDispatch_Release(This->pOOPage);
    }
    This->pOOPage = V_DISPATCH(&vRet);
    IDispatch_AddRef(This->pOOPage);

    VariantClear(&vframe);
    VariantClear(&vRet);
    VariantClear(&param1);
    return S_OK;
}

HRESULT MSO_TO_OO_I_Shape_Line_Initialize(
        I_Shape* iface,
        I_Shapes *ishapes,
        float x1, float y1, float x2, float y2)
{
    ShapeImpl *This = (ShapeImpl*)iface;
    ShapesImpl *shapes = (ShapesImpl*)ishapes;
    WorksheetImpl *wsh = (WorksheetImpl*)(shapes->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);
    VARIANT vline, param1, size, position, vRet;
    HRESULT hres;

    TRACE("\n");

    VariantInit(&vline);
    VariantInit(&param1);
    VariantInit(&position);
    VariantInit(&size);
    VariantInit(&vRet);

    if (This->pShapes!=NULL) {
        I_Shapes_Release((I_Shapes*)This->pShapes);
    }
    This->pShapes = (IDispatch*)ishapes;
    I_Shapes_AddRef((I_Shapes*)This->pShapes);

    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"com.sun.star.drawing.LineShape");
    hres = AutoWrap(DISPATCH_METHOD, &vline, wb->pDoc, L"createInstance",1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when get createInstance \n");
        return E_FAIL;
    }
    if (This->pOOShape!=NULL) {
        IDispatch_Release(This->pOOShape);
    }
    This->pOOShape = V_DISPATCH(&vline);
    IDispatch_AddRef(This->pOOShape);

    hres = AutoWrap(DISPATCH_METHOD, &position, V_DISPATCH(&vline), L"getPosition",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getPosition \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_METHOD, &size, V_DISPATCH(&vline), L"getSize",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getSize\n");
        return E_FAIL;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_R4;
    V_R4(&param1) = x1;
    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRet, V_DISPATCH(&position), L"X",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when set X\n");
        return E_FAIL;
    }
    VariantClear(&vRet);
    VariantClear(&param1);
    V_VT(&param1) = VT_R4;
    V_R4(&param1) = y1;
    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRet, V_DISPATCH(&position), L"Y",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when set Y\n");
        return E_FAIL;
    }
    VariantClear(&vRet);
    VariantClear(&param1);
    V_VT(&param1) = VT_R4;
    V_R4(&param1) = x2-x1;
    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRet, V_DISPATCH(&size), L"Width",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when set Width\n");
        return E_FAIL;
    }
    VariantClear(&vRet);
    VariantClear(&param1);
    V_VT(&param1) = VT_R4;
    V_R4(&param1) = y2-y1;
    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRet, V_DISPATCH(&size), L"Height",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when set Height\n");
        return E_FAIL;
    }
    VariantClear(&vRet);
    VariantClear(&param1);
    V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = V_DISPATCH(&position);
    IDispatch_AddRef(V_DISPATCH(&param1));
    hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vline), L"setPosition",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when setPosition \n");
        return E_FAIL;
    }
    VariantClear(&vRet);
    VariantClear(&param1);
    V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = V_DISPATCH(&size);
    IDispatch_AddRef(V_DISPATCH(&param1));
    hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vline), L"setSize",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when setSize\n");
        return E_FAIL;
    }

    /*add shape to page*/
    VariantClear(&vRet);
    VariantClear(&param1);
    V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = This->pOOShape;
    IDispatch_AddRef(This->pOOShape);
    hres = AutoWrap(DISPATCH_METHOD, &vRet, shapes->pOOPage, L"add",1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when add\n");
        return E_FAIL;
    }

    VariantClear(&vline);
    VariantClear(&param1);
    VariantClear(&position);
    VariantClear(&size);
    VariantClear(&vRet);
    return S_OK;
}

HRESULT MSO_TO_OO_Names_Initialize(
        Names* iface,
        I_Workbook *wb)
{
    NamesImpl *This = (NamesImpl*)iface;
    WorkbookImpl *wbi = (WorkbookImpl*)wb;
    VARIANT vRet;
    HRESULT hres;

    TRACE("\n");

    VariantInit(&vRet);

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pwb!=NULL) {
         I_Workbook_Release((I_Workbook*)(This->pwb));
    }
    This->pwb = (IDispatch*)wb;
    if (This->pwb != NULL) I_Workbook_AddRef((I_Workbook*)(This->pwb));

    if (This->pApplication!=NULL) {
         I_ApplicationExcel_Release((I_ApplicationExcel*)(This->pApplication));
    }
    This->pApplication = (IDispatch*)(wbi->pApplication);
    if (This->pApplication != NULL) I_ApplicationExcel_AddRef((I_ApplicationExcel*)(This->pApplication));

    if (wbi->pDoc==NULL) {
        TRACE("Object pDoc is NULL\n");
        return E_FAIL;
    }
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRet, wbi->pDoc, L"NamedRanges",0);
    if (FAILED(hres)) {
        TRACE("ERROR when NamedRanges \n");
        return E_FAIL;
    }
    if (This->pOONames!=NULL) {
         IDispatch_Release(This->pOONames);
    }
    This->pOONames = V_DISPATCH(&vRet);
    if (This->pOONames != NULL) IDispatch_AddRef(This->pOONames);

    VariantClear(&vRet);

    return S_OK;
}

HRESULT MSO_TO_OO_Workbook_SetVisible(
        I_Workbook *wb,
        VARIANT_BOOL vbvisible)
{
    WorkbookImpl *This = (WorkbookImpl*)wb;
    HRESULT hres;
    VARIANT oocontr, ooframe, oocontwindow, param, res;

    TRACE("\n");

    VariantInit(&oocontr);
    VariantInit(&ooframe);
    VariantInit(&oocontwindow);
    VariantInit(&param);
    VariantInit(&res);

    if (This==NULL) {
        TRACE("Object is NULL \n");
        return S_OK;
    }

    if (This->pDoc==NULL) {
        TRACE("Object pDoc is NULL \n");
        return S_OK;
    }

    hres = AutoWrap(DISPATCH_METHOD, &oocontr,This->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &ooframe, V_DISPATCH(&oocontr), L"getFrame",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getFrame \n");
        VariantClear(&oocontr);
        return hres;
    }
    hres = AutoWrap(DISPATCH_METHOD, &oocontwindow, V_DISPATCH(&ooframe), L"getContainerWindow",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getContainerWindow \n");
        VariantClear(&oocontr);
        VariantClear(&ooframe);
        return hres;
    }

    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = vbvisible;

    hres = AutoWrap(DISPATCH_METHOD, &res, V_DISPATCH(&oocontwindow), L"setVisible",1,param);
    if (FAILED(hres)) {
        TRACE("ERROR when setVisible \n");
        VariantClear(&oocontr);
        VariantClear(&ooframe);
        VariantClear(&oocontwindow);
        VariantClear(&param);
        return hres;
    }

    VariantClear(&oocontr);
    VariantClear(&ooframe);
    VariantClear(&oocontwindow);
    VariantClear(&param);
    VariantClear(&res);

    return S_OK;
}

HRESULT MSO_TO_OO_I_Outline_Initialize(
        I_Outline* iface,
        I_Worksheet *iwsh)
{
    OutlineImpl *This = (OutlineImpl*)iface;

    if (This->pwsh!=NULL) {
        I_Worksheet_Release((I_Worksheet*)This->pwsh);
    }
    This->pwsh = (IDispatch*)iwsh;
    I_Worksheet_AddRef((I_Worksheet*)This->pwsh);

    return S_OK;
}
