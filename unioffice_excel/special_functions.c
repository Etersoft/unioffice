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


#define WORKBOOKS_THIS(iface) DEFINE_THIS(WorkbooksImpl, workbooks, iface)
HRESULT MSO_TO_OO_I_Workbooks_Initialize(
        I_Workbooks* iface,
        _Application *app)
{
    WorkbooksImpl *This = WORKBOOKS_THIS(iface);
    TRACE_IN;

    This->pApplication = app;
/*    if (This->pApplication != NULL) _Application_AddRef(This->pApplication);*/
    This->count_workbooks = 0;
    This->current_workbook = -1;
    This->pworkbook = NULL;

    TRACE_OUT;
    return S_OK;
}
#undef WORKBOOKS_THIS


#define FONT_THIS(iface) DEFINE_THIS(FontImpl, font, iface)
#define RANGE_THIS(iface) DEFINE_THIS(RangeImpl, range, iface)
HRESULT MSO_TO_OO_I_Font_Initialize(
        I_Font* iface,
        I_Range *range)
{
    FontImpl *This = FONT_THIS(iface);
    RangeImpl *This_range = RANGE_THIS(range);
    TRACE_IN;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    if (This->pRange) {
         I_Range_Release(This->pRange);
    }
    This->pRange = range;
    if (This->pRange) I_Range_AddRef(This->pRange);

    if (This->pOORange) {
         IDispatch_Release(This->pOORange);
    }
    This->pOORange = This_range->pOORange;
    if (This->pOORange) IDispatch_AddRef(This->pOORange);


    TRACE_OUT;
    return S_OK;
}
#undef RANGE_THIS
#undef FONT_THIS

#define INTERIOR_THIS(iface) DEFINE_THIS(InteriorImpl, interior, iface)
#define RANGE_THIS(iface) DEFINE_THIS(RangeImpl, range, iface)
HRESULT MSO_TO_OO_I_Interior_Initialize(
        I_Interior* iface,
        I_Range *range)
{
    InteriorImpl *This = INTERIOR_THIS(iface);
    RangeImpl *This_range = RANGE_THIS(range);
    TRACE_IN;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }
    
    if (!This_range) {
        ERR("This_range is NULL \n");
        return E_POINTER;
    }
    
    if (This->pRange) {
        I_Range_Release(This->pRange);
    }
    This->pRange = range;
    if (This->pRange) I_Range_AddRef((This->pRange));

    if (This->pOORange) {
        IDispatch_Release(This->pOORange);
    }
    This->pOORange = This_range->pOORange;
    if (This->pOORange) IDispatch_AddRef((This->pOORange));


    TRACE_OUT;
    return S_OK;
}
#undef RANGE_THIS
#undef INTERIOR_THIS

#define BORDERS_THIS(iface) DEFINE_THIS(BordersImpl, borders, iface)
HRESULT MSO_TO_OO_I_Borders_Initialize(
        I_Borders* iface,
        I_Range *range)
{
    BordersImpl *This = BORDERS_THIS(iface);
    RangeImpl* This_range = (RangeImpl*)range;
    TRACE_IN;

    if (!This) {
        ERR("object is NULL \n");
        return E_POINTER;
    }

    if (This->pRange) {
        I_Range_Release((This->pRange));
    }
    This->pRange = range;
    if (This->pRange) I_Range_AddRef(This->pRange);

    if (This->pOORange) {
        IDispatch_Release(This->pOORange);
    }
    This->pOORange = This_range->pOORange;
    if (This->pOORange) IDispatch_AddRef((This->pOORange));

    TRACE_OUT;
    return S_OK;
}
#undef BORDERS_THIS

#define BORDER_THIS(iface) DEFINE_THIS(BorderImpl, border, iface)
#define BORDERS_THIS(iface) DEFINE_THIS(BordersImpl, borders, iface)
HRESULT MSO_TO_OO_I_Border_Initialize(
        I_Border* iface,
        I_Borders *borders,
        XlBordersIndex key)
{
    BorderImpl *This = BORDER_THIS(iface);
    BordersImpl *This_borders = BORDERS_THIS(borders);
    TRACE_IN;

    if (!This) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pBorders) {
        I_Borders_Release(This->pBorders);
    }
    This->pBorders = borders;
    if (This->pBorders) 
        I_Borders_AddRef(This->pBorders);
    else {
        TRACE("ERROR parent object is NULL \n");
        return E_FAIL;    
    }

    if (This->pOORange) {
        IDispatch_Release(This->pOORange);
    }
    This->pOORange = This_borders->pOORange;
    if (This->pOORange) 
        IDispatch_AddRef(This->pOORange);
    else {
        TRACE("ERROR OORange object is NULL \n");
        return E_FAIL;    
    }

    This->key = key;

    TRACE_OUT;
    return S_OK;
}
#undef BORDER_THIS
#undef BORDERS_THIS

#define PAGESETUP_THIS(iface) DEFINE_THIS(PageSetupImpl, pagesetup, iface)
#define WORKSHEET_THIS(iface) DEFINE_THIS(WorksheetImpl, worksheet, iface)
#define WORKBOOK_THIS(iface) DEFINE_THIS(WorkbookImpl, workbook, iface)
HRESULT MSO_TO_OO_I_PageSetup_Initialize(
     I_PageSetup* pPageSetup,
     I_Worksheet* wsh)
{
    PageSetupImpl *This = PAGESETUP_THIS(pPageSetup);
    WorksheetImpl *This_worksheet = WORKSHEET_THIS(wsh);
    WorkbookImpl *This_workbook = WORKBOOK_THIS(This_worksheet->pwb);
    
    TRACE_IN;

    if (This->pWorksheet) {
        I_Worksheet_Release(This->pWorksheet);
    }
    This->pWorksheet = wsh;
    if (This->pWorksheet) {
        I_Worksheet_AddRef(This->pWorksheet);
    }

    if (This->pOOSheet) {
        IDispatch_Release(This->pOOSheet);
    }
    This->pOOSheet = This_worksheet->pOOSheet;
    if (This->pOOSheet) {
        IDispatch_AddRef(This->pOOSheet);
    }

    if (This->pOODocument) {
        IDispatch_Release(This->pOODocument);
    }
    This->pOODocument = This_workbook->pDoc;
    if (This->pOODocument) {
        IDispatch_AddRef(This->pOODocument);
    }

    TRACE_OUT;
    return S_OK;
}
#undef WORKBOOK_THIS
#undef WORKSHEET_THIS
#undef PAGESETUP_THIS

HRESULT MSO_TO_OO_I_Workbook_Initialize(
        I_Workbook* iface,
        I_Workbooks *pwrks)
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
    TRACE_IN;

    This->pworkbooks = (IDispatch*)pwrks;
    if (This->pworkbooks != NULL) I_Workbooks_AddRef(pwrks);

    WorkbooksImpl *wbks = (WorkbooksImpl*)pwrks;
    _ApplicationImpl *Thisapp = (_ApplicationImpl*)wbks->pApplication;

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
    MSO_TO_OO_GetDispatchPropertyValue(APPEXCEL(Thisapp), &dpv);
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
        return hres;
    }
    This->pDoc = V_DISPATCH(&resultDoc);
    IDispatch_AddRef(This->pDoc);

    /*���� ������� pSheets*/
    hres = _I_SheetsConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Sheets_QueryInterface(punk, &IID_I_Sheets, (void**) &(This->pSheets));
/*    I_Sheets_Release(punk);*/
    if (FAILED(hres)) return E_NOINTERFACE;
    /*������ ��������������*/
    hres = MSO_TO_OO_I_Sheets_Initialize((I_Sheets*)(This->pSheets), iface);
    if (FAILED(hres)){
        TRACE("ERROR FAILED Sheets_Initialize \n");
    }
    /*������ ���������� ������� ��������� ����� ������*/
    I_Sheets_get_Count((I_Sheets*)(This->pSheets), &count_list);
    if (count_list>Thisapp->sheetsinnewworkbook) {
        /*����� ������� ����� �� ���������� ���-��*/
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
            /*����� �������� ������ �� ���������� ���-��*/
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


    /*����������� ������ ���������� ��� ������*/
    VariantClear(&param0);
    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&param3);
    VariantClear(&resultDoc);

    TRACE_OUT;
    return hres;
}

HRESULT MSO_TO_OO_I_Sheets_Initialize(
        I_Sheets* iface,
        I_Workbook *wb)
{
    SheetsImpl *This = (SheetsImpl*)iface;
    VARIANT resultSheets;
    TRACE_IN;

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

    TRACE_OUT;
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
    VARIANT res;
    HRESULT hres;
    VariantInit(&resultSheet);
    TRACE_IN;

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

    /*���� ������� WorkSheet*/
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

    TRACE_OUT;
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
    TRACE_IN;

    This->pwb = wb;
    I_Workbook_AddRef(This->pwb);
    This->pOOSheet = oosheet;
    IDispatch_AddRef(This->pOOSheet);


    if (This->pAllRange != NULL) {
        IDispatch_Release(This->pAllRange);
        This->pAllRange = NULL;
    }
    /*���� This->pAllRange = NULL ��� ���� �������*/
    hres = _I_RangeConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &(This->pAllRange));

    if (FAILED(hres)) return E_NOINTERFACE;

    /*����������� ��������� �� worksheet*/
    RangeImpl *this_range = (RangeImpl*) ((I_Range*)This->pAllRange);
    this_range->pwsheet = (IDispatch*)iface;
    this_range->is_release = 0;
//    IDispatch_AddRef(this_range->pwsheet);

    WorkbookImpl *wbtemp = (WorkbookImpl*)This->pwb;
    /*����������� ��������� �� Application*/

    MSO_TO_OO_I_Range_Initialize2((I_Range*)This->pAllRange, This->pOOSheet);

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_GetDispatchPropertyValue(
         _Application *app,
         IDispatch** pIDispatch)
{
    /* there are many of the Open Office functions use "com.sun.star.beans.PropertyValue",
    using this method */
    HRESULT hres;
    VARIANT res;
    VARIANT objstr;
    TRACE_IN;

    _ApplicationImpl *This = (_ApplicationImpl*)app;

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

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_GetDispatchHelper(
         _Application *app,
         IDispatch** pIDispatch)
{
    /* there are many of the Open Office functions use "com.sun.star.frame.DispatchHelper",
    using this method */
    HRESULT hres;
    VARIANT res;
    VARIANT objstr;
    TRACE_IN;

    _ApplicationImpl *This = (_ApplicationImpl*)app;
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

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_ExecuteDispatchHelper_ActiveWorkBook(
        _Application *app,
        BSTR ooCommand,
        VARIANT ooParams) /*������ ���� ������*/
{
    _ApplicationImpl *This = (_ApplicationImpl*)app;
    HRESULT hres;
    VARIANT res;
    TRACE_IN;

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

    hres = _Application_get_ActiveWorkbook(app,(IDispatch**) &wb);
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

    TRACE_OUT;
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
    WorkbooksImpl *This_wbks = (WorkbooksImpl*)(This_wb->pworkbooks);
    _ApplicationImpl *This_app = (_ApplicationImpl*)(This_wbks->pApplication);
    HRESULT hres;
    VARIANT res;
    TRACE_IN;

    if (This_wb==NULL) {
        return E_POINTER;
    }
    if (This_app==NULL) {
        return E_POINTER;
    }

    IDispatch *oodispatcher;
    hres = MSO_TO_OO_GetDispatchHelper((_Application*)This_app, &oodispatcher);
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

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_CloseWorkbook(
         I_Workbook *wb,
         BSTR filename)
{
    WorkbookImpl *This = (WorkbookImpl*)wb;
    WorkbooksImpl *This_wbks = (WorkbooksImpl*)(This->pworkbooks);
    _ApplicationImpl *this_app = (_ApplicationImpl*)(This_wbks->pApplication);
    VARIANT res;
    SAFEARRAY FAR* pPropVals;
    long ix = 0;
    VARIANT p3,p2;
    HRESULT hres;
    V_VT(&p2) = VT_BOOL;
    V_BOOL(&p2) = VARIANT_TRUE;
    TRACE_IN;

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
    MSO_TO_OO_GetDispatchPropertyValue(APPEXCEL(this_app), &dpv);
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

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_I_Workbook_Initialize2(
        I_Workbook* iface,
        I_Workbooks *pwrks,
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
    TRACE_IN;

    This->pworkbooks = (IDispatch*)pwrks;
    if (This->pworkbooks != NULL) I_Workbooks_AddRef(This->pworkbooks);

    WorkbooksImpl *wbks = (WorkbooksImpl*)pwrks;
    _ApplicationImpl *Thisapp = (_ApplicationImpl*)(wbks->pApplication);

    VariantInit(&param0);
    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&param3);
    V_VT(&param0) = VT_BSTR;
    V_BSTR(&param0) = SysAllocString(Filename); /* Name of document */
    /* This->filename = SysAllocString(Filename);  ���������� ��� ����� */
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(L"_blank");  /* Template */
    /* ��� ��������� ����������� ������ ���� �������� ��������� ������� ��������*/


    /*��� ��������� ������������ �� ���������*/
    V_VT(&param2) = VT_I2;
    V_I2(&param2) = 0;  // Another params count
    if (astemplate==VARIANT_FALSE) {
        MSO_TO_OO_GetDispatchPropertyValue(APPEXCEL(Thisapp), &dpv);
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
        /*��������� ������ �� ������*/
        MSO_TO_OO_GetDispatchPropertyValue(APPEXCEL(Thisapp), &dpv);
        if (dpv == NULL)
            return E_FAIL;
        V_VT(&p1) = VT_BSTR;
        V_BSTR(&p1) = SysAllocString(L"AsTemplate");
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Name", 1, p1);
        SysFreeString(V_BSTR(&p1));
        V_VT(&p2) = VT_BOOL; 
        V_BOOL(&p2) = VARIANT_TRUE;
        AutoWrap(DISPATCH_PROPERTYPUT, &res, dpv, L"Value", 1, p2);
        MSO_TO_OO_GetDispatchPropertyValue(APPEXCEL(Thisapp), &dpv2);
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
        return hres;
    }
    This->pDoc = V_DISPATCH(&resultDoc);
    IDispatch_AddRef(This->pDoc);

    /*���� ������� pSheets*/
    hres = _I_SheetsConstructor((LPVOID*) &punk);
    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Sheets_QueryInterface(punk, &IID_I_Sheets, (void**) &(This->pSheets));

    if (FAILED(hres)) return E_NOINTERFACE;
    /*������ ��������������*/
    hres = MSO_TO_OO_I_Sheets_Initialize((I_Sheets*)(This->pSheets), iface);

    VariantClear(&param0);
    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&param3);
    VariantClear(&resultDoc);

    TRACE_OUT;
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
    TRACE_IN;

    if (This_parent->pOORange == NULL) {
        ERR("Object is NULL \n");                          
       return E_POINTER;
    }

    VARIANT resRange;
    VariantInit (&resRange);
    VARIANT vLeft, vRight, vTop, vBottom;
    VariantInit(&vLeft);
    V_VT(&vLeft) = VT_I4;
    V_I4(&vLeft) = topLeft.x - 1; /* �������� 1, �.�. ��������� � ���� */
    VariantInit(&vTop);
    V_VT(&vTop) = VT_I4;
    V_I4(&vTop) = topLeft.y - 1;
    VariantInit(&vRight);
    V_VT(&vRight) = VT_I4;
    V_I4(&vRight) = bottomRight.x - 1;
    VariantInit(&vBottom);
    V_VT(&vBottom) = VT_I4;
    V_I4(&vBottom) = bottomRight.y - 1;

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &resRange, This_parent->pOORange, L"getCellRangeByPosition", 4, vBottom, vRight, vTop, vLeft);
    if (FAILED(hres)) {
       This->pOORange = NULL;
       return hres;
    }
    /*����������� ��������� �� worksheet*/
    This->pwsheet = This_parent->pwsheet;
    IDispatch_AddRef(This->pwsheet);
    /*����������� ��������� �� Application*/

    This->pOORange = V_DISPATCH(&resRange);
    IDispatch_AddRef(V_DISPATCH(&resRange));
    VariantClear(&resRange);

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_I_Range_Initialize2(
        I_Range* iface,
        IDispatch *oosheet)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

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
        TRACE_OUT;
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
    TRACE_IN;

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (oosheet == NULL) {
        TRACE("ERROR oosheet= NULL \n");
        return E_POINTER;
    }    

    if (This->pOORange != NULL) {
        IDispatch_Release(This->pOORange);
        This->pOORange = NULL;
    }
    This->pOORange = oosheet;
    if (This->pOORange != NULL) {
        IDispatch_AddRef(This->pOORange);
    }

    if (This->pwsheet != NULL) {
        IDispatch_Release(This->pwsheet);
        This->pwsheet = NULL;
    }
    This->pwsheet = psheet;
    if (This->pwsheet != NULL) {
        IDispatch_AddRef(This->pwsheet);
    }

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_GetRangeAddress(
        I_Range* iface,
        long *lLeft,
        long *lTop,
        long *lRight,
        long *lBottom)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

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

    TRACE_OUT;
    return hres;
}

HRESULT MSO_TO_OO_GetActiveWorkbook(
        I_Workbooks* iface,
        I_Workbook **wb)
{
    WorkbooksImpl *This = (WorkbooksImpl*)iface;
    TRACE_IN;

    if (This->current_workbook<0) {
        TRACE("ERROR No active Workbook \n");
        *wb = NULL;
        return E_FAIL;
    }

    *wb = (I_Workbook*)This->pworkbook[This->current_workbook];
    I_Workbook_AddRef(*wb);

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_GetActiveCells(
        I_Workbooks* iface,
        I_Range **ppRange)
{
    IDispatch *pWorkbook;
    HRESULT hres;
    TRACE_IN;

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

    /*������� ����� ������ I_Range*/
    IDispatch *pRange;
    IUnknown *punk = NULL;

    hres = _I_RangeConstructor((LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);

    if (pRange == NULL) {
        return E_FAIL;
    }

    RangeImpl *this_range = (RangeImpl*) ((I_Range*)pRange);
    /*����������� ��������� �� parent worksheet*/
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
    /*����������� ��������� �� parent worksheet*/
    hres = MSO_TO_OO_I_Range_Initialize2((I_Range*)pRange,pCurrentCell);

    *ppRange = (I_Range*)pRange;
    I_Range_AddRef(*ppRange);
    I_Range_Release((I_Range*)pRange);

    /*����������� ������*/
    I_Workbook_Release((I_Workbook*)pWorkbook);
    IDispatch_Release(pCurrentCell);
    I_Sheets_Release(pSheets);
    I_Worksheet_Release(pworksheet);

    TRACE_OUT;
    return hres;
}

HRESULT MSO_TO_OO_I_Range_Initialize_ByName(
         I_Range *iface,
         I_Range *pParentRange,
         VARIANT rangename)
{
    RangeImpl *This = (RangeImpl*)iface;
    RangeImpl *This_parent = (RangeImpl*)pParentRange;
    TRACE_IN;

    if (This_parent->pOORange == NULL) {
       return E_POINTER;
    }

    VARIANT resRange;
    VariantInit (&resRange);

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &resRange, This_parent->pOORange, L"getCellRangeByName", 1, rangename);
    if (FAILED(hres)) {
        This->pOORange = NULL;
        return hres;
    }

    /*����������� ��������� �� worksheet*/
    This->pwsheet = This_parent->pwsheet;
    IDispatch_AddRef(This->pwsheet);
    /*����������� ��������� �� worksheet*/

    This->pOORange = V_DISPATCH(&resRange);
    IDispatch_AddRef(This->pOORange);
    VariantClear(&resRange);

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_CorrectArg(
         VARIANT value,
         VARIANT *retval)
{
VariantInit(retval);
if (V_ISBYREF(&value)) {  
                       
    if ((V_VT(&value) - VT_BYREF) & VT_ARRAY) {
       V_VT(retval) = V_VT(&value) - VT_BYREF;
       V_ARRAY(retval) =*(V_ARRAYREF(&value));   
       return S_OK;           
    }  
              
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
/*    case VT_ARRAY:{
        V_VT(retval) = VT_ARRAY;
        V_ARRAY(retval) =*(V_ARRAYREF(&value));
        break;
        }*/
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
    TRACE_IN;

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

    TRACE_OUT;
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
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}


//���������� ������ ������� �� �� �����.
//���� �� ������� ���������� -1
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
    TRACE_IN;

    if (This==NULL) {
        return E_POINTER;
    }

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
                    TRACE_OUT;
                    return i - 1;
                }
                SysFreeString(tmp_name);
            }
            IDispatch_Release(wsh);
            wsh = NULL;
        }
        i++;
    }
    TRACE_OUT;
    return -1;
}

/*���������� Workbook � ������*/
long MSO_TO_OO_GlobalFindIndexWorksheetByName(
        _Application *app,
        BSTR name,
        IDispatch **retval)
{
    _ApplicationImpl *This_app = (_ApplicationImpl*)app;
    int i,id;
    WorkbooksImpl *wbs = (WorkbooksImpl*)This_app->pdWorkbooks;
    SheetsImpl *wsheets;
    WorkbookImpl *wb;
    TRACE_IN;

    for (i=0;i<wbs->count_workbooks;i++){
        if (wbs->pworkbook[i]!=NULL) {
            wb = (WorkbookImpl*)(wbs->pworkbook[i]);
            wsheets = (SheetsImpl*)wb->pSheets;
            id = MSO_TO_OO_FindIndexWorksheetByName((I_Sheets*)wsheets, name);
            if (id>=0) {
               *retval = (IDispatch*)wb;
               TRACE_OUT;
               return id;
            }
        }
    }
    TRACE_OUT;
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
    TRACE_IN;

    VariantInit(&vframe);
    VariantInit(&vRet);
    VariantInit(&param1);

    if (This->pwsheet!=NULL) {
        I_Worksheet_Release((I_Worksheet*)This->pwsheet);
    }
    This->pwsheet = (IDispatch*)wsh;
    I_Worksheet_AddRef((I_Worksheet*)This->pwsheet);

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

    TRACE_OUT;
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
    TRACE_IN;

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

    TRACE_OUT;
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
    TRACE_IN;

    VariantInit(&vRet);

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pWorkbook!=NULL) {
         I_Workbook_Release((I_Workbook*)(This->pWorkbook));
    }
    This->pWorkbook = (IDispatch*)wb;
    if (This->pWorkbook != NULL) I_Workbook_AddRef((I_Workbook*)(This->pWorkbook));

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

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_Workbook_SetVisible(
        I_Workbook *wb,
        VARIANT_BOOL vbvisible)
{
    WorkbookImpl *This = (WorkbookImpl*)wb;
    HRESULT hres;
    VARIANT oocontr, ooframe, oocontwindow, param, res;
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}

#define OUTLINE_THIS(iface) DEFINE_THIS(OutlineImpl, outline, iface)
#define WORKSHEET_THIS(iface) DEFINE_THIS(WorksheetImpl, worksheet, iface)
HRESULT MSO_TO_OO_I_Outline_Initialize(
        I_Outline* iface,
        I_Worksheet *iwsh)
{
    OutlineImpl *This = OUTLINE_THIS(iface);
    WorksheetImpl *This_worksheet = WORKSHEET_THIS(iwsh);
    TRACE_IN;

    if (This->pWorksheet) {
        I_Worksheet_Release(This->pWorksheet);
    }
    This->pWorksheet = iwsh;
    I_Worksheet_AddRef(This->pWorksheet);

    if (This->pOOSheet) {
        IDispatch_Release(This->pOOSheet);
    }
    This->pOOSheet = This_worksheet->pOOSheet;
    IDispatch_AddRef(This->pOOSheet);

    TRACE_OUT;
    return S_OK;
}
#undef WORKSHEET_THIS
#undef OUTLINE_THIS

HRESULT MSO_TO_OO_Name_Initialize_By_Name(
        Name* iface,
        Names *pnames,
        VARIANT varname)
{
    NameImpl *This = (NameImpl*)iface;
    NamesImpl *onames = (NamesImpl*)pnames;
    WorkbookImpl *wbi = (WorkbookImpl*)(onames->pWorkbook);
    VARIANT vRet,vRet2;
    HRESULT hres;
    TRACE_IN;

    VariantInit(&vRet);

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pNames!=NULL) {
         Names_Release((Names*)(This->pNames));
    }
    This->pNames = (IDispatch*)pnames;
    if (This->pNames != NULL) Names_AddRef((Names*)(This->pNames));

    if (wbi->pDoc==NULL) {
        TRACE("Object pDoc is NULL\n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRet, wbi->pDoc, L"NamedRanges",0);
    if (FAILED(hres)) {
        TRACE("ERROR when NamedRange \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vRet2, V_DISPATCH(&vRet), L"getByName",1,varname);
    if (FAILED(hres)) {
        TRACE("ERROR when NamedRange \n");
        return E_FAIL;
    }

    if (This->pOOName!=NULL) {
         IDispatch_Release(This->pOOName);
    }
    This->pOOName = V_DISPATCH(&vRet2);
    if (This->pOOName != NULL) IDispatch_AddRef(This->pOOName);

    VariantClear(&vRet);

    TRACE_OUT;
    return S_OK;
}

HRESULT MSO_TO_OO_Name_Initialize_By_Index(
        Name* iface,
        Names *pnames,
        VARIANT varindex)
{
    NameImpl *This = (NameImpl*)iface;
    NamesImpl *onames = (NamesImpl*)pnames;
    WorkbookImpl *wbi = (WorkbookImpl*)(onames->pWorkbook);
    VARIANT vRet,vRet2;
    HRESULT hres;
    TRACE_IN;

    VariantInit(&vRet);

    if (This == NULL) {
        TRACE("ERROR THIS = NULL \n");
        return E_POINTER;
    }

    if (This->pNames!=NULL) {
         Names_Release((Names*)(This->pNames));
    }
    This->pNames = (IDispatch*)pnames;
    if (This->pNames != NULL) Names_AddRef((Names*)(This->pNames));

    if (wbi->pDoc==NULL) {
        TRACE("Object pDoc is NULL\n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRet, wbi->pDoc, L"NamedRanges",0);
    if (FAILED(hres)) {
        TRACE("ERROR when NamedRange \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vRet2, V_DISPATCH(&vRet), L"getByIndex",1,varindex);
    if (FAILED(hres)) {
        TRACE("ERROR when NamedRange \n");
        return E_FAIL;
    }

    if (This->pOOName!=NULL) {
         IDispatch_Release(This->pOOName);
    }
    This->pOOName = V_DISPATCH(&vRet2);
    if (This->pOOName != NULL) IDispatch_AddRef(This->pOOName);

    VariantClear(&vRet);

    TRACE_OUT;
    return S_OK;
}

BOOL    Is_Variant_Null(
        VARIANT var)
{
if ((V_VT(&var)==VT_EMPTY) || (V_VT(&var)==VT_NULL) || (V_VT(&var)==VT_ERROR)) return 1;
return 0;
}
