/*
 * IWorksheet interface functions
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


static WCHAR const str_name[] = {
    'N','a','m','e',0};
static WCHAR const str_cells[] = {
    'C','e','l','l','s',0};
static WCHAR const str_range[] = {
    'R','a','n','g','e',0};
static WCHAR const str_paste[] = {
    'P','a','s','t','e',0};
static WCHAR const str_activate[] = {
    'A','c','t','i','v','a','t','e',0};
static WCHAR const str_rows[] = {
    'R','o','w','s',0};
static WCHAR const str_columns[] = {
    'C','o','l','u','m','n','s',0};
static WCHAR const str_copy[] = {
    'C','o','p','y',0};
static WCHAR const str_delete[] = {
    'D','e','l','e','t','e',0};
static WCHAR const str_pagesetup[] = {
    'P','a','g','e','S','e','t','u','p',0};

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Worksheet_AddRef(
        I_Worksheet* iface)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    ULONG ref;

    TRACE("REF=%i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_QueryInterface(
        I_Worksheet* iface,
        REFIID riid,
        void **ppvObject)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Worksheet)) {
        *ppvObject = &This->_worksheetVtbl;
        MSO_TO_OO_I_Worksheet_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Worksheet_Release(
        I_Worksheet* iface)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    ULONG ref;

    TRACE("REF=%i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pOOSheet != NULL) {
            IDispatch_Release(This->pOOSheet);
            This->pOOSheet = NULL;
        }
        if (This->pwb != NULL) {
            I_Workbook_Release(This->pwb);
            This->pwb = NULL;
        }
        if (This->pAllRange != NULL) {
            IDispatch_Release(This->pAllRange);
            This->pAllRange = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Worksheet methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Name(
        I_Worksheet* iface,
        BSTR *pbstrName)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    VARIANT res;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheet, L"getName", 0);

    if (hres == S_OK)
       *pbstrName = V_BSTR(&res);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_Name(
        I_Worksheet* iface,
        BSTR bstrName)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    VARIANT res;
    VARIANT new_str;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    VariantInit(&new_str);

    V_VT(&new_str) = VT_BSTR;
    V_BSTR(&new_str) = bstrName;

    hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheet, L"setName", 1, new_str);

    VariantClear(&res);
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Cells(
        I_Worksheet* iface,
        IDispatch **ppRange)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;

    TRACE("\n");

    *ppRange = This->pAllRange;
    IDispatch_AddRef(This->pAllRange);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Range(
        I_Worksheet* iface,
        VARIANT Cell1,
        VARIANT Cell2,
        IDispatch **ppRange)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    IDispatch *pRange;
    I_Range *pCell1;
    I_Range *pCell2;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    HRESULT hres;
    VARIANT vNull;
    VariantInit(&vNull);

    TRACE("\n");

    if ((V_VT(&Cell2)==VT_NULL)||(V_VT(&Cell2)==VT_EMPTY)) {
        hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &pRange);
        if (FAILED(hres)) {
            TRACE("ERROR when get_Cells\n");
            return hres;
        }
        hres = I_Range_get__Default((I_Range*)pRange,Cell1,vNull,ppRange);
        if (FAILED(hres)) {
            TRACE("ERROR when get__Default\n");
            IDispatch_Release(pRange);
            return hres;
        }
        IDispatch_Release(pRange);
        return S_OK;
    }

    if ((V_VT(&Cell1)==VT_BSTR)&&(V_VT(&Cell2)==VT_BSTR)) {
            /*Два параметра и оба строковые переменные*/
            WTRACE(L"2 Parametra BSTR %s   %s \n", V_BSTR(&Cell1), V_BSTR(&Cell2));

            hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &pRange);
            if (FAILED(hres)) {
                TRACE("ERROR when get_Cells\n");
                return hres;
            }

            hres = I_Range_get__Default((I_Range*)pRange,Cell1,vNull, (IDispatch**)&pCell1);
            if (FAILED(hres)) {
                TRACE("ERROR when get__Default\n");
                IDispatch_Release(pRange);
                return hres;
            }
            hres = I_Range_get__Default((I_Range*)pRange,Cell2,vNull, (IDispatch**)&pCell2);
            if (FAILED(hres)) {
                TRACE("ERROR when get__Default\n");
                IDispatch_Release(pRange);
                return hres;
            }
        IDispatch_Release(pRange);
        } else {
            pCell1 = (I_Range*) V_DISPATCH(&Cell1);
            pCell2 = (I_Range*) V_DISPATCH(&Cell2);
        }

    if ((pCell1 == NULL) || (pCell2 == NULL)) {
        TRACE("Error - one of the pointers is NULL \n");
        return E_INVALIDARG;
    }

    long lCell1L, lCell1R, lCell1T, lCell1B;
    long lCell2L, lCell2R, lCell2T, lCell2B;
    /*long lRangeL, lRangeR, lRangeT, lRangeB;*/

    hres = MSO_TO_OO_GetRangeAddress(pCell1, &lCell1L, &lCell1T, &lCell1R, &lCell1B);
    if (hres != S_OK) {
        return hres;
    }

    hres = MSO_TO_OO_GetRangeAddress(pCell2, &lCell2L, &lCell2T, &lCell2R, &lCell2B);
    if (hres != S_OK) {
        return hres;
    }

    /*Создаем новый объект I_Range*/
    hres = _I_RangeConstructor(pUnkOuter, (LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);
    if (pRange == NULL) {
       return E_FAIL;
    }
    struct CELL_COORD cell1,cell2;
    cell1.x = lCell1T + 1;
    cell1.y = lCell1L + 1;
    cell2.x = lCell2B + 1;
    cell2.y = lCell2R + 1;
    
    TRACE(" cell1.x=%i \n    cell1.y=%i \n    cell2.x=%i \n    cell2.y=%i \n", cell1.x, cell1.y, cell2.x, cell2.y);
    
    hres = MSO_TO_OO_I_Range_Initialize(pRange,This->pAllRange, cell1, cell2);
    if (hres != S_OK) {
        IDispatch_Release(pRange);
        return hres;
    }

    *ppRange = pRange;
    I_Range_AddRef((I_Range*)*ppRange);
    I_Range_Release(pRange);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Paste(
        I_Worksheet* iface,
        VARIANT Destination,
        VARIANT Link,
        long lcid)
{
    TRACE("NOT REALISED\n");
/*
    WRITE_LOG_R_W((L"CWorksheet::Paste(...)"));
    if ((CApplication::s_pdOOApp == NULL) || (Destination.vt != VT_DISPATCH))
    {
        return E_FULT hres;
    ::OleInitialize(NULL);
    IDataObject *pDataObj;
    hres = ::OleGetClipboard(&pDataObj);
    if (hres != S_OK)
    {
        return hres;
    }
        FORMATETC               fmtetc;
        STGMEDIUM               stgmed = {0};
    fmtetc.cfFormat = CF_TEXT;
        fmtetc.dwAspect = DVASPECT_CONTENT;
        fmtetc.lindex   = -1;
        fmtetc.ptd              = 0;
        fmtetc.tymed    = TYMED_HGLOBAL;

    hres = pDataObj->GetData(&fmtetc, &stgmed);
    pDataObj->Release();
    if (hres != S_OK)
    {
        pDataObj->Release();
        return hres;
    }

    char* szPasteString = (char*) GlobalLock(stgmed.hGlobal);
    GlobalUnlock(stgmed.hGlobal);
    if (szPasteString == NULL)
    {
        return E_FAIL;
    }

    //VARIANT vRes;

    CRange *pRange = (CRange*) Destination.pdispVal;
    CComVariant vStr = szPasteString;
    CComVariant vAnyValue;
    hres = pRange->put_Value(vAnyValue, vStr);
    return hres;
*/

    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Activate(
        I_Worksheet* iface)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    BSTR command;
    VARIANT param, res;
    long index;
    SAFEARRAY FAR* pPropVals;
    long ix = 0;
    WorkbookImpl *paren_wb = (WorkbookImpl*)This->pwb;
    HRESULT hres;
    BSTR name;

    TRACE("\n");

    if (This==NULL) {
        return E_POINTER;
    }
    hres = I_Worksheet_get_Name(iface, &name);
    if (FAILED(hres)) {
       TRACE("ERROR when get_Name\n");
       /*просто выходим из процедуры*/ 
       return S_OK;
    }
    index = MSO_TO_OO_FindIndexWorksheetByName((I_Sheets*)(paren_wb->pSheets), name);
    if (index==-1) {
       TRACE("ERROR not find such name\n");
       /*Если не нашли, то просто выходим из процедуры*/ 
       return S_OK;
    }
    command = SysAllocString(L".uno:JumpToTable");
    /* Create PropertyValue with save-format-data */
    IDispatch *ooParams;
    MSO_TO_OO_GetDispatchPropertyValue((I_ApplicationExcel*)(paren_wb->pApplication), &ooParams);
    if (ooParams == NULL)
        return E_FAIL;

    VARIANT p1;
    V_VT(&p1) = VT_BSTR;
    V_BSTR(&p1) = SysAllocString(L"Nr");
    AutoWrap(DISPATCH_PROPERTYPUT, &res, ooParams, L"Name", 1, p1);
    VariantClear(&p1);
    V_VT(&p1) = VT_I4;
    V_I4(&p1) = index+1;
    AutoWrap(DISPATCH_PROPERTYPUT, &res, ooParams, L"Value", 1, p1);
    /* Init params */
    pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 1 );
    hres = SafeArrayPutElement( pPropVals, &ix, ooParams );
    if (FAILED(hres)){
        return hres;
    }

    VariantInit (&param);
    V_VT(&param) = VT_DISPATCH | VT_ARRAY;
    V_ARRAY(&param) = pPropVals;

    hres = MSO_TO_OO_ExecuteDispatchHelper_WB((I_Workbook*)paren_wb, command, param);
    if (FAILED(hres)){
        return hres;
    }

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Rows(
        I_Worksheet* iface,
        VARIANT Row,
        IDispatch **ppRange)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    IDispatch *pRange;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    HRESULT hres;
    IDispatch *tmp_range;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if ((V_VT(&Row)==VT_NULL)||(V_VT(&Row)==VT_EMPTY)) {
        TRACE("PARAMETER IS NULL\n",V_I4(&Row));
        /*Без параметра*/
        /*Возвращаем всю таблицу*/
        return MSO_TO_OO_I_Worksheet_get_Cells(iface, ppRange);
    } else {
        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&Row, &Row, 0, 0, VT_I4);

        if (V_VT(&Row) == VT_I4) {
            /*параметр это индекс*/
            TRACE("PARAMETER IS %i\n",V_I4(&Row));
            hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &tmp_range);
            if (FAILED(hres)) {
                TRACE("ERROR get_Cells\n");
                return hres;
            }
            /*Создаем новый объект I_Range*/
            hres = _I_RangeConstructor(pUnkOuter, (LPVOID*) &punk);

            if (FAILED(hres)) return E_NOINTERFACE;

            hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);
            if (pRange == NULL) {
                return E_FAIL;
            }
            struct CELL_COORD cell1,cell2;
            cell1.x = 1;
            cell1.y = V_I4(&Row);
            cell2.x = 256;
            cell2.y = V_I4(&Row);;
            hres = MSO_TO_OO_I_Range_Initialize(pRange,tmp_range, cell1, cell2);
            if (hres != S_OK) {
                IDispatch_Release(pRange);
                I_Range_Release(tmp_range);
                return hres;
            }

            *ppRange = pRange;
            I_Range_AddRef((I_Range*)*ppRange);
            I_Range_Release(pRange);
            I_Range_Release(tmp_range);
            return S_OK;
        } else 
            if (V_VT(&Row) == VT_BSTR) {
                /*параметр это строка*/
                WTRACE(L"PARAMETER IS %s\n",V_BSTR(&Row));
                WCHAR *tmp_str;
                int i, row1, row2, itmp;
                tmp_str = V_BSTR(&Row);
                i=0;
                itmp=0;
                while (tmp_str[i]!=0) {
                    if ((tmp_str[i]>=L'0')&&(tmp_str[i]<=L'9')) {
                        itmp=itmp*10+(tmp_str[i]-L'0');
                    }
                    if (tmp_str[i]==L':') {
                        row1 = itmp;
                        itmp = 0;
                    }
                    i++;
                }
                row2 = itmp;

                hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &tmp_range);
                if (FAILED(hres)) {
                    TRACE("ERROR get_Cells\n");
                    return hres;
                }

                /*Создаем новый объект I_Range*/
                hres = _I_RangeConstructor(pUnkOuter, (LPVOID*) &punk);

                if (FAILED(hres)) return E_NOINTERFACE;

                hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);

                if (pRange == NULL) {
                    return E_FAIL;
                }
                struct CELL_COORD cell1,cell2;
                cell1.x = 1;
                cell1.y = row1;
                cell2.x = 256;
                cell2.y = row2;
TRACE("PARAMETRS IS %i     %i \n", row1,row2);
                hres = MSO_TO_OO_I_Range_Initialize(pRange,tmp_range, cell1, cell2);
                if (FAILED(hres)) {
                    TRACE("ERROR Range_Initialize\n");
                    IDispatch_Release(pRange);
                    IDispatch_Release(tmp_range);
                    return hres;
                }

                *ppRange = pRange;
                I_Range_AddRef((I_Range*)*ppRange);
                I_Range_Release(pRange);
                I_Range_Release(tmp_range);

                return S_OK;
            } else {
                *ppRange = NULL;
                return E_FAIL;
            }
    }

    return E_FAIL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Columns(
        I_Worksheet* iface,
        VARIANT Column,
        IDispatch **ppRange)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    IDispatch *pRange;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    HRESULT hres;
    IDispatch *tmp_range;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if ((V_VT(&Column)==VT_NULL)||(V_VT(&Column)==VT_EMPTY)) {
        TRACE("PARAMETER IS NULL\n",V_I4(&Column));
        /*Без параметра*/
        /*Возвращаем всю таблицу*/
        return MSO_TO_OO_I_Worksheet_get_Cells(iface, ppRange);
    } else {
        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&Column, &Column, 0, 0, VT_I4);

        if (V_VT(&Column) == VT_I4) {
            /*параметр это индекс*/
            TRACE("PARAMETER IS %i\n",V_I4(&Column));
            hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &tmp_range);
            if (FAILED(hres)) {
                TRACE("ERROR get_Cells\n");
                return hres;
            }
            /*Создаем новый объект I_Range*/
            hres = _I_RangeConstructor(pUnkOuter, (LPVOID*) &punk);

            if (FAILED(hres)) return E_NOINTERFACE;

            hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);

            if (pRange == NULL) {
                return E_FAIL;
            }
            struct CELL_COORD cell1,cell2;
            cell1.x = V_I4(&Column);
            cell1.y = 1;
            cell2.x = V_I4(&Column);
            cell2.y = 65536;
            hres = MSO_TO_OO_I_Range_Initialize(pRange,tmp_range, cell1, cell2);
            if (FAILED(hres)) {
                IDispatch_Release(pRange);
                IDispatch_Release(tmp_range);
                return hres;
            }

            *ppRange = pRange;
            I_Range_AddRef((I_Range*)*ppRange);
            IDispatch_Release(pRange);
            IDispatch_Release(tmp_range);

            return S_OK;
        } else 
            if (V_VT(&Column) == VT_BSTR) {
                /*параметр это строка*/
                /*Разбираем строки вида `col1:col2`*/
                WTRACE(L"PARAMETER IS STRING %s\n",V_BSTR(&Column));
                WCHAR *tmp_str;
                int i, col1, col2, itmp;
                tmp_str = V_BSTR(&Column);
                i=0;
                itmp=0;
                while (tmp_str[i]!=0) {
                    if ((tmp_str[i]>=L'0')&&(tmp_str[i]<=L'9')) {
                        itmp=itmp*10+(tmp_str[i]-L'0');
                    }
                    if ((tmp_str[i]>=L'A')&&(tmp_str[i]<=L'Z')) {
                        itmp=itmp*26+(tmp_str[i]-L'A')+1;
                    }
                    if (tmp_str[i]==L':') {
                        col1 = itmp;
                        itmp = 0;
                    }
                    i++;
                }
                col2 = itmp;

                hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &tmp_range);
                if (FAILED(hres)) {
                    TRACE("ERROR get_Cells\n");
                    return hres;
                }

                /*Создаем новый объект I_Range*/
                hres = _I_RangeConstructor(pUnkOuter, (LPVOID*) &punk);

                if (FAILED(hres)) return E_NOINTERFACE;

                hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);

                if (pRange == NULL) {
                    return E_FAIL;
                }
                struct CELL_COORD cell1,cell2;
                cell1.x = col1;
                cell1.y = 1;
                cell2.x = col2;
                cell2.y = 65536;
TRACE("PARAMETRS IS %i     %i \n", col1,col2);
                hres = MSO_TO_OO_I_Range_Initialize(pRange,tmp_range, cell1, cell2);
                if (FAILED(hres)) {
                    TRACE("ERROR Range_Initialize\n");
                    IDispatch_Release(pRange);
                    IDispatch_Release(tmp_range);
                    return hres;
                }

                *ppRange = pRange;
                I_Range_AddRef((I_Range*)pRange);
                IDispatch_Release(pRange);
                IDispatch_Release(tmp_range);

                return S_OK;
            } else {
                *ppRange = NULL;
                return E_FAIL;
            }
    }

    return E_FAIL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Copy(
        I_Worksheet* iface,
        VARIANT Before,
        VARIANT After)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    int ftype_add = 0,i;
    HRESULT hres;
    VARIANT vNull, find_name;
    BSTR name,name_of_copy,tmp_name;
    WorksheetImpl *tmp_wsh;
    WorkbookImpl *parent_wb = (WorkbookImpl*)This->pwb;
    SheetsImpl *sheets_find;
    IDispatch *wb_find; 
    VariantInit(&vNull);
    VariantClear(&find_name);
    IDispatch *new_wsh;
    _ApplicationExcelImpl *app = (_ApplicationExcelImpl*)parent_wb->pApplication;
    IDispatch *range1,*range2, *range3;
    VARIANT cols,torange;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    /*Приводим все значения к необходимому виду.*/
    if ((V_VT(&Before)==VT_EMPTY) || (V_VT(&Before)==VT_NULL)) {
        VariantClear(&Before);
    } else {
        tmp_wsh =(WorksheetImpl*)V_DISPATCH(&Before);
        ftype_add = 1;
    }
    if ((V_VT(&After)==VT_EMPTY) || (V_VT(&After)==VT_NULL)) {
        VariantClear(&After);
    } else {
        tmp_wsh =(WorksheetImpl*)V_DISPATCH(&After);
        ftype_add = 2;
    }

hres = I_Worksheet_get_Name((I_Worksheet*)tmp_wsh, &tmp_name);
if (FAILED(hres)) {
   TRACE("ERROR when get_Name\n");
   /*просто выходим из процедуры*/ 
   return S_OK;
}

i = MSO_TO_OO_GlobalFindIndexWorksheetByName((I_ApplicationExcel*)app, tmp_name, &wb_find);
if (i<0) {
    TRACE("Target not find \n");
    return E_FAIL;
} else {
    TRACE(" INDEX = %i \n", i);
    sheets_find = (SheetsImpl*)(((WorkbookImpl*)wb_find)->pSheets);
}

switch (ftype_add) {
case 1:
    WTRACE(L"PAR-------> BEFORE %s\n",V_BSTR(&Before));
    VariantClear(&find_name);
    V_VT(&find_name) = VT_BSTR;
    V_BSTR(&find_name) = SysAllocString(tmp_name);
    hres = I_Sheets_Add((I_Sheets*)sheets_find, find_name, vNull, vNull, vNull, &new_wsh);
    break;
case 2:
    WTRACE(L"PAR-------> AFTER %s \n",V_BSTR(&After));
    VariantClear(&find_name);
    V_VT(&find_name) = VT_BSTR;
    V_BSTR(&find_name) = SysAllocString(tmp_name);
    hres = I_Sheets_Add((I_Sheets*)sheets_find, vNull, find_name, vNull, vNull, &new_wsh);
    break;
default:
    TRACE("to the new Workbook \n");
    return E_NOTIMPL;
}
/*Теперь просто копируем все ячеки из одного Worksheet в другой*/
VariantInit(&cols);
VariantInit(&torange);
V_VT(&cols) = VT_BSTR;
V_BSTR(&cols) = SysAllocString(L"1:256");
I_Worksheet_get_Columns(iface,cols,&range1);
I_Worksheet_get_Columns((I_Worksheet*)new_wsh,cols,&range2);
V_VT(&torange) = VT_DISPATCH;
V_DISPATCH(&torange) = range2;
I_Range_Copy((I_Range*)range1, torange, &range3);
VariantClear(&cols);
/*Необходимо еще скопировать PAGESETUP*/
IDispatch *src,*trg;
double dtmp;
long ltmp;
VARIANT vtmp;
VariantInit(&vtmp);
VARIANT_BOOL vbtmp;
I_Worksheet_get_PageSetup(iface,&src);
I_Worksheet_get_PageSetup((I_Worksheet*)new_wsh,&trg);

I_PageSetup_get_LeftMargin((I_PageSetup*)src, &dtmp);
I_PageSetup_put_LeftMargin((I_PageSetup*)trg, dtmp);
I_PageSetup_get_RightMargin((I_PageSetup*)src, &dtmp);
I_PageSetup_put_RightMargin((I_PageSetup*)trg, dtmp);
I_PageSetup_get_TopMargin((I_PageSetup*)src, &dtmp);
I_PageSetup_put_TopMargin((I_PageSetup*)trg, dtmp);
I_PageSetup_get_BottomMargin((I_PageSetup*)src, &dtmp);
I_PageSetup_put_BottomMargin((I_PageSetup*)trg, dtmp);
I_PageSetup_get_HeaderMargin((I_PageSetup*)src, &dtmp);
I_PageSetup_put_HeaderMargin((I_PageSetup*)trg, dtmp);
I_PageSetup_get_FooterMargin((I_PageSetup*)src, &dtmp);
I_PageSetup_put_FooterMargin((I_PageSetup*)trg, dtmp);
I_PageSetup_get_Orientation((I_PageSetup*)src, &ltmp);
I_PageSetup_put_Orientation((I_PageSetup*)trg, ltmp);
I_PageSetup_get_Zoom((I_PageSetup*)src, &vtmp);
I_PageSetup_put_Zoom((I_PageSetup*)trg, vtmp);
I_PageSetup_get_FitToPagesTall((I_PageSetup*)src, &vtmp);
I_PageSetup_put_FitToPagesTall((I_PageSetup*)trg, vtmp);
I_PageSetup_get_FitToPagesWide((I_PageSetup*)src, &vtmp);
I_PageSetup_put_FitToPagesWide((I_PageSetup*)trg, vtmp);
I_PageSetup_get_CenterHorizontally((I_PageSetup*)src, &vbtmp);
I_PageSetup_put_CenterHorizontally((I_PageSetup*)trg, vbtmp);
I_PageSetup_get_CenterVertically((I_PageSetup*)src, &vbtmp);
I_PageSetup_put_CenterVertically((I_PageSetup*)trg, vbtmp);
VariantClear(&vtmp);
IDispatch_Release(src);
IDispatch_Release(trg);
/*Закончили копировать PAGESETUP*/

/*Переименовываем*/
hres = I_Worksheet_get_Name(iface, &name);
if (FAILED(hres)) {
    TRACE("ERROR when get_Name\n");
    /*просто выходим из процедуры*/ 
    return S_OK;
}
VarBstrCat(name, L" 2",&name_of_copy);

hres = I_Worksheet_put_Name((I_Worksheet*)new_wsh, name_of_copy);
if (FAILED(hres)) {
    TRACE("ERROR when get_Name\n");
    /*просто выходим из процедуры*/ 
    return S_OK;
}

return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Delete(
        I_Worksheet* iface)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    BSTR name;
    HRESULT hres;
    VARIANT param1,res;
    WorkbookImpl *paren_wb = (WorkbookImpl*)This->pwb;
    SheetsImpl *sheets = (SheetsImpl*)((I_Sheets*)(paren_wb->pSheets));

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    hres = MSO_TO_OO_I_Worksheet_get_Name(iface, &name);
    if (FAILED(hres)) {
       TRACE("ERROR when get_Name\n");
       return hres;
    }

    VariantInit(&param1);
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(name);

    hres = AutoWrap(DISPATCH_METHOD, &res, sheets->pOOSheets, L"removeByName", 1,param1);
    if (FAILED(hres)) {
        TRACE("ERROR when removeByName \n");
        return hres;
    }

    SysFreeString(name);
    VariantClear(&param1);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_PageSetup(
        I_Worksheet* iface,
        IDispatch **ppValue)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    HRESULT hr;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    IDispatch *pPageSetup;


    TRACE("(GET) \n");

    if (This==NULL) {
        TRACE("(GET) ERROR Object is NULL \n");
        return E_POINTER;
    }

    *ppValue = NULL;

    hr = _I_PageSetupConstructor(pUnkOuter, (LPVOID*) &punk);

    if (FAILED(hr)) return E_NOINTERFACE;

    hr = I_PageSetup_QueryInterface(punk, &IID_I_PageSetup, (void**) &pPageSetup);
    if (pPageSetup == NULL) {
        return E_FAIL;
    }

    hr = MSO_TO_OO_I_PageSetup_Initialize((I_PageSetup*)pPageSetup, iface);

    if (FAILED(hr)) {
        IDispatch_Release(pPageSetup);
        return hr;
    }

    *ppValue = pPageSetup;

    return S_OK;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Worksheet_GetTypeInfoCount(
        I_Worksheet* iface,
        UINT *pctinfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_GetTypeInfo(
        I_Worksheet* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_GetIDsOfNames(
        I_Worksheet* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    if (!lstrcmpiW(*rgszNames, str_name)) {
        *rgDispId = 1;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_cells)) {
        *rgDispId = 2;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_range)) {
        *rgDispId = 3;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_paste)) {
        *rgDispId = 4;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_activate)) {
        *rgDispId = 5;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_rows)) {
        *rgDispId = 6;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_columns)) {
        *rgDispId = 7;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_copy)) {
        *rgDispId = 8;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_delete)) {
        *rgDispId = 9;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_pagesetup)) {
        *rgDispId = 10;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L" %s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Invoke(
        I_Worksheet* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    HRESULT hres;
    BSTR res;
    IDispatch *dret;
    IDispatch *pretdisp;
    VARIANT vNull;
    VARIANT cell1,cell2,tmpval;

    VariantInit(&vNull);
    VariantInit(&cell1);
    VariantInit(&cell2);
    VariantInit(&tmpval);

    TRACE("dispIdMember = %i\n",dispIdMember);

    if (This == NULL) {
        TRACE("ERROR E_POINTER \n");
        return E_POINTER;
    }

    switch (dispIdMember)
    {
    case 1:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            hres = MSO_TO_OO_I_Worksheet_put_Name(iface, V_BSTR(&(pDispParams->rgvarg[0])));
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pDispParams->cArgs!=0) return E_FAIL;

            hres = MSO_TO_OO_I_Worksheet_get_Name(iface, &res);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BSTR;
                V_BSTR(pVarResult) = res;
            }
            return S_OK;
        }
    case 2:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            switch (pDispParams->cArgs) {
            case 3:
                hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &dret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    TRACE("(case 2) ERROR get_cells hres = %08x\n",hres);
                    return hres;
                }
                /*необходимо привести к значению , т.к. иногда присылаются ссылки*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[2], &cell1);
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &cell2);
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &tmpval);
                I_Range_get__Default((I_Range*)dret,cell1, cell2,&pretdisp);
                I_Range_put_Value((I_Range*)pretdisp, vNull, tmpval);
                IDispatch_Release(dret);
                IDispatch_Release(pretdisp);
                return S_OK;
            }
            TRACE("(case 2) (PUT) only realized with 3 parameters \n");
            return E_NOTIMPL;
        } else {

            hres = MSO_TO_OO_I_Worksheet_get_Cells(iface, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                TRACE("(case 2) ERROR get_cells hres = %08x\n",hres);
                return hres;
            }

            /*здесь надо проверить параметры если они есть, то вызвать метод Range->_Default.*/
            switch(pDispParams->cArgs) {
            case 0:
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=dret;
                }
                break;
            case 1:
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2);
                I_Range_get__Default((I_Range*)dret,cell2, vNull, &pretdisp);
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=pretdisp;
                }
                I_Range_Release((I_Range*)dret);
               break;
            case 2:
                /*необходимо привести к значению , т.к. иногда присылаются ссылки*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &cell1);
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2);

                I_Range_get__Default((I_Range*)dret,cell1, cell2,&pretdisp);
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=pretdisp;
                }
                I_Range_Release((I_Range*)dret);
               break;
            }
            return S_OK;
        }
    case 3:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 1:
                TRACE("(case 3) ONE PARAMETR IS SEND\n");
                return E_NOTIMPL;
            case 2:
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &cell1);
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2);
            hres = MSO_TO_OO_I_Worksheet_get_Range(iface, cell1, cell2, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_DISPATCH;
                V_DISPATCH(pVarResult) = dret;
            }
            return S_OK;
            default:
                TRACE("(case 3) ERROR PARAMETR IS SEND\n");
                return E_FAIL;
            }
        }
    case 4:
        /*method paste   MSO_TO_OO_I_Worksheet_Paste*/

       return E_NOTIMPL;
    case 5:
       return MSO_TO_OO_I_Worksheet_Activate(iface);
    case 6://Rows
       if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 0:
                TRACE("(case 6) 0 Parameter\n");
                hres = MSO_TO_OO_I_Worksheet_get_Rows(iface, vNull, &dret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = dret;
                }
                return S_OK;
            case 1:
                TRACE("(case 6) 1 Parameter\n");
                /*Привести параметры к типу VARIANT если они переданы по ссылке*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1);
                hres = MSO_TO_OO_I_Worksheet_get_Rows(iface, cell1, &dret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = dret;
                }
                return S_OK;
            }
        }
    case 7://Columns
       if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 0:
                TRACE("(case 7) 0 Parameter\n");
                hres = MSO_TO_OO_I_Worksheet_get_Columns(iface, vNull, &dret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = dret;
                }
                return S_OK;
            case 1:
                TRACE("(case 7) 1 Parameter\n");
                /*Привести параметры к типу VARIANT если они переданы по ссылке*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1);
                hres = MSO_TO_OO_I_Worksheet_get_Columns(iface, cell1, &dret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = dret;
                }
                return S_OK;
            }
        }
    case 8:
        switch (pDispParams->cArgs) {
        case 0:
            TRACE("(case 8) 0 Parameter\n");
            hres = MSO_TO_OO_I_Worksheet_Copy(iface, vNull, vNull);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        case 1:
            TRACE("(case 8) 1 Parameter\n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2))) return E_FAIL;

            hres = MSO_TO_OO_I_Worksheet_Copy(iface, cell2, vNull);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        case 2:
            TRACE("(case 8) 2 Parameter\n");
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &cell1))) return E_FAIL;
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2))) return E_FAIL;
            hres = MSO_TO_OO_I_Worksheet_Copy(iface, cell1, cell2);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        }
    case 9:
        return MSO_TO_OO_I_Worksheet_Delete(iface);
    case 10:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            if (pDispParams->cArgs!=0) return E_FAIL;

            hres = MSO_TO_OO_I_Worksheet_get_PageSetup(iface, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_DISPATCH;
                V_DISPATCH(pVarResult) = dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    }

    return E_NOTIMPL;
}

const I_WorksheetVtbl MSO_TO_OO_I_WorksheetVtbl =
{
    MSO_TO_OO_I_Worksheet_QueryInterface,
    MSO_TO_OO_I_Worksheet_AddRef,
    MSO_TO_OO_I_Worksheet_Release,
    MSO_TO_OO_I_Worksheet_GetTypeInfoCount,
    MSO_TO_OO_I_Worksheet_GetTypeInfo,
    MSO_TO_OO_I_Worksheet_GetIDsOfNames,
    MSO_TO_OO_I_Worksheet_Invoke,
    MSO_TO_OO_I_Worksheet_get_Name,
    MSO_TO_OO_I_Worksheet_put_Name,
    MSO_TO_OO_I_Worksheet_get_Cells,
    MSO_TO_OO_I_Worksheet_get_Range,
    MSO_TO_OO_I_Worksheet_Paste,
    MSO_TO_OO_I_Worksheet_Activate,
    MSO_TO_OO_I_Worksheet_get_Rows,
    MSO_TO_OO_I_Worksheet_get_Columns,
    MSO_TO_OO_I_Worksheet_Copy,
    MSO_TO_OO_I_Worksheet_Delete,
    MSO_TO_OO_I_Worksheet_get_PageSetup
};

WorksheetImpl MSO_TO_OO_Worksheet =
{
    &MSO_TO_OO_I_WorksheetVtbl,
    0,
    NULL,
    NULL,
    NULL
};


extern HRESULT _I_WorksheetConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    WorksheetImpl *worksheet;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

    worksheet = HeapAlloc(GetProcessHeap(), 0, sizeof(*worksheet));
    if (!worksheet)
    {
        return E_OUTOFMEMORY;
    }

    worksheet->_worksheetVtbl = &MSO_TO_OO_I_WorksheetVtbl;
    worksheet->ref = 0;
    worksheet->pOOSheet = NULL;
    worksheet->pwb = NULL;
    worksheet->pAllRange = NULL;

    *ppObj = &worksheet->_worksheetVtbl;

    return S_OK;
}
