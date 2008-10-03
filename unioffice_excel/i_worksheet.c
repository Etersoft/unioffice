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


ITypeInfo *ti_worksheet = NULL;

HRESULT get_typeinfo_worksheet(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_worksheet) {
        *typeinfo = ti_worksheet;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Worksheet, &ti_worksheet);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_worksheet;
    return hres;
}

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

    if (This->pAllRange==NULL) return E_FAIL;

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
    hres = _I_RangeConstructor((LPVOID*) &punk);

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
        I_Worksheet* iface, long lcid)
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
            hres = _I_RangeConstructor((LPVOID*) &punk);

            if (FAILED(hres)) return E_NOINTERFACE;

            hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pRange);
            if (pRange == NULL) {
                return E_FAIL;
            }
            struct CELL_COORD cell1,cell2;
            cell1.x = 1;
            cell1.y = V_I4(&Row);
            cell2.x = 256;
            cell2.y = V_I4(&Row);
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
                hres = _I_RangeConstructor((LPVOID*) &punk);

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
            hres = _I_RangeConstructor((LPVOID*) &punk);

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
                hres = _I_RangeConstructor((LPVOID*) &punk);

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
        VARIANT After,
        long lcid)
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
        I_Worksheet* iface, long lcid)
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

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Protect(
        I_Worksheet* iface,
        VARIANT Password,
        VARIANT DrawingObjects,
        VARIANT Contents,
        VARIANT Scenarios,
        VARIANT UserInterfaceOnly,
        VARIANT AllowFormattingCells,
        VARIANT AllowFormattingColumns,
        VARIANT AllowFormattingRows,
        VARIANT AllowInsertingColumns,
        VARIANT AllowInsertingRows,
        VARIANT AllowInsertingHyperlinks,
        VARIANT AllowDeletingColumns,
        VARIANT AllowDeletingRows,
        VARIANT AllowSorting,
        VARIANT AllowFiltering,
        VARIANT AllowUsingPivotTables)
{
    /*TODO Think about other parameters*/
    WorksheetImpl *This = (WorksheetImpl*)iface;
    VARIANT param, res;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    VariantInit(&param);
    if ((V_VT(&Password)==VT_EMPTY)||(V_VT(&Password)==VT_NULL)) {
        V_VT(&param) = VT_BSTR;
        V_BSTR(&param) = SysAllocString(L"");
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheet, L"protect", 1, param);
        if (FAILED(hres)) {
            TRACE("ERROR when protect\n");
            return hres;
        }
    } else {
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheet, L"protect", 1, Password);
        if (FAILED(hres)) {
            TRACE("ERROR when protect\n");
            return hres;
        }
    }

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Unprotect(
        I_Worksheet* iface,
        VARIANT Password,
        long lcid)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    VARIANT param, res;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    VariantInit(&param);
    if ((V_VT(&Password)==VT_EMPTY)||(V_VT(&Password)==VT_NULL)) {
        V_VT(&param) = VT_BSTR;
        V_BSTR(&param) = SysAllocString(L"");
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheet, L"unprotect", 1, param);
        if (FAILED(hres)) {
            TRACE("ERROR when unprotect\n");
            return hres;
        }
    } else {
        hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheet, L"unprotect", 1, Password);
        if (FAILED(hres)) {
            TRACE("ERROR when unprotect\n");
            return hres;
        }
    }

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Shapes(
        I_Worksheet* iface,
        IDispatch **ppValue)
{
    HRESULT hres;
    IUnknown *pObj;

    TRACE("\n");

    hres = _I_ShapesConstructor((void**)&pObj);
    if (FAILED(hres)) {
        TRACE(" ERROR when call constructor IShapes\n");
        return E_FAIL;
    }

    hres = I_Shapes_QueryInterface(pObj, &IID_I_Shapes, (void**)ppValue);
    if (FAILED(hres)) {
        TRACE(" ERROR when call IShapes->QueryInterface\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Shapes_Initialize((I_Shapes*)*ppValue, iface);
    if (FAILED(hres)) {
        TRACE(" ERROR when call Shape initialize\n");
        return E_FAIL;
    }

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Application(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Creator(
        I_Worksheet* iface,
        XlCreator *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Parent(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_CodeName(
        I_Worksheet* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get__CodeName(
        I_Worksheet* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put__CodeName(
        I_Worksheet* iface,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Index(
        I_Worksheet* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Move(
        I_Worksheet* iface,
        VARIANT Before,
        VARIANT After,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Next(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_OnDoubleClick(
        I_Worksheet* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_OnDoubleClick(
        I_Worksheet* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_OnSheetActivate(
        I_Worksheet* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_OnSheetActivate(
        I_Worksheet* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_OnSheetDeactivate(
        I_Worksheet* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_OnSheetDeactivate(
        I_Worksheet* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Previous(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet__PrintOut(
        I_Worksheet* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_PrintPreview(
        I_Worksheet* iface,
        VARIANT EnableChanges,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet__Protect(
        I_Worksheet* iface,
        VARIANT Password,
        VARIANT DrawingObjects,
        VARIANT Contents,
        VARIANT Scenarios,
        VARIANT UserInterfaceOnly,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ProtectContents(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ProtectDrawingObjects(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ProtectionMode(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ProtectScenarios(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet__SaveAs(
        I_Worksheet* iface,
        BSTR Filename,
        VARIANT FileFormat,
        VARIANT Password,
        VARIANT WriteResPassword,
        VARIANT ReadOnlyRecommended,
        VARIANT CreateBackup,
        VARIANT AddToMru,
        VARIANT TextCodepage,
        VARIANT TextVisualLayout,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Select(
        I_Worksheet* iface,
        VARIANT Replace,
        long lcid)
{
    WorksheetImpl *This = (WorksheetImpl*)(iface);

    TRACE("\n");

    if (This == NULL) return E_POINTER;
    if (This->pOOSheet == NULL) {
        TRACE("ERROR OOSheet = NULL \n");
        return E_POINTER;
    }

    WorkbookImpl *wb = (WorkbookImpl*)(This->pwb);

    VARIANT vRes,vRet,param;
    VariantInit(&vRes);
    VariantInit(&vRet);
    VariantInit(&param);
    HRESULT hres;

    hres = VariantChangeTypeEx(&Replace, &Replace, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }

    if (V_BOOL(&Replace)==VARIANT_TRUE) {
            hres = AutoWrap(DISPATCH_METHOD, &vRes, wb->pDoc,   L"getCurrentController",0);
        if (FAILED(hres)) {
            TRACE("ERROR when getCurrentController \n");
            return hres;
        }

        V_VT(&param) = VT_DISPATCH;
        V_DISPATCH(&param) = This->pOOSheet;
        IDispatch_AddRef(V_DISPATCH(&param));

        hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vRes), L"Select",1,param);

        if (FAILED(hres)) {
            TRACE("ERROR when Select \n");
            return hres;
        }
    } else {
        TRACE("Get VARIANT_FALSE as parameter \n");
    }
    VariantClear(&vRes);
    VariantClear(&vRet);
    VariantClear(&param);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Visible(
        I_Worksheet* iface,
        long lcid,
        XlSheetVisibility *RHS)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    VARIANT res;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VariantInit(&res);
    hres = AutoWrap(DISPATCH_PROPERTYGET, &res, This->pOOSheet, L"IsVisible", 0);

    if (FAILED(hres)) TRACE("ERROR when IsVisible \n");

    hres = VariantChangeTypeEx(&res, &res, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx \n");
        return E_FAIL;
    }

    switch (V_BOOL(&res)) {
        case VARIANT_TRUE:
            *RHS = xlSheetVisible;
            break;
        case VARIANT_FALSE:
            *RHS = xlSheetHidden;
            break;
    }

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_Visible(
        I_Worksheet* iface,
        long lcid,
        XlSheetVisibility RHS)
{
    WorksheetImpl *This = (WorksheetImpl*)iface;
    VARIANT res, param1;
    HRESULT hres;

    TRACE("\n");
    VariantInit(&param1);

    if (This == NULL) return E_POINTER;

    V_VT(&param1) = VT_BOOL;
    switch (RHS) {
        case xlSheetVeryHidden:
        case xlSheetHidden:
            V_BOOL(&param1) = VARIANT_FALSE;
            break;
        case xlSheetVisible:
            V_BOOL(&param1) = VARIANT_TRUE;
            break;
    }

    VariantInit(&res);
    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOOSheet, L"IsVisible", 1, param1);

    if (FAILED(hres)) TRACE("ERROR when IsVisible \n");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_TransitionExpEval(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_TransitionExpEval(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Arcs(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_AutoFilterMode(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_AutoFilterMode(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_SetBackgroundPicture(
        I_Worksheet* iface,
        BSTR Filename)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Buttons(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Calculate(
        I_Worksheet* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_EnableCalculation(
        I_Worksheet* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_EnableCalculation(
        I_Worksheet* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ChartObjects(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_CheckBoxes(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_CheckSpelling(
        I_Worksheet* iface,
        VARIANT CustomDictionary,
        VARIANT IgnoreUppercase,
        VARIANT AlwaysSuggest,
        VARIANT SpellLang,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_CircularReference(
        I_Worksheet* iface,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ClearArrows(
        I_Worksheet* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ConsolidationFunction(
        I_Worksheet* iface,
        long lcid,
        XlConsolidationFunction *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ConsolidationOptions(
        I_Worksheet* iface,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ConsolidationSources(
        I_Worksheet* iface,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_DisplayAutomaticPageBreaks(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_DisplayAutomaticPageBreaks(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Drawings(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_DrawingObjects(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_DropDowns(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_EnableAutoFilter(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_EnableAutoFilter(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_EnableSelection(
        I_Worksheet* iface,
        XlEnableSelection *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_EnableSelection(
        I_Worksheet* iface,
        XlEnableSelection RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_EnableOutlining(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_EnableOutlining(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_EnablePivotTable(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_EnablePivotTable(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Evaluate(
        I_Worksheet* iface,
        VARIANT Name,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet__Evaluate(
        I_Worksheet* iface,
        VARIANT Name,
        long lcid,
        VARIANT *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_FilterMode(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ResetAllPageBreaks(
        I_Worksheet* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_GroupBoxes(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_GroupObjects(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Labels(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Lines(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ListBoxes(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Names(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_OLEObjects(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_OnCalculate(
        I_Worksheet* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_OnCalculate(
        I_Worksheet* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_OnData(
        I_Worksheet* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_OnData(
        I_Worksheet* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_OnEntry(
        I_Worksheet* iface,
        long lcid,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_OnEntry(
        I_Worksheet* iface,
        long lcid,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_OptionButtons(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Outline(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    HRESULT hres;
    IUnknown *pObj;

    TRACE("\n");

    hres = _I_OutlineConstructor((void**)&pObj);
    if (FAILED(hres)) {
        TRACE(" ERROR when call constructor IOutline\n");
        return E_FAIL;
    }

    hres = I_Outline_QueryInterface(pObj, &IID_I_Outline, (void**)RHS);
    if (FAILED(hres)) {
        TRACE(" ERROR when call IOutline->QueryInterface\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Outline_Initialize((I_Outline*)*RHS, iface);
    if (FAILED(hres)) {
        TRACE(" ERROR when call Outline initialize\n");
        return E_FAIL;
    }

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Ovals(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet__PasteSpecial(
        I_Worksheet* iface,
        VARIANT Format,
        VARIANT Link,
        VARIANT DisplayAsIcon,
        VARIANT IconFileName,
        VARIANT IconIndex,
        VARIANT IconLabel,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Pictures(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_PivotTables(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_PivotTableWizard(
        I_Worksheet* iface,
        VARIANT SourceType,
        VARIANT SourceData,
        VARIANT TableDestination,
        VARIANT TableName,
        VARIANT RowGrand,
        VARIANT ColumnGrand,
        VARIANT SaveData,
        VARIANT HasAutoFormat,
        VARIANT AutoPage,
        VARIANT Reserved,
        VARIANT BackgroundQuery,
        VARIANT OptimizeCache,
        VARIANT PageFieldOrder,
        VARIANT PageFieldWrapCount,
        VARIANT ReadData,
        VARIANT Connection,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Rectangles(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Scenarios(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ScrollArea(
        I_Worksheet* iface,
        BSTR *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_ScrollArea(
        I_Worksheet* iface,
        BSTR RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ScrollBars(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ShowAllData(
        I_Worksheet* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ShowDataForm(
        I_Worksheet* iface,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_Spinners(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_StandardHeight(
        I_Worksheet* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_StandardWidth(
        I_Worksheet* iface,
        long lcid,
        double *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_StandardWidth(
        I_Worksheet* iface,
        long lcid,
        double RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_TextBoxes(
        I_Worksheet* iface,
        VARIANT Index,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_TransitionFormEntry(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_TransitionFormEntry(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Type(
        I_Worksheet* iface,
        long lcid,
        XlSheetType *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_UsedRange(
        I_Worksheet* iface,
        long lcid,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_HPageBreaks(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_VPageBreaks(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_QueryTables(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_DisplayPageBreaks(
        I_Worksheet* iface,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_DisplayPageBreaks(
        I_Worksheet* iface,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Comments(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Hyperlinks(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_ClearCircles(
        I_Worksheet* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_CircleInvalid(
        I_Worksheet* iface)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get__DisplayRightToLeft(
        I_Worksheet* iface,
        long lcid,
        long *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put__DisplayRightToLeft(
        I_Worksheet* iface,
        long lcid,
        long RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_AutoFilter(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_DisplayRightToLeft(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL *RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_put_DisplayRightToLeft(
        I_Worksheet* iface,
        long lcid,
        VARIANT_BOOL RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Scripts(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_PrintOut(
        I_Worksheet* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        VARIANT PrToFileName,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet__CheckSpelling(
        I_Worksheet* iface,
        VARIANT CustomDictionary,
        VARIANT IgnoreUppercase,
        VARIANT AlwaysSuggest,
        VARIANT SpellLang,
        VARIANT IgnoreFinalYaa,
        VARIANT SpellScript,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Tab(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_MailEnvelope(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_SaveAs(
        I_Worksheet* iface,
        BSTR Filename,
        VARIANT FileFormat,
        VARIANT Password,
        VARIANT WriteResPassword,
        VARIANT ReadOnlyRecommended,
        VARIANT CreateBackup,
        VARIANT AddToMru,
        VARIANT TextCodepage,
        VARIANT TextVisualLayout,
        VARIANT Local)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_CustomProperties(
        I_Worksheet* iface,
        IDispatch     **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_SmartTags(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_Protection(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_PasteSpecial(
        I_Worksheet* iface,
        VARIANT Format,
        VARIANT Link,
        VARIANT DisplayAsIcon,
        VARIANT IconFileName,
        VARIANT IconIndex,
        VARIANT IconLabel,
        VARIANT NoHTMLFormatting,
        long lcid)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_get_ListObjects(
        I_Worksheet* iface,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_XmlDataQuery(
        I_Worksheet* iface,
        BSTR XPath,
        VARIANT SelectionNamespaces,
        VARIANT Map,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Worksheet_XmlMapQuery(
        I_Worksheet* iface,
        BSTR XPath,
        VARIANT SelectionNamespaces,
        VARIANT Map,
        IDispatch **RHS)
{
    TRACE("\n");
    return E_NOTIMPL;
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
    ITypeInfo *typeinfo;
    HRESULT hres;

    hres = get_typeinfo_worksheet(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }

    return hres;
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
    VARIANT vmas[16];
    int i;
    long ltmp;

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
    case dispid_worksheet_name:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &tmpval);
            hres = MSO_TO_OO_I_Worksheet_put_Name(iface, V_BSTR(&tmpval));
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
    case dispid_worksheet_cells:
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
                I_Range_put_Value((I_Range*)pretdisp, vNull, 0, tmpval);
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
                } else {
                    IDispatch_Release(dret);
                }
                break;
            case 1:
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2);
                I_Range_get__Default((I_Range*)dret,cell2, vNull, &pretdisp);
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=pretdisp;
                } else {
                    IDispatch_Release(pretdisp);
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
                } else {
                    IDispatch_Release(pretdisp);
                }
                I_Range_Release((I_Range*)dret);
               break;
            }
            return S_OK;
        }
    case dispid_worksheet_range:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 1:
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1);
                hres = MSO_TO_OO_I_Worksheet_get_Range(iface, cell1, vNull, &dret);
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
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
            default:
                TRACE("(case 3) ERROR PARAMETR IS SEND\n");
                return E_FAIL;
            }
        }
    case dispid_worksheet_paste:
        /*method paste   MSO_TO_OO_I_Worksheet_Paste*/

       return E_NOTIMPL;
    case dispid_worksheet_activate:
       return MSO_TO_OO_I_Worksheet_Activate(iface, 0);
    case dispid_worksheet_rows://Rows
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
                } else {
                    IDispatch_Release(dret);
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
                } else {
                    IDispatch_Release(dret);
                }
                return S_OK;
            }
        }
    case dispid_worksheet_columns://Columns
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
                } else {
                    IDispatch_Release(dret);
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
                } else {
                    IDispatch_Release(dret);
                }
                return S_OK;
            }
        }
    case dispid_worksheet_copy:
        switch (pDispParams->cArgs) {
        case 0:
            TRACE("(case 8) 0 Parameter\n");
            hres = MSO_TO_OO_I_Worksheet_Copy(iface, vNull, vNull, 0);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        case 1:
            TRACE("(case 8) 1 Parameter\n");
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2))) return E_FAIL;

            hres = MSO_TO_OO_I_Worksheet_Copy(iface, cell2, vNull, 0);
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
            hres = MSO_TO_OO_I_Worksheet_Copy(iface, cell1, cell2, 0);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        }
    case dispid_worksheet_delete:
        return MSO_TO_OO_I_Worksheet_Delete(iface, 0);
    case dispid_worksheet_pagesetup:
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
    case dispid_worksheet_protect://Protect
        for (i=0;i<16;i++) {
            VariantInit(&vmas[i]);
            V_VT(&vmas[i])=VT_EMPTY;
        }
        /*необходимо перевернуть параметры*/
        for (i=0;i<pDispParams->cArgs;i++) {
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[pDispParams->cArgs-i-1], &vmas[i]))) return E_FAIL;
        }
        return MSO_TO_OO_I_Worksheet_Protect(iface, vmas[0], vmas[1], vmas[2], vmas[3], vmas[4], vmas[5], vmas[6], vmas[7], vmas[8], vmas[9], vmas[10], vmas[11], vmas[12], vmas[13], vmas[14], vmas[15]);
    case dispid_worksheet_unprotect://UnProtect
        switch (pDispParams->cArgs) {
        case 0:
            VariantClear(&cell1);
            V_VT(&cell1) = VT_EMPTY;
            break;
        case 1:
            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell1))) return E_FAIL;
            break;
        default:
            TRACE("ERROR parameters \n");
            return E_INVALIDARG;
        }
        return MSO_TO_OO_I_Worksheet_Unprotect(iface,cell1, 0);
    case dispid_worksheet_shapes:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 0:
                TRACE("(case 13) 0 Parameter\n");
                hres = MSO_TO_OO_I_Worksheet_get_Shapes(iface, &dret);
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
            case 1:
                TRACE("(case 13) 1 Parameter\n");
                if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2))) return E_FAIL;

                hres = MSO_TO_OO_I_Worksheet_get_Shapes(iface, &dret);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                /*TODO use parameters*/

                if (pVarResult!=NULL){
                    V_VT(pVarResult) = VT_DISPATCH;
                    V_DISPATCH(pVarResult) = dret;
                } else {
                    IDispatch_Release(dret);
                }
                return S_OK;
            }
        }
    case dispid_worksheet_outline:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            switch (pDispParams->cArgs) {
            case 0:
                TRACE("0 Parameter\n");
                hres = MSO_TO_OO_I_Worksheet_get_Outline(iface, &dret);
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
            default:
                TRACE("ERROR parameters \n");
                return E_FAIL;
            }
        }
    case dispid_worksheet_visible:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &tmpval);
            hres = VariantChangeTypeEx(&tmpval, &tmpval, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("(case 8) ERROR VariantChangeTypeEx   %08x\n",hres);
                return E_FAIL;
            }
            ltmp = V_I4(&tmpval);
            hres = MSO_TO_OO_I_Worksheet_put_Visible(iface, 0, (XlSheetVisibility)ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Worksheet_get_Visible(iface, 0, (XlSheetVisibility*)&ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = ltmp;
            }
            return S_OK;
        }
    case dispid_worksheet_select://Method select
        switch (pDispParams->cArgs) {
        case 1:
            TRACE("ERROR parameters number = %i \n", pDispParams->cArgs);
            return E_FAIL;

            if (FAILED(MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &cell2))) {
                TRACE("ERROR when CorrectArg \n");
                return E_FAIL;
            }
            break;
        case 0:
            V_VT(&cell2) = VT_BOOL;
            V_BOOL(&cell2) = VARIANT_TRUE;
            break;
        }
        return MSO_TO_OO_I_Worksheet_Select(iface, cell2, 0);
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
    MSO_TO_OO_I_Worksheet_get_Application,
    MSO_TO_OO_I_Worksheet_get_Creator,
    MSO_TO_OO_I_Worksheet_get_Parent,
    MSO_TO_OO_I_Worksheet_Activate,
    MSO_TO_OO_I_Worksheet_Copy,
    MSO_TO_OO_I_Worksheet_Delete,
    MSO_TO_OO_I_Worksheet_get_CodeName,
    MSO_TO_OO_I_Worksheet_get__CodeName,
    MSO_TO_OO_I_Worksheet_put__CodeName,
    MSO_TO_OO_I_Worksheet_get_Index,
    MSO_TO_OO_I_Worksheet_Move,
    MSO_TO_OO_I_Worksheet_get_Name,
    MSO_TO_OO_I_Worksheet_put_Name,
    MSO_TO_OO_I_Worksheet_get_Next,
    MSO_TO_OO_I_Worksheet_get_OnDoubleClick,
    MSO_TO_OO_I_Worksheet_put_OnDoubleClick,
    MSO_TO_OO_I_Worksheet_get_OnSheetActivate,
    MSO_TO_OO_I_Worksheet_put_OnSheetActivate,
    MSO_TO_OO_I_Worksheet_get_OnSheetDeactivate,
    MSO_TO_OO_I_Worksheet_put_OnSheetDeactivate,
    MSO_TO_OO_I_Worksheet_get_PageSetup,
    MSO_TO_OO_I_Worksheet_get_Previous,
    MSO_TO_OO_I_Worksheet__PrintOut,
    MSO_TO_OO_I_Worksheet_PrintPreview,
    MSO_TO_OO_I_Worksheet__Protect,
    MSO_TO_OO_I_Worksheet_get_ProtectContents,
    MSO_TO_OO_I_Worksheet_get_ProtectDrawingObjects,
    MSO_TO_OO_I_Worksheet_get_ProtectionMode,
    MSO_TO_OO_I_Worksheet_get_ProtectScenarios,
    MSO_TO_OO_I_Worksheet__SaveAs,
    MSO_TO_OO_I_Worksheet_Select,
    MSO_TO_OO_I_Worksheet_Unprotect,
    MSO_TO_OO_I_Worksheet_get_Visible,
    MSO_TO_OO_I_Worksheet_put_Visible,
    MSO_TO_OO_I_Worksheet_get_Shapes,
    MSO_TO_OO_I_Worksheet_get_TransitionExpEval,
    MSO_TO_OO_I_Worksheet_put_TransitionExpEval,
    MSO_TO_OO_I_Worksheet_Arcs,
    MSO_TO_OO_I_Worksheet_get_AutoFilterMode,
    MSO_TO_OO_I_Worksheet_put_AutoFilterMode,
    MSO_TO_OO_I_Worksheet_SetBackgroundPicture,
    MSO_TO_OO_I_Worksheet_Buttons,
    MSO_TO_OO_I_Worksheet_Calculate,
    MSO_TO_OO_I_Worksheet_get_EnableCalculation,
    MSO_TO_OO_I_Worksheet_put_EnableCalculation,
    MSO_TO_OO_I_Worksheet_get_Cells,
    MSO_TO_OO_I_Worksheet_ChartObjects,
    MSO_TO_OO_I_Worksheet_CheckBoxes,
    MSO_TO_OO_I_Worksheet_CheckSpelling,
    MSO_TO_OO_I_Worksheet_get_CircularReference,
    MSO_TO_OO_I_Worksheet_ClearArrows,
    MSO_TO_OO_I_Worksheet_get_Columns,
    MSO_TO_OO_I_Worksheet_get_ConsolidationFunction,
    MSO_TO_OO_I_Worksheet_get_ConsolidationOptions,
    MSO_TO_OO_I_Worksheet_get_ConsolidationSources,
    MSO_TO_OO_I_Worksheet_get_DisplayAutomaticPageBreaks,
    MSO_TO_OO_I_Worksheet_put_DisplayAutomaticPageBreaks,
    MSO_TO_OO_I_Worksheet_Drawings,
    MSO_TO_OO_I_Worksheet_DrawingObjects,
    MSO_TO_OO_I_Worksheet_DropDowns,
    MSO_TO_OO_I_Worksheet_get_EnableAutoFilter,
    MSO_TO_OO_I_Worksheet_put_EnableAutoFilter,
    MSO_TO_OO_I_Worksheet_get_EnableSelection,
    MSO_TO_OO_I_Worksheet_put_EnableSelection,
    MSO_TO_OO_I_Worksheet_get_EnableOutlining,
    MSO_TO_OO_I_Worksheet_put_EnableOutlining,
    MSO_TO_OO_I_Worksheet_get_EnablePivotTable,
    MSO_TO_OO_I_Worksheet_put_EnablePivotTable,
    MSO_TO_OO_I_Worksheet_Evaluate,
    MSO_TO_OO_I_Worksheet__Evaluate,
    MSO_TO_OO_I_Worksheet_get_FilterMode,
    MSO_TO_OO_I_Worksheet_ResetAllPageBreaks,
    MSO_TO_OO_I_Worksheet_GroupBoxes,
    MSO_TO_OO_I_Worksheet_GroupObjects,
    MSO_TO_OO_I_Worksheet_Labels,
    MSO_TO_OO_I_Worksheet_Lines,
    MSO_TO_OO_I_Worksheet_ListBoxes,
    MSO_TO_OO_I_Worksheet_get_Names,
    MSO_TO_OO_I_Worksheet_OLEObjects,
    MSO_TO_OO_I_Worksheet_get_OnCalculate,
    MSO_TO_OO_I_Worksheet_put_OnCalculate,
    MSO_TO_OO_I_Worksheet_get_OnData,
    MSO_TO_OO_I_Worksheet_put_OnData,
    MSO_TO_OO_I_Worksheet_get_OnEntry,
    MSO_TO_OO_I_Worksheet_put_OnEntry,
    MSO_TO_OO_I_Worksheet_OptionButtons,
    MSO_TO_OO_I_Worksheet_get_Outline,
    MSO_TO_OO_I_Worksheet_Ovals,
    MSO_TO_OO_I_Worksheet_Paste,
    MSO_TO_OO_I_Worksheet__PasteSpecial,
    MSO_TO_OO_I_Worksheet_Pictures,
    MSO_TO_OO_I_Worksheet_PivotTables,
    MSO_TO_OO_I_Worksheet_PivotTableWizard,
    MSO_TO_OO_I_Worksheet_get_Range,
    MSO_TO_OO_I_Worksheet_Rectangles,
    MSO_TO_OO_I_Worksheet_get_Rows,
    MSO_TO_OO_I_Worksheet_Scenarios,
    MSO_TO_OO_I_Worksheet_get_ScrollArea,
    MSO_TO_OO_I_Worksheet_put_ScrollArea,
    MSO_TO_OO_I_Worksheet_ScrollBars,
    MSO_TO_OO_I_Worksheet_ShowAllData,
    MSO_TO_OO_I_Worksheet_ShowDataForm,
    MSO_TO_OO_I_Worksheet_Spinners,
    MSO_TO_OO_I_Worksheet_get_StandardHeight,
    MSO_TO_OO_I_Worksheet_get_StandardWidth,
    MSO_TO_OO_I_Worksheet_put_StandardWidth,
    MSO_TO_OO_I_Worksheet_TextBoxes,
    MSO_TO_OO_I_Worksheet_get_TransitionFormEntry,
    MSO_TO_OO_I_Worksheet_put_TransitionFormEntry,
    MSO_TO_OO_I_Worksheet_get_Type,
    MSO_TO_OO_I_Worksheet_get_UsedRange,
    MSO_TO_OO_I_Worksheet_get_HPageBreaks,
    MSO_TO_OO_I_Worksheet_get_VPageBreaks,
    MSO_TO_OO_I_Worksheet_get_QueryTables,
    MSO_TO_OO_I_Worksheet_get_DisplayPageBreaks,
    MSO_TO_OO_I_Worksheet_put_DisplayPageBreaks,
    MSO_TO_OO_I_Worksheet_get_Comments,
    MSO_TO_OO_I_Worksheet_get_Hyperlinks,
    MSO_TO_OO_I_Worksheet_ClearCircles,
    MSO_TO_OO_I_Worksheet_CircleInvalid,
    MSO_TO_OO_I_Worksheet_get__DisplayRightToLeft,
    MSO_TO_OO_I_Worksheet_put__DisplayRightToLeft,
    MSO_TO_OO_I_Worksheet_get_AutoFilter,
    MSO_TO_OO_I_Worksheet_get_DisplayRightToLeft,
    MSO_TO_OO_I_Worksheet_put_DisplayRightToLeft,
    MSO_TO_OO_I_Worksheet_get_Scripts,
    MSO_TO_OO_I_Worksheet_PrintOut,
    MSO_TO_OO_I_Worksheet__CheckSpelling,
    MSO_TO_OO_I_Worksheet_get_Tab,
    MSO_TO_OO_I_Worksheet_get_MailEnvelope,
    MSO_TO_OO_I_Worksheet_SaveAs,
    MSO_TO_OO_I_Worksheet_get_CustomProperties,
    MSO_TO_OO_I_Worksheet_get_SmartTags,
    MSO_TO_OO_I_Worksheet_get_Protection,
    MSO_TO_OO_I_Worksheet_PasteSpecial,
    MSO_TO_OO_I_Worksheet_Protect,
    MSO_TO_OO_I_Worksheet_get_ListObjects,
    MSO_TO_OO_I_Worksheet_XmlDataQuery,
    MSO_TO_OO_I_Worksheet_XmlMapQuery
};

extern HRESULT _I_WorksheetConstructor(LPVOID *ppObj)
{
    WorksheetImpl *worksheet;

    TRACE("(%p)\n", ppObj);

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
