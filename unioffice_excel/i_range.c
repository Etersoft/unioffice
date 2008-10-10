/*
 * IRange interface functions
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
#include "oleauto.h"

ITypeInfo *ti_range = NULL;

HRESULT get_typeinfo_range(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_range) {
        *typeinfo = ti_range;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Range, &ti_range);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_range;
    return hres;
}

/*флаги для работы с ячейками*/
const long VALUE 	= 1;
const long DATETIME 	= 2;
const long STRING 	= 4;
const long ANNOTATION 	= 8;
const long FORMULA 	= 16;
const long HARDATTR 	= 32;
const long STYLES 	= 64;
const long OBJECTS 	= 128;
const long EDITATTR 	= 256;
const long FORMATTED 	= 512;

typedef enum {
    NONE,
    DOWN,
    RIGHT,
    ROWS,
    COLUMNS
} CellInsertMode;

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Range_AddRef(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_QueryInterface(
        I_Range* iface,
        REFIID riid,
        void **ppvObject)
{
    RangeImpl *This = (RangeImpl*)iface;

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Range)) {
        *ppvObject = &This->_rangeVtbl;
        MSO_TO_OO_I_Range_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Range_Release(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pOORange != NULL) {
            IDispatch_Release(This->pOORange);
            This->pOORange = NULL;
        }
        if ((This->pwsheet != NULL)&&(This->is_release==1)) {
            IDispatch_Release(This->pwsheet);
            This->pwsheet = NULL;
        }
        if (This->pApplication != NULL) {
            IDispatch_Release(This->pApplication);
            This->pApplication = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Range methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Range_get__Default(
        I_Range* iface,
        VARIANT varRowIndex,
        VARIANT varColumnIndex,
        IDispatch **ppObject)
{
    RangeImpl *This = (RangeImpl*)iface;
    I_Range *pCell;
    IUnknown *punk = NULL;
    HRESULT hres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(varRowIndex, &varRowIndex);
    MSO_TO_OO_CorrectArg(varColumnIndex, &varColumnIndex);

    if (This == NULL) return E_POINTER;

    *ppObject = NULL;

    if (V_VT(&varRowIndex)==VT_BSTR) {

        hres = _I_RangeConstructor((LPVOID*) &punk);
        if (FAILED(hres)) return E_NOINTERFACE;

        hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pCell);

        if (pCell == NULL) {
            return E_FAIL;
        }

        hres = MSO_TO_OO_I_Range_Initialize_ByName(pCell, iface, varRowIndex);
        if (FAILED(hres)){
            I_Range_Release(pCell);
            return hres;
        }

        *ppObject = (IDispatch*)pCell;
        I_Range_AddRef((I_Range*)*ppObject);
        I_Range_Release(pCell);
        TRACE_OUT;
        return S_OK;
    } else {
        /*преобразовываем любой тип к I4*/

        hres = VariantChangeTypeEx(&varRowIndex, &varRowIndex, 0, 0, VT_I4);

        hres = VariantChangeTypeEx(&varColumnIndex, &varColumnIndex, 0, 0, VT_I4);

        if ((V_VT(&varRowIndex) != VT_I4) || (V_VT(&varColumnIndex) != VT_I4))
            return E_FAIL;

        /*Создаем новый объект I_Range that occupy one cell*/
        struct CELL_COORD cellCoord;

        hres = _I_RangeConstructor((LPVOID*) &punk);

        if (FAILED(hres)) return E_NOINTERFACE;

        hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pCell);

        if (pCell == NULL) {
            return E_FAIL;
        }

        cellCoord.x = V_I4(&varColumnIndex);
        cellCoord.y = V_I4(&varRowIndex);
	
        TRACE("cellCoord.x=%i, cellCoord.y=%i \n",cellCoord.x,cellCoord.y);
        hres = MSO_TO_OO_I_Range_Initialize((I_Range*)pCell, iface, cellCoord, cellCoord);
        if (FAILED(hres)){
            I_Range_Release(pCell);
            return hres;
        }

        *ppObject = (IDispatch*)pCell;
        I_Range_AddRef((I_Range*)*ppObject);
        I_Range_Release(pCell);
        TRACE_OUT;
        return S_OK;
    }
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_ColumnWidth(
        I_Range* iface,
        long *pnColumnWidth)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    if (This->pOORange == NULL)
        return E_POINTER;

    VARIANT columns;
    HRESULT hres;
    VARIANT vColWidth;

    hres = AutoWrap(DISPATCH_METHOD, &columns, This->pOORange, L"getColumns", 0);
    if (hres != S_OK)
        return hres;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vColWidth, V_DISPATCH(&columns), L"Width", 0);

    IDispatch_Release(V_DISPATCH(&columns));
    *pnColumnWidth = V_I2(&vColWidth)/200;

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_ColumnWidth(
        I_Range* iface,
        long nColumnWidth)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    if (This->pOORange == NULL)
        return E_POINTER;

/* It is a some different between MS excel and OO spreadsheet
 in MS excell we may get and put ColumnWidth in chars,
 in OO spreadsheet we setiing this sizes in units 10HTMM
 for default font the appropriate ratio about 1 char  == 200(10HTMM) */

    VARIANT res;
    VARIANT columns;
    HRESULT hres;
    VARIANT vColWidth;

    hres = AutoWrap(DISPATCH_METHOD, &columns, This->pOORange, L"getColumns", 0);
    if (hres != S_OK)
        return hres;

    VariantInit(&vColWidth);
    V_VT(&vColWidth) = VT_I4;
    V_I4(&vColWidth) = nColumnWidth*210;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, V_DISPATCH(&columns), L"Width", 1, vColWidth);
    IDispatch_Release(V_DISPATCH(&columns));

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Font(
        I_Range* iface,
        IDispatch **ppFont)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hr;
    IUnknown *punk = NULL;
    IDispatch *pFont;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *ppFont = NULL;

    hr = _I_FontConstructor((LPVOID*) &punk);

    if (FAILED(hr)) return E_NOINTERFACE;

    hr = I_Font_QueryInterface(punk, &IID_I_Font, (void**) &pFont);
    if (pFont == NULL) {
        return E_FAIL;
    }

    hr = MSO_TO_OO_I_Font_Initialize((I_Font*)pFont, iface);

    if (FAILED(hr)) {
        IDispatch_Release(pFont);
        return hr;
    }

    *ppFont = pFont;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Value(
        I_Range* iface,
        VARIANT varRangeValueDataType,
        LCID lcid,
        VARIANT *pvarValue)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(varRangeValueDataType, &varRangeValueDataType);

    if (This == NULL) return E_POINTER;

    VARIANT resultCell;
    HRESULT hres;
    VARIANT parmRow; // XPOS
    VARIANT vartype;

    VariantInit(&vartype);
    V_VT(&parmRow) = VT_I2;
    V_I2(&parmRow) = 0;
    VARIANT parmColumn; // nYPos
    V_VT(&parmColumn) = VT_I2;
    V_I2(&parmColumn) = 0;

    hres = AutoWrap(DISPATCH_METHOD, &resultCell, This->pOORange, L"getCellByPosition", 2, parmRow, parmColumn);
    if (hres != S_OK) {
        TRACE("ERROR when getCellByPosition \n");
        return hres;
    }
/*Необходимо узнать тип ячейки и после этого уже читать значение*/
    hres = AutoWrap(DISPATCH_METHOD, &vartype, V_DISPATCH(&resultCell), L"getType", 0);

    switch V_I4(&vartype){
    case vtFORMULA:
    case vtVALUE:
        hres = AutoWrap(DISPATCH_METHOD, pvarValue, V_DISPATCH(&resultCell), L"getValue", 0);
        break;
    case vtEMPTY:
        V_VT(pvarValue)=VT_EMPTY;
        hres = S_OK;
        break;
    case vtTEXT:
    default:
        hres = AutoWrap(DISPATCH_METHOD, pvarValue, V_DISPATCH(&resultCell), L"getFormula", 0);
    } 

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Value(
        I_Range* iface,
        VARIANT varRangeValueDataType,
        LCID lcid,
        VARIANT varValue)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(varRangeValueDataType, &varRangeValueDataType);
    MSO_TO_OO_CorrectArg(varValue, &varValue);

    if (This == NULL) return E_POINTER;

    VARIANT resultCell;
    HRESULT hres;
    VARIANT res;
    VARIANT parmRow; /* XPOS */
    V_VT(&parmRow) = VT_I2;
    V_I2(&parmRow) = 0;
    VARIANT parmColumn; /* nYPos */
    V_VT(&parmColumn) = VT_I2;
    V_I2(&parmColumn) = 0;

    if (V_VT(&varValue) & VT_ARRAY) {
/*Для Excel первое измерение массива это ряды, второе колонки*/
/*Для OpenOffice наоборот первое измерение массива это колонки, второе ряды*/
    int arr_dim;
    VARIANT *pvar;
    int i,j;

    arr_dim = SafeArrayGetDim(V_ARRAY(&varValue));

    /*Если одно измерение*/
    if (arr_dim == 1) {
    /*TODO*/
        TRACE("1 Demension array NOT REALIZE NOW \n");
    }
    /*Если два измерения*/
    if (arr_dim == 2) {
        long startrow,endrow,startcolumn,endcolumn;
        VARIANT row,col;
        VARIANT vNull;
        VariantInit(&vNull);
        I_Range *temp_range;

        hres=SafeArrayAccessData(V_ARRAY(&varValue), (void **)&pvar);
        if (FAILED(hres)) return hres;

        hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);
        TRACE("startrow=%i, startcolumn=%i, endrow=%i, endcolumn=%i \n",startrow, startcolumn, endrow, endcolumn);

        int maxj = (V_ARRAY(&varValue))->rgsabound[0].cElements;
        int maxi = (V_ARRAY(&varValue))->rgsabound[1].cElements;

        for (i=0; i<maxi; i++) {
            for (j=0; j<maxj; j++) {

                V_VT(&row) = VT_I4;
                V_I4(&row) = i + 1;
                V_VT(&col) = VT_I4;
                V_I4(&col) = j + 1;

                if ((i<=(endrow-startrow))&&(j<=(endcolumn-startcolumn))) {
                    hres = MSO_TO_OO_I_Range_get__Default(iface, row, col, (IDispatch**)&temp_range);
                    if (FAILED(hres)) {
                        TRACE("Error when Range_get_default row=%i col=%i\n",V_I4(&row),V_I4(&col));
                        hres = SafeArrayUnaccessData(V_ARRAY(&varValue));
                        return hres;
                    }

                    hres = MSO_TO_OO_I_Range_put_Value(temp_range, vNull, 0, pvar[j*maxi+i]);
                    if (FAILED(hres)) {
                        TRACE("Error when Range_put_Value \n");
                        hres = SafeArrayUnaccessData(V_ARRAY(&varValue));
                        return hres;
                    }
                    I_Range_Release(temp_range);
                }
            }
        }
        hres = SafeArrayUnaccessData(V_ARRAY(&varValue));
        if (FAILED(hres)) {
            TRACE("Error when SafeArrayUnaccessData \n");
        }
        return hres;
    }
    /*Если не равно ни 1 ни 2, то ничего не делаем*/
    TRACE_OUT;
    return S_OK;
    } else {

        hres = AutoWrap(DISPATCH_METHOD, &resultCell, This->pOORange, L"getCellByPosition", 2, parmRow, parmColumn);
        if (hres != S_OK) {
            TRACE("ERROR when getCellByPosition \n");
            return hres;
        }
        long tmp;
        /*Необходимо преобразование типов OpenOffice понимает не все типы*/
        switch V_VT(&varValue) {
        case VT_I8:/*надо преобразовать к VT_I4*/
            tmp = (long) V_I8(&varValue);
            VariantClear(&varValue);
            V_VT(&varValue) = VT_I4;
            V_I4(&varValue) = tmp;
        }
        switch V_VT(&varValue) {
        case VT_BSTR:
            if (lstrlenW(V_BSTR(&varValue))!=0) {
                if (V_BSTR(&varValue)[0]==L'=') {
                    hres = AutoWrap(DISPATCH_METHOD, &res, V_DISPATCH(&resultCell), L"setFormula", 1, varValue);
                    if (hres != S_OK) {
                        TRACE("ERROR when setFormula \n");
                        TRACE("    VT = %i \n",V_VT(&varValue));
                    }
                    return hres;
                }
            }
            hres = AutoWrap(DISPATCH_METHOD, &res, V_DISPATCH(&resultCell), L"setString", 1, varValue);
            if (hres != S_OK) {
                TRACE("ERROR when setString \n");
                TRACE("    VT = %i \n",V_VT(&varValue));
                return hres;
            }
            return hres;
        default:
            hres = AutoWrap(DISPATCH_METHOD, &res, V_DISPATCH(&resultCell), L"setValue", 1, varValue);
            if (hres != S_OK) {
                TRACE("ERROR when setFormula \n");
                TRACE("    VT = %i \n",V_VT(&varValue));
            }
            return hres;
        }
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Select(
        I_Range* iface,
        VARIANT *pvarResult)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;
    if (This->pOORange == NULL) {
        TRACE("ERROR OORange = NULL \n");
        return E_POINTER;
    }
    WorksheetImpl *wsh = (WorksheetImpl*)(This->pwsheet);
    WorkbookImpl *wb = (WorkbookImpl*)(wsh->pwb);

    VARIANT vRes,vRet,param;
    VariantInit(&vRes);
    VariantInit(&vRet);
    VariantInit(&param);
    HRESULT hres;
    V_VT(pvarResult) = VT_BOOL;
    V_BOOL(pvarResult) = VARIANT_FALSE;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, wb->pDoc, L"getCurrentController",0);
    if (FAILED(hres)) {
        TRACE("ERROR when getCurrentController \n");
        return hres;
    }

    V_VT(&param) = VT_DISPATCH;
    V_DISPATCH(&param) = This->pOORange;
    IDispatch_AddRef(V_DISPATCH(&param));

    hres = AutoWrap(DISPATCH_METHOD, &vRet, V_DISPATCH(&vRes), L"Select",1,param);

    if (FAILED(hres)) {
        TRACE("ERROR when Select \n");
        return hres;
    }

    V_VT(pvarResult) = VT_BOOL;
    V_BOOL(pvarResult) = VARIANT_TRUE;

    VariantClear(&vRes);
    VariantClear(&vRet);
    VariantClear(&param);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_NoteText(
        I_Range* iface,
        VARIANT Text,
        VARIANT Start,
        VARIANT Length,
        BSTR *pText)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(Text, &Text);
    MSO_TO_OO_CorrectArg(Start, &Start);
    MSO_TO_OO_CorrectArg(Length, &Length);

    if (This == NULL) return E_POINTER;

    VARIANT vRes;
    VARIANT parmRow;    /* XPOS */
    VARIANT parmColumn; /* YPOS */

    V_VT(&parmRow)=VT_I4;
    V_I4(&parmRow) = 0;
    V_VT(&parmColumn)=VT_I4;
    V_I4(&parmColumn) = 0;

    IDispatch *pdCell       = NULL;
    IDispatch *pdAnnotation = NULL;

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"getCellByPosition", 2, parmRow, parmColumn);
    if (hres != S_OK) {
        TRACE("Error when getCellByPosition\n");
        return hres;
    }
    pdCell = V_DISPATCH(&vRes);
    VariantInit(&vRes);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdCell, L"Annotation", 0);
    if (hres != S_OK) {
        TRACE("Annotation\n");
        IDispatch_Release(pdCell);
        return hres;
    }
    pdAnnotation = V_DISPATCH(&vRes);
    VariantInit(&vRes);

/*
    if (Text.vt == VT_ERROR)
    {
        hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdAnnotation, L"String", 0);
        pdCell->Release();
        pdAnnotation->Release();

        if (hres == S_OK)
        {

            int iLength;
            if (Length.vt == VT_ERROR)
            {
                iLength = 0xffff;
            }
            else
            {
                iLength = Length.intVal;
            }

            int iStart;
            if (Start.vt == VT_ERROR)
            {
                iStart = 0;
            }
            else
            {
                iStart = Start.intVal - 1;
            }

            int iRealLen    = ::SysStringLen(vRes.bstrVal);
            wchar_t *wscResult  = new wchar_t[iRealLen];
            wchar_t *wscBeg     = wscResult;
            memcpy(wscResult, vRes.bstrVal, iRealLen * sizeof(wchar_t));
            if ((iStart < iRealLen) && (iRealLen != 0))
            {

                wscBeg += iStart;
                if (iLength < iRealLen)
                {
                    *(wscBeg + iLength) = 0;
                } else
                {
                    *(wscBeg + iRealLen) = 0;
                }
            }
            else
            {
                wscBeg = 0;
            }
            *pText = ::SysAllocString(wscBeg);
        }
    } else
    {
        std::wstring sNewAnnotation = L"";
        if (Start.vt != VT_ERROR)
        {
            hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdAnnotation, L"String", 0);
            if (hres != S_OK)
            {
                pdCell->Release();
                pdAnnotation->Release();
                return hres;
            }
            sNewAnnotation = vRes.bstrVal;
        }
        sNewAnnotation = sNewAnnotation  + Text.bstrVal;
        CComVariant p1 = sNewAnnotation.c_str();
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, pdAnnotation, L"String", 1, p1);
        pdCell->Release();
        pdAnnotation->Release();
    }
    // TODO : Insert delete
        return hres;

*/
    TRACE_OUT;
    return E_NOTIMPL;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearContents(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT result,param,vRes;
    VariantInit(&result);
    VariantInit(&param);
    VariantInit(&vRes);
    V_VT(&result) = VT_NULL;
    HRESULT hres;
    TRACE_IN;

    if (This == NULL) return result;

    V_VT(&param) = VT_I4;
    V_I4(&param) = VALUE + DATETIME + STRING + FORMULA + OBJECTS;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"clearContents", 1, param);

    if (FAILED(hres)) {
        TRACE("ERROR when clearContents \n");
        return result;
    }

    TRACE_OUT;
    return result;
}

static long WINAPI MSO_TO_OO_I_Range_get_Column(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    IDispatch *pdRangeAddress = NULL;
    VARIANT vRes,vtmp;
    HRESULT hres;
    long lres;
    VariantInit(&vtmp);

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"getRangeAddress", 0);
    if (hres != S_OK) {
        TRACE("  OO->getRangeAddress FAILED \n");
        return -1;
    }

    pdRangeAddress = V_DISPATCH(&vRes);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdRangeAddress, L"StartColumn", 0);
    IDispatch_Release(pdRangeAddress);
    if (hres != S_OK) {
        TRACE(" OO->StartColumn FAILED \n");
        return -1;
    }

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_I4);
    if (FAILED(hres)) {
        TRACE(" ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    lres = V_I4(&vtmp);

    TRACE_OUT;
    return lres + 1;
}

static long WINAPI MSO_TO_OO_I_Range_get_Row(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    IDispatch *pdRangeAddress = NULL;
    VARIANT vRes,vtmp;
    HRESULT hres;
    long lres;
    VariantInit(&vtmp);

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"getRangeAddress", 0);
    if (hres != S_OK) {
       TRACE("  OO->getRangeAddress FAILED \n");
       return -1;
    }

    pdRangeAddress = V_DISPATCH(&vRes);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, pdRangeAddress, L"StartRow", 0);
    IDispatch_Release(pdRangeAddress);
    if (hres != S_OK)
    {
        TRACE(" OO->StartRow FAILED \n");
        return -1;
    }

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_I4);
    if (FAILED(hres)) {
        TRACE(" ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    lres = V_I4(&vtmp);

    TRACE_OUT;
    return lres + 1;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_HorizontalAlignment(
        I_Range* iface,
        XlHAlign *halign)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT vRes,vtmp;
    VariantInit(&vRes);
    long lret;
    VariantInit(&vtmp);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, This->pOORange, L"HoriJustify", 0);

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_I4);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    lret = V_I4(&vtmp);

    switch (lret) {
    case 0:*halign=xlHAlignGeneral ;break;
    case 1:*halign=xlHAlignLeft ;break;
    case 2:*halign=xlHAlignCenter ;break;
    case 3:*halign=xlHAlignRight ;break;
    case 4:*halign=xlHAlignJustify ;break;
    case 5:*halign=xlHAlignFill ;break;
    default:*halign=xlHAlignGeneral;break;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_HorizontalAlignment(
        I_Range* iface,
        XlHAlign halign)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes;
    VariantInit(&param);
    V_VT(&param) = VT_I4;
    switch (halign) {
    case xlHAlignCenter:V_I4(&param) = 2;break;
    case xlHAlignCenterAcrossSelection:V_I4(&param) = 0;break;
    case xlHAlignDistributed:V_I4(&param) = 0;break;
    case xlHAlignFill:V_I4(&param) = 5;break;
    case xlHAlignGeneral:V_I4(&param) = 0;break;
    case xlHAlignJustify:V_I4(&param) = 4;break;
    case xlHAlignLeft:V_I4(&param) = 1;break;
    case xlHAlignRight:V_I4(&param) = 3;break;
    }

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"HoriJustify", 1, param);
    if (hres != S_OK)  {
       return hres;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_VerticalAlignment(
        I_Range* iface,
        XlVAlign *valign)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT vRes,vtmp;
    VariantInit(&vRes);
    long lret;
    VariantInit(&vtmp);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, This->pOORange, L"VertJustify", 0);

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_I4);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    lret = V_I4(&vtmp);

    switch (lret) {
    case 0:*valign=xlVAlignJustify ;break;
    case 1:*valign=xlVAlignTop ;break;
    case 2:*valign=xlVAlignCenter ;break;
    case 3:*valign=xlVAlignBottom ;break;
    default:*valign=xlVAlignDistributed;break;
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_VerticalAlignment(
        I_Range* iface,
        XlVAlign valign)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes;
    VariantInit(&param);
    V_VT(&param) = VT_I4;
    switch (valign) {
    case xlVAlignBottom:V_I4(&param) = 3;break;
    case xlVAlignCenter:V_I4(&param) = 2;break;
    case xlVAlignDistributed:V_I4(&param) = 0;break;
    case xlVAlignJustify:V_I4(&param) = 0;break;
    case xlVAlignTop:V_I4(&param) = 1;break;
    }

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"VertJustify", 1, param);
    if (hres != S_OK)  {
       return hres;
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Merge(
        I_Range* iface,
        VARIANT_BOOL flag)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes, vraddr;
    HRESULT hres;
    VariantInit(&param);
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = VARIANT_TRUE;

    if (flag == VARIANT_FALSE) {
        hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"Merge", 1, param);
        TRACE_OUT;
        return hres;
    } else {
        /*надо пробежаться по всем рядам и объеденить их все отдельно*/
        long startrow,endrow,startcolumn,endcolumn;
        int i;

        hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);

        if (hres != S_OK)  {
            TRACE("Error when GetRangeAddress \n");
            return hres;
        }

        IDispatch *newrange;

        VARIANT vLeft, vRight, vTop, vBottom;

        for (i=0;i<=endrow-startrow;i++) {
            VariantInit(&vLeft);
            V_VT(&vLeft) = VT_I4;
            V_I4(&vLeft) = 0;
            VariantInit(&vTop);
            V_VT(&vTop) = VT_I4;
            V_I4(&vTop) = i;
            VariantInit(&vRight);
            V_VT(&vRight) = VT_I4;
            V_I4(&vRight) = endcolumn-startcolumn;
            VariantInit(&vBottom);
            V_VT(&vBottom) = VT_I4;
            V_I4(&vBottom) = i;

            hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"getCellRangeByPosition", 4, vBottom, vRight, vTop, vLeft);
            if (hres != S_OK)  {
                TRACE("Error when getCellRangeByPosition \n");
                return hres;
            }
            newrange = V_DISPATCH(&vRes);
            hres = AutoWrap(DISPATCH_METHOD, &vRes, newrange, L"Merge", 1, param);
            IDispatch_Release(newrange);
        }
    TRACE_OUT;
    return S_OK;
    }
}

static HRESULT WINAPI MSO_TO_OO_I_Range_UnMerge(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes;
    HRESULT hres;
    VariantInit(&param);
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = VARIANT_FALSE; /*т.к. мы разбиваем ячейки*/

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"Merge", 1, param);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_WrapText(
        I_Range* iface,
        VARIANT_BOOL *pvbwraptext)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT vRes;
    HRESULT hres;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, This->pOORange, L"isTextWrapped", 0);

    if (FAILED(hres)) {
        return hres;
    }
    *pvbwraptext = V_BOOL(&vRes);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_WrapText(
        I_Range* iface,
        VARIANT_BOOL pvbwraptext)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes;
    HRESULT hres;
    VariantInit(&param);
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = pvbwraptext;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"isTextWrapped", 1, param);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Application(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *value = This->pApplication;
    I_ApplicationExcel_AddRef(This->pApplication);

    if (value==NULL)
        return E_POINTER;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Parent(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *value = This->pwsheet;
    I_Worksheet_AddRef(*value);

    if (value==NULL)
        return E_POINTER;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Worksheet(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *value = This->pwsheet;
    I_Worksheet_AddRef(*value);

    if (value==NULL)
        return E_POINTER;

    TRACE_OUT;
    return S_OK;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_Clear(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    VARIANT result;
    VariantInit(&result);
    V_VT(&result) = VT_NULL;


    VARIANT param,vRes;
    VariantInit(&param);
    VariantInit(&vRes);
    HRESULT hres;

    if (This == NULL) return result;

    V_VT(&param) = VT_I4;
    V_I4(&param) = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + HARDATTR + STYLES + OBJECTS + EDITATTR + FORMATTED;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"clearContents", 1, param);

    if (hres != S_OK) {
        TRACE("ERROR when clearContents \n");
        return result;
    }

    TRACE_OUT;
    return result;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ClearComments(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    VARIANT param,vRes;
    VariantInit(&param);
    VariantInit(&vRes);
    HRESULT hres;

    if (This == NULL) return E_FAIL;

    V_VT(&param) = VT_I4;
    V_I4(&param) = ANNOTATION;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"clearContents", 1, param);

    if (hres != S_OK) {
        TRACE("ERROR when clearComments\n");
        return hres;
    }

    TRACE_OUT;
    return S_OK;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearFormats(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    VARIANT result;
    VariantInit(&result);
    V_VT(&result) = VT_NULL;

    VARIANT param,vRes;
    VariantInit(&param);
    VariantInit(&vRes);
    HRESULT hres;

    if (This == NULL) return result;

    V_VT(&param) = VT_I4;
    V_I4(&param) = HARDATTR + STYLES + EDITATTR + FORMATTED;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"clearContents", 1, param);

    if (hres != S_OK) {
        TRACE("ERROR when clearFormats \n");
        return result;
    }

    TRACE_OUT;
    return result;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearNotes(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    VARIANT result;
    VariantInit(&result);
    V_VT(&result) = VT_NULL;

    VARIANT param,vRes;
    VariantInit(&param);
    VariantInit(&vRes);
    HRESULT hres;

    if (This == NULL) return result;

    V_VT(&param) = VT_I4;
    V_I4(&param) = ANNOTATION;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"clearContents", 1, param);

    if (hres != S_OK) {
        TRACE("ERROR when clearContents \n");
        return result;
    }

    TRACE_OUT;
    return result;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearOutline(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    VARIANT result;
    VariantInit(&result);
    V_VT(&result) = VT_NULL;

    VARIANT param,vRes;
    VariantInit(&param);
    VariantInit(&vRes);
    HRESULT hres;

    if (This == NULL) return result;

    V_VT(&param) = VT_I4;
    V_I4(&param) = STYLES;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"clearContents", 1, param);

    if (hres != S_OK) {
        TRACE("ERROR when clearOutline \n");
        return result;
    }

    TRACE_OUT;
    return result;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Interior(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;

    HRESULT hr;
    IUnknown *punk = NULL;
    IDispatch *pInterior;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *value = NULL;

    hr = _I_InteriorConstructor((LPVOID*) &punk);

    if (FAILED(hr)) return E_NOINTERFACE;

    hr = I_Interior_QueryInterface(punk, &IID_I_Interior, (void**) &pInterior);
    if (pInterior == NULL) {
        return E_FAIL;
    }
    /*проинициализируем его*/
    hr = MSO_TO_OO_I_Interior_Initialize((I_Interior*)pInterior, iface);

    if (FAILED(hr)) {
        IDispatch_Release(pInterior);
        return hr;
    }

    *value = pInterior;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Borders(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;

    HRESULT hr;
    IUnknown *punk = NULL;
    IDispatch *pBorders;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    *value = NULL;

    hr = _I_BordersConstructor((LPVOID*) &punk);

    if (FAILED(hr)) return E_NOINTERFACE;

    hr = I_Borders_QueryInterface(punk, &IID_I_Borders, (void**) &pBorders);
    if (pBorders == NULL) {
        return E_FAIL;
    }
    /*проинициализируем его*/

    hr = MSO_TO_OO_I_Borders_Initialize((I_Borders*)pBorders, iface);

    if (FAILED(hr)) {
        IDispatch_Release(pBorders);
        return hr;
    }

    *value = pBorders;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Count(
        I_Range* iface,
        long *lcount)
{
    HRESULT hres;
    long startrow, startcolumn, endrow, endcolumn,w,h;
    TRACE_IN;

    hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress\n");
        return hres;
    }
    w = endcolumn - startcolumn + 1;
    h = endrow - startrow + 1;

    *lcount = w*h;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Delete(
        I_Range* iface,
        VARIANT param,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    XlDeleteShiftDirection action;
    HRESULT hres;
    long startrow, startcolumn, endrow, endcolumn;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    if (This == NULL) return E_POINTER;
    *value = NULL;

    hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress\n");
        return hres;
    }

    if (Is_Variant_Null(param)) {
        if ((endcolumn-startcolumn)>(endrow-startrow)) action = xlShiftUp; else action = xlShiftToLeft;
    } else {
        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&param, &param, 0, 0, VT_I4);
        if (FAILED(hres)) {
            TRACE("ERROR VariantChangeTypeEx VT=%i \n", V_VT(&param));
        }
        switch(V_I4(&param)) {
        case xlShiftToLeft: action = xlShiftToLeft;break;
        case xlShiftUp: action =xlShiftUp;break;
        default: action = xlShiftToLeft;break;
        }
    }

    /*Теперь в зависимости от того куда надо сдвигать
    получаем или ряды или колонки*/
    VARIANT tmp_range, vRes, par1, par2;
    VariantInit(&tmp_range);
    VariantInit(&vRes);
    VariantInit(&par1);
    VariantInit(&par2);

    switch (action) {
    case xlShiftToLeft:
        hres = AutoWrap(DISPATCH_METHOD, &tmp_range, This->pOORange, L"getColumns", 0);
        if (hres != S_OK) {
            TRACE("ERROR when getColumns\n");
            return hres;
        }
        V_VT(&par1) = VT_I4;
        V_I4(&par1) = 0;
        V_VT(&par2) = VT_I4;
        V_I4(&par2) = endcolumn - startcolumn + 1;
        TRACE("Delete Columns: index = %i    count = %i \n", V_I4(&par1), V_I4(&par2));
        TRACE("startrow=%i, startcolumn=%i, endrow=%i, endcolumn=%i \n",startrow, startcolumn, endrow, endcolumn);
        hres = AutoWrap(DISPATCH_METHOD, &vRes, V_DISPATCH(&tmp_range), L"removeByIndex", 2, par2, par1);
        if (hres != S_OK) {
            TRACE("ERROR when removeByIndex\n");
            return hres;
        }
        *value = (IDispatch*)iface;
        break;
    case xlShiftUp:
        hres = AutoWrap(DISPATCH_METHOD, &tmp_range, This->pOORange, L"getRows", 0);
        if (hres != S_OK) {
            TRACE("ERROR when getRows\n");
            return hres;
        }
        V_VT(&par1) = VT_I4;
        V_I4(&par1) = 0;
        V_VT(&par2) = VT_I4;
        V_I4(&par2) = endrow - startrow + 1;
        TRACE("Delete Rows: index = %i    count = %i \n", V_I4(&par1), V_I4(&par2));
        TRACE("startrow=%i, startcolumn=%i, endrow=%i, endcolumn=%i \n",startrow, startcolumn, endrow, endcolumn);
        hres = AutoWrap(DISPATCH_METHOD, &vRes, V_DISPATCH(&tmp_range), L"removeByIndex", 2, par2, par1);
        if (hres != S_OK) {
            TRACE("ERROR when removeByIndex\n");
            return hres;
        }
        *value = (IDispatch*)iface;
        break;
    }
    I_Range_AddRef(*value);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_RowHeight(
        I_Range* iface,
        long *pnrowheight)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    if (This->pOORange == NULL)
        return E_POINTER;

    VARIANT rows;
    HRESULT hres;
    VARIANT vRowHeight;

    hres = AutoWrap(DISPATCH_METHOD, &rows, This->pOORange, L"getRows", 0);
    if (hres != S_OK)
        return hres;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRowHeight, V_DISPATCH(&rows), L"Height", 0);

    IDispatch_Release(V_DISPATCH(&rows));
    *pnrowheight = V_I2(&vRowHeight)/100;

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_RowHeight(
        I_Range* iface,
        long nrowheight)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    if (This->pOORange == NULL)
        return E_POINTER;

    VARIANT res;
    VARIANT rows;
    HRESULT hres;
    VARIANT vRowHeight;

    hres = AutoWrap(DISPATCH_METHOD, &rows, This->pOORange, L"getRows", 0);
    if (hres != S_OK)
        return hres;

    VariantInit(&vRowHeight);
    V_VT(&vRowHeight) = VT_I4;
    V_I4(&vRowHeight) = nrowheight*100;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, V_DISPATCH(&rows), L"Height", 1, vRowHeight);
    IDispatch_Release(V_DISPATCH(&rows));

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Copy(
        I_Range* iface,
        VARIANT RangeTo,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT tmp_var;
    VariantInit(&tmp_var);
    HRESULT hres;
    SAFEARRAY FAR* pPropVals;
    BSTR command;
    VARIANT param;
    WorksheetImpl *wsh = (WorksheetImpl*)This->pwsheet;
    WorkbookImpl *parent_wb = (WorkbookImpl*)wsh->pwb;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(RangeTo, &RangeTo);

    if (This == NULL) return E_POINTER;
    if (This->pOORange == NULL)
        return E_POINTER;

    hres = MSO_TO_OO_I_Range_Select(iface, &tmp_var);
    if (FAILED(hres)) {
        TRACE("ERROR Select\n");
        return E_FAIL;
    }

    command = SysAllocString(L".uno:Copy");

    pPropVals = SafeArrayCreateVector( VT_DISPATCH, 0, 0);
    VariantInit (&param);
    V_VT(&param) = VT_DISPATCH | VT_ARRAY;
    V_ARRAY(&param) = pPropVals;
    hres = MSO_TO_OO_ExecuteDispatchHelper_WB((I_Workbook*)parent_wb, command, param);
    if (FAILED(hres)){
        TRACE("ERROR Dispatcher (.uno:Copy)\n");
        return hres;
    }

    if (Is_Variant_Null(RangeTo)) {
        TRACE("(To Clipboard)\n");
        *value = (IDispatch*)iface;
        IDispatch_AddRef(*value);
        TRACE_OUT;
        return S_OK;
    } else {
        TRACE("(To another range)\n");
        SysAllocString(command);
       command = SysAllocString(L".uno:Paste");

        if (V_VT(&RangeTo)!=VT_DISPATCH) {
            TRACE("ERROR parameter\n");
            return E_FAIL;
        }

        hres = MSO_TO_OO_I_Range_Select((I_Range*)(V_DISPATCH(&RangeTo)),&tmp_var);
        if (FAILED(hres)) {
            TRACE("ERROR Select\n");
            return hres;
        }
        IDispatch *irange_tmp;

        RangeImpl *range_tmp = (RangeImpl*)(V_DISPATCH(&RangeTo));
        WorksheetImpl *wsh2 = (WorksheetImpl*)(range_tmp->pwsheet);
        WorkbookImpl *parent_wb2 = (WorkbookImpl*)wsh2->pwb;

        hres = MSO_TO_OO_ExecuteDispatchHelper_WB((I_Workbook*)parent_wb2, command, param);
        if (FAILED(hres)){
            TRACE("ERROR Dispatcher (.uno:Paste)\n");
            return hres;
        }

        *value = (V_DISPATCH(&RangeTo));
        IDispatch_AddRef(*value);
        TRACE_OUT;
        return S_OK;
    }

    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_NumberFormat(
        I_Range* iface,
        VARIANT *pnumbformat)
{
/*TODO*/
    TRACE_IN;
    V_VT(pnumbformat) = VT_BSTR;
    V_BSTR(pnumbformat) = SysAllocString(L"");
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_NumberFormat(
        I_Range* iface,
        VARIANT numbformat)
{
/*TODO*/
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_NumberFormatLocal(
        I_Range* iface,
        VARIANT *pnumbformat)
{
/*TODO*/
    TRACE_IN;
    V_VT(pnumbformat) = VT_BSTR;
    V_BSTR(pnumbformat) = SysAllocString(L"");
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_NumberFormatLocal(
        I_Range* iface,
        VARIANT numbformat)
{
/*TODO*/
    TRACE_NOTIMPL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Height(
        I_Range* iface,
        VARIANT *value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    VARIANT vsize;
    TRACE_IN;

    VariantInit(&vsize);

    if (value==NULL) return E_FAIL;

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vsize, This->pOORange, L"Size",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Size \n");
        return hres;
    }

    /*Подумать над 1/100 мм */

    hres = AutoWrap(DISPATCH_PROPERTYGET, value, V_DISPATCH(&vsize), L"Height",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Height \n");
        return hres;
    }

    VariantClear(&vsize);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Width(
        I_Range* iface,
        VARIANT *value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    VARIANT vsize;
    TRACE_IN;

    VariantInit(&vsize);

    if (value==NULL) return E_FAIL;

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vsize, This->pOORange, L"Size",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Size \n");
        return hres;
    }

    /*Подумать над 1/100 мм */

    hres = AutoWrap(DISPATCH_PROPERTYGET, value, V_DISPATCH(&vsize), L"Width",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Width \n");
        return hres;
    }

    VariantClear(&vsize);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Left(
        I_Range* iface,
        VARIANT *value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    VARIANT vsize;
    TRACE_IN;

    VariantInit(&vsize);

    if (value==NULL) return E_FAIL;

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vsize, This->pOORange, L"Position",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Position \n");
        return hres;
    }

    /*Подумать над 1/100 мм */

    hres = AutoWrap(DISPATCH_PROPERTYGET, value, V_DISPATCH(&vsize), L"X",0);
    if (FAILED(hres)) {
        TRACE("ERROR when X \n");
        return hres;
    }

    VariantClear(&vsize);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Top(
        I_Range* iface,
        VARIANT *value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    VARIANT vsize;

    VariantInit(&vsize);
    TRACE_IN;

    if (value==NULL) return E_FAIL;

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vsize, This->pOORange, L"Position",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Position \n");
        return hres;
    }

    /*Подумать над 1/100 мм */

    hres = AutoWrap(DISPATCH_PROPERTYGET, value, V_DISPATCH(&vsize), L"Y",0);
    if (FAILED(hres)) {
        TRACE("ERROR when Y \n");
        return hres;
    }

    VariantClear(&vsize);

    TRACE_OUT;
    return S_OK;
}


static HRESULT WINAPI MSO_TO_OO_I_Range_get_ShrinkToFit(
        I_Range* iface,
        VARIANT *pparam)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes,vtmp;
    VariantInit(&vRes);
    VariantInit(&vtmp);
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, This->pOORange, L"ShrinkToFit", 0);

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    *pparam = vtmp;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_ShrinkToFit(
        I_Range* iface,
        VARIANT param)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes;
    HRESULT hres;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    if (This == NULL) return E_POINTER;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"ShrinkToFit", 1, param);
    if (FAILED(hres))  {
       TRACE("ERROR when ShrinkToFit \n");
       return hres;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_MergeCells(
        I_Range* iface,
        VARIANT *pparam)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes,vtmp;
    VariantInit(&vRes);
    VariantInit(&vtmp);
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"getIsMerged", 0);

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    *pparam = vtmp;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_MergeCells(
        I_Range* iface,
        VARIANT param)
{
    HRESULT hres;
    VARIANT vtmp;
    TRACE_IN;

    VariantInit(&vtmp);

    MSO_TO_OO_CorrectArg(param, &param);

    hres = VariantChangeTypeEx(&vtmp, &param, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
    }

    if (V_BOOL(&vtmp)==VARIANT_FALSE) {
        hres = MSO_TO_OO_I_Range_UnMerge(iface);
    } else {
        hres = MSO_TO_OO_I_Range_Merge(iface, VARIANT_TRUE);
    }

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Locked(
        I_Range* iface,
        VARIANT *pparam)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes,vtmp, vCellProt;
    HRESULT hres;
    VariantInit(&vRes);
    VariantInit(&vCellProt);
    VariantInit(&vtmp);
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vCellProt, This->pOORange, L"CellProtection", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get CellProtection \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, V_DISPATCH(&vCellProt), L"IsLocked", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get IsLocked \n");
        return E_FAIL;
    }

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    *pparam = vtmp;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Locked(
        I_Range* iface,
        VARIANT param)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes, vCellProt, vtmp;
    HRESULT hres;
    VariantInit(&vRes);
    VariantInit(&vCellProt);
    VariantInit(&vtmp);
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    if (This == NULL) return E_POINTER;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vCellProt, This->pOORange, L"CellProtection", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get CellProtection \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, V_DISPATCH(&vCellProt), L"IsLocked", 1, param);
    if (FAILED(hres))  {
       TRACE("ERROR when IsLocked \n");
       return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"CellProtection", 1, vCellProt);
    if (FAILED(hres))  {
       TRACE("ERROR when CellProtection \n");
       return hres;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Hidden(
        I_Range* iface,
        VARIANT *pparam)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes,vtmp, vCellProt;
    HRESULT hres;
    VariantInit(&vRes);
    VariantInit(&vCellProt);
    VariantInit(&vtmp);
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vCellProt, This->pOORange, L"CellProtection", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get CellProtection \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, V_DISPATCH(&vCellProt), L"IsHidden", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get IsHidden \n");
        return E_FAIL;
    }

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    *pparam = vtmp;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Hidden(
        I_Range* iface,
        VARIANT param)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes, vCellProt, vtmp;
    HRESULT hres;
    VariantInit(&vRes);
    VariantInit(&vCellProt);
    VariantInit(&vtmp);
    TRACE_IN;

    MSO_TO_OO_CorrectArg(param, &param);

    if (This == NULL) return E_POINTER;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vCellProt, This->pOORange, L"CellProtection", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get CellProtection \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, V_DISPATCH(&vCellProt), L"IsHidden", 1, param);
    if (FAILED(hres))  {
       TRACE("ERROR when IsHidden \n");
       return hres;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"CellProtection", 1, vCellProt);
    if (FAILED(hres))  {
       TRACE("ERROR when CellProtection \n");
       return hres;
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_MergeArea(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    long left, right,top,bottom;
    TRACE_IN;

    hres = MSO_TO_OO_GetRangeAddress(iface, &left, &top, &right, &bottom);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress\n");
    }
    /*Если не 1 ячейка, значит ошибка*/
    if ((left!=right)||(top!=bottom)) return E_FAIL;
    /*Получить объединенную область или ячейку*/

    *value = (IDispatch*)This;
    IDispatch_AddRef(*value);

    TRACE_OUT;
    return S_OK;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_AutoFit(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT result;
    VARIANT res;
    VARIANT range;
    VARIANT param;
    HRESULT hres;
    TRACE_IN;

    VariantInit(&result);
    VariantInit(&res);
    VariantInit(&range);
    VariantInit(&param);

    V_VT(&result) = VT_NULL;

    if (This == NULL) return result;

    hres = AutoWrap(DISPATCH_METHOD, &range, This->pOORange, L"getColumns", 0);
    if (hres != S_OK) {
        TRACE("Error when getColumns\n");
        return result;
    }

    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = VARIANT_TRUE;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, V_DISPATCH(&range), L"OptimalWidth", 1, param);
    if (FAILED(hres)) TRACE("ERROR when OptimalWidth\n");
    IDispatch_Release(V_DISPATCH(&range));

    VariantClear(&res);

    hres = AutoWrap(DISPATCH_METHOD, &range, This->pOORange, L"getRows", 0);
    if (hres != S_OK) {
        TRACE("Error when getRows\n");
        return result;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, V_DISPATCH(&range), L"OptimalHeight", 1, param);
    if (FAILED(hres)) TRACE("ERROR when OptimalHeight\n");
    IDispatch_Release(V_DISPATCH(&range));

    VariantClear(&res);
    VariantClear(&param);

    TRACE_OUT;
    return result;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Insert(
        I_Range* iface,
        VARIANT Shift,
        VARIANT CopyOrigin,
        VARIANT* RHS)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT result, shift, param1, param2, res;
    HRESULT hres;
    CellInsertMode insert_mode;
    WorksheetImpl* wsh = (WorksheetImpl*)(This->pwsheet);
    long startrow=0, startcolumn=0, endrow=0, endcolumn=0;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(Shift, &Shift);
    MSO_TO_OO_CorrectArg(CopyOrigin, &CopyOrigin);

    VariantInit(&result);
    VariantInit(&shift);
    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&res);
    V_VT(&result) = VT_NULL;

    //CopyOrigin is ignore now

    if (Is_Variant_Null(Shift)) {
        hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);
        if (FAILED(hres)) {
            TRACE("ERROR when GetRangeAddress\n");
        }
        if ((endcolumn - startcolumn)>(endrow - startrow)) insert_mode = RIGHT; else insert_mode = DOWN;
    } else {
        hres = VariantChangeTypeEx(&shift, &Shift, 0, 0, VT_I4);
        if (FAILED(hres)) {
           TRACE("ERROR when VariantChangeTypeEx\n");
            VariantClear(&shift);
            return E_FAIL;
        }

        switch (V_I4(&shift)) {
        case -4121:
            insert_mode = DOWN;
            break;
        case -4161:
            insert_mode = RIGHT;
            break;
        default:
            TRACE("ERROR invalid argument Shift = %i", V_I4(&shift));
            VariantClear(&shift);
            return E_INVALIDARG;
        }
    }

    hres = AutoWrap(DISPATCH_METHOD, &param1, This->pOORange, L"getRangeAddress", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when getRangeAddress \n");
        VariantClear(&shift);
        return E_FAIL;
    }

    V_VT(&param2) = VT_I4;
    V_I4(&param2) = insert_mode;
    hres = AutoWrap(DISPATCH_METHOD, &res, wsh->pOOSheet, L"insertCells", 2,param2, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when insertCells \n");
    }

    VariantClear(&shift);
    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&res);
    *RHS = result;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_EntireColumn(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    IUnknown *punk;
    VARIANT range, vcount, vstr, param1, vcolumn, vname1, vname2;
    HRESULT hres;
    WCHAR str[10];
    long start,end;
    TRACE_IN;

    VariantInit(&range);
    VariantInit(&vcount);
    VariantInit(&vstr);
    VariantInit(&param1);
    VariantInit(&vcolumn);
    VariantInit(&vname1);
    VariantInit(&vname2);

    if (This == NULL) {
        TRACE("ERROR object is NULL\n");
        return E_POINTER;
    }

    hres = AutoWrap(DISPATCH_METHOD, &range, This->pOORange, L"getColumns", 0);
    if (hres != S_OK) {
        TRACE("Error when getColumns\n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vcount, V_DISPATCH(&range), L"getCount", 0);
    if (hres != S_OK) {
        TRACE("Error when getCount\n");
        return hres;
    }
    start = 0;
    end = V_I4(&vcount) - 1;

    /*получаем имена столбцов*/
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = start;
    hres = AutoWrap(DISPATCH_METHOD, &vcolumn, V_DISPATCH(&range), L"getByIndex", 1, param1);
    if (hres != S_OK) {
        TRACE("Error when getByIndex\n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vname1, V_DISPATCH(&vcolumn), L"getName", 0);
    if (hres != S_OK) {
        TRACE("Error when getName\n");
        return hres;
    }

    VariantClear(&vcolumn);
    VariantClear(&param1);

    V_VT(&param1) = VT_I4;
    V_I4(&param1) = end;
    hres = AutoWrap(DISPATCH_METHOD, &vcolumn, V_DISPATCH(&range), L"getByIndex", 1, param1);
    if (hres != S_OK) {
        TRACE("Error when getByIndex\n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vname2, V_DISPATCH(&vcolumn), L"getName", 0);
    if (hres != S_OK) {
        TRACE("Error when getName\n");
        return hres;
    }
    /*соединяем в одну строку*/
    wsprintfW(str, L"%s:%s", V_BSTR(&vname1), V_BSTR(&vname2));

    V_VT(&vstr) = VT_BSTR;
    V_BSTR(&vstr) = SysAllocString(str);
    hres = I_Worksheet_get_Columns((I_Worksheet*)(This->pwsheet), vstr, value);
    if (FAILED(hres)) {
        TRACE("ERROR when initialize Range\n");
        return hres;
    }

    VariantClear(&range);
    VariantClear(&vcount);
    VariantClear(&vstr);
    VariantClear(&param1);
    VariantClear(&vcolumn);
    VariantClear(&vname1);
    VariantClear(&vname2);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_EntireRow(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    IUnknown *punk;
    VARIANT range, vcount, vstr, param1, vrow, vname1, vname2;
    HRESULT hres;
    WCHAR str[10];
    long start,end;
    long startrow=0, startcolumn=0, endrow=0, endcolumn=0;
    TRACE_IN;

    VariantInit(&range);
    VariantInit(&vcount);
    VariantInit(&vstr);
    VariantInit(&param1);
    VariantInit(&vrow);
    VariantInit(&vname1);
    VariantInit(&vname2);

    if (This == NULL) {
        TRACE("ERROR object is NULL\n");
        return E_POINTER;
    }

    hres = AutoWrap(DISPATCH_METHOD, &range, This->pOORange, L"getRows", 0);
    if (hres != S_OK) {
        TRACE("Error when getRows\n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, &vcount, V_DISPATCH(&range), L"getCount", 0);
    if (hres != S_OK) {
        TRACE("Error when getCount\n");
        return hres;
    }
    start = 0;
    end = V_I4(&vcount) - 1;
    TRACE("start=%i end=%i \n", start, end);

    hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress\n");
    }
    TRACE("%i    %i    %i    %i \n", startrow, startcolumn, endrow, endcolumn);

    /*соединяем в одну строку*/
    wsprintfW(str, L"%i:%i", startrow+1, endrow+1);

    V_VT(&vstr) = VT_BSTR;
    V_BSTR(&vstr) = SysAllocString(str);
    hres = I_Worksheet_get_Rows((I_Worksheet*)(This->pwsheet), vstr, value);
    if (FAILED(hres)) {
        TRACE("ERROR when initialize Range\n");
        return hres;
    }

    VariantClear(&range);
    VariantClear(&vcount);
    VariantClear(&vstr);
    VariantClear(&param1);
    VariantClear(&vrow);
    VariantClear(&vname1);
    VariantClear(&vname2);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Creator(
        I_Range* iface,
        VARIANT *result)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Activate(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_AddIndent(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_AddIndent(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Address(
        I_Range* iface,
        VARIANT RowAbsolute,
        VARIANT ColumnAbsolute,
        XlReferenceStyle ReferenceStyle,
        VARIANT External,
        VARIANT RelativeTo,
        LCID lcid,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_AddressLocal(
        I_Range* iface,
        VARIANT RowAbsolute,
        VARIANT ColumnAbsolute,
        XlReferenceStyle ReferenceStyle,
        VARIANT External,
        VARIANT RelativeTo,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_AdvancedFilter(
        I_Range* iface,
        XlFilterAction Action,
        VARIANT CriteriaRange,
        VARIANT CopyToRange,
        VARIANT Unique,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ApplyNames(
        I_Range* iface,
        VARIANT Names,
        VARIANT IgnoreRelativeAbsolute,
        VARIANT UseRowColumnNames,
        VARIANT OmitColumn,
        VARIANT OmitRow,
        XlApplyNamesOrder Order,
        VARIANT AppendLast,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ApplyOutlineStyles(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Areas(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_AutoComplete(
        I_Range* iface,
        BSTR String,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_AutoFill(
        I_Range* iface,
        IDispatch *Destination,
        XlAutoFillType Type,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_AutoFilter(
        I_Range* iface,
        VARIANT Field,
        VARIANT Criteria1,
        XlAutoFilterOperator Operator,
        VARIANT Criteria2,
        VARIANT VisibleDropDown,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_AutoFormat(
        I_Range* iface,
        XlRangeAutoFormat Format,
        VARIANT Number,
        VARIANT IXLFont,
        VARIANT Alignment,
        VARIANT Border,
        VARIANT Pattern,
        VARIANT Width,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_AutoOutline(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_BorderAround(
        I_Range* iface,
        VARIANT LineStyle,
        XlBorderWeight Weight,
        XlColorIndex ColorIndex,
        VARIANT Color,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Calculate(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Cells(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_IN;

    *RHS = (IDispatch*)iface;
    I_Range_AddRef(*RHS);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Characters(
        I_Range* iface,
        VARIANT Start,
        VARIANT Length,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_CheckSpelling(
        I_Range* iface,
        VARIANT CustomDictionary,
        VARIANT IgnoreUppercase,
        VARIANT AlwaysSuggest,
        VARIANT SpellLang,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ColumnDifferences(
        I_Range* iface,
        VARIANT Comparison,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Columns(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return I_Range_get_EntireColumn(iface, RHS);;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Consolidate(
        I_Range* iface,
        VARIANT Sources,
        VARIANT Function,
        VARIANT TopRow,
        VARIANT LeftColumn,
        VARIANT CreateLinks,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_CopyFromRecordset(
        I_Range* iface,
        IUnknown *Data,
        VARIANT MaxRows,
        VARIANT MaxColumns,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_CopyPicture(
        I_Range* iface,
        XlPictureAppearance Appearance,
        XlCopyPictureFormat Format,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_CreateNames(
        I_Range* iface,
        VARIANT Top,
        VARIANT Left,
        VARIANT Bottom,
        VARIANT Right,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_CreatePublisher(
        I_Range* iface,
        VARIANT Edition,
        XlPictureAppearance Appearance,
        VARIANT ContainsPICT,
        VARIANT ContainsBIFF,
        VARIANT ContainsRTF,
        VARIANT ContainsVALU,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_CurrentArray(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_CurrentRegion(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Cut(
        I_Range* iface,
        VARIANT Destination,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_DataSeries(
        I_Range* iface,
        VARIANT Rowcol,
        XlDataSeriesType Type,
        XlDataSeriesDate Date,
        VARIANT Step,
        VARIANT Stop,
        VARIANT Trend,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put__Default(
        I_Range* iface,
        VARIANT varRowIndex,
        VARIANT varColumnIndex,
        LCID lcid,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Dependents(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_DialogBox1(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_DirectDependents(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_EditionOptions(
        I_Range* iface,
        XlEditionType Type,
        XlEditionOptionsOption Option,
        VARIANT Name,
        VARIANT Reference,
        XlPictureAppearance Appearance,
        XlPictureAppearance ChartSize,
        VARIANT Format,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_End(
        I_Range* iface,
        XlDirection Direction,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_FillDown(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_FillLeft(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_FillRight(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_FillUp(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Find(
        I_Range* iface,
        VARIANT What,
        VARIANT After,
        VARIANT LookIn,
        VARIANT LookAt,
        VARIANT SearchOrder,
        XlSearchDirection SearchDirection,
        VARIANT MatchCase,
        VARIANT MatchByte,
        VARIANT SearchFormat,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_FindNext(
        I_Range* iface,
        VARIANT After,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_FindPrevious(
        I_Range* iface,
        VARIANT After,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Formula(
        I_Range* iface,
        LCID lcid,
        VARIANT *RHS)
{
    RangeImpl *This = (RangeImpl*)iface;
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    VARIANT resultCell;
    HRESULT hres;
    VARIANT parmRow; // XPOS

    V_VT(&parmRow) = VT_I2;
    V_I2(&parmRow) = 0;
    VARIANT parmColumn; // nYPos
    V_VT(&parmColumn) = VT_I2;
    V_I2(&parmColumn) = 0;

    hres = AutoWrap(DISPATCH_METHOD, &resultCell, This->pOORange, L"getCellByPosition", 2, parmRow, parmColumn);
    if (hres != S_OK) {
        TRACE("ERROR when getCellByPosition \n");
        return hres;
    }

    hres = AutoWrap(DISPATCH_METHOD, RHS, V_DISPATCH(&resultCell), L"getFormula", 0);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Formula(
        I_Range* iface,
        LCID lcid,
        VARIANT RHS)
{
    VARIANT vNull;
    TRACE_IN;
    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;
    TRACE_OUT;
    return I_Range_put_Value(iface, vNull, 0, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_FormulaArray(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_FormulaArray(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_FormulaLabel(
        I_Range* iface,
        XlFormulaLabel *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_FormulaLabel(
        I_Range* iface,
        XlFormulaLabel RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_FormulaHidden(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_FormulaHidden(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_FormulaLocal(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_FormulaLocal(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_FormulaR1C1(
        I_Range* iface,
        LCID lcid,
        VARIANT *RHS)
{
    VARIANT tmp;
    TRACE_IN;
    VariantInit(&tmp);
    TRACE_OUT;
    return I_Range_get_Value(iface, tmp, lcid, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_FormulaR1C1(
        I_Range* iface,
        LCID lcid,
        VARIANT RHS)
{
    /* .uno:SheetUseR1C1 */
    VARIANT tmp;
    TRACE_IN;
    VariantInit(&tmp);
    TRACE_OUT;
    return I_Range_put_Value(iface, tmp, lcid, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_FormulaR1C1Local(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_FormulaR1C1Local(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_FunctionWizard(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_GoalSeek(
        I_Range* iface,
        VARIANT Goal,
        IDispatch *ChangingCell,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Group(
        I_Range* iface,
        VARIANT Start,
        VARIANT End,
        VARIANT By,
        VARIANT Periods,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_HasArray(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_HasFormula(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_IndentLevel(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_IndentLevel(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_InsertIndent(
        I_Range* iface,
        long InsertAmount)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Item(
        I_Range* iface,
        VARIANT RowIndex,
        VARIANT ColumnIndex,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Item(
        I_Range* iface,
        VARIANT RowIndex,
        VARIANT ColumnIndex,
        LCID lcid,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Justify(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_ListHeaderRows(
        I_Range* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ListNames(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_LocationInTable(
        I_Range* iface,
        XlLocationInTable *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Name(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Name(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_NavigateArrow(
        I_Range* iface,
        VARIANT TowardPrecedent,
        VARIANT ArrowNumber,
        VARIANT LinkNumber,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get__NewEnum(
        I_Range* iface,
        IUnknown **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Next(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Offset(
        I_Range* iface,
        VARIANT RowOffset,
        VARIANT ColumnOffset,
        IDispatch **RHS)
{
    RangeImpl* This = (RangeImpl*)iface;
    WorksheetImpl* wsh = (WorksheetImpl*)This->pwsheet;
    HRESULT hres;
    long left, top, right, bottom;
    long drow = 0, dcol = 0;
    struct CELL_COORD lefttop, rightbottom;
    IDispatch *pCell;
    IUnknown *punk;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(RowOffset, &RowOffset);
    MSO_TO_OO_CorrectArg(ColumnOffset, &ColumnOffset);

    if (!Is_Variant_Null(RowOffset)) {
         hres = VariantChangeTypeEx(&RowOffset, &RowOffset, 0, 0, VT_I4);
         if (FAILED(hres)) {
             TRACE("ERROR when VariantChangeTypeEx\n");
         }
         drow = V_I4(&RowOffset);
    }

    if (!Is_Variant_Null(ColumnOffset)) {
         hres = VariantChangeTypeEx(&ColumnOffset, &ColumnOffset, 0, 0, VT_I4);
         if (FAILED(hres)) {
             TRACE("ERROR when VariantChangeTypeEx\n");
         }
         dcol = V_I4(&ColumnOffset);
    }

    hres = MSO_TO_OO_GetRangeAddress(iface, &left, &top, &right, &bottom);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress \n");
        return E_FAIL;
    }
    TRACE("drow = %i , dcol = %i \n", drow, dcol);
    TRACE("StartRow = %i, StartColumn = %i, EndRow = %i, EndColumn = %i \n", left, top, right, bottom);
    /*В OpenOffice оси направлены наоборот*/
    left += drow;
    right += drow;
    top += dcol;
    bottom += dcol;
    TRACE("StartRow = %i, StartColumn = %i, EndRow = %i, EndColumn = %i \n", left, top, right, bottom);
    /*создаем новый */
    lefttop.x = top+1;
    lefttop.y = left+1;
    rightbottom.x = bottom+1;
    rightbottom.y = right+1;


    hres = _I_RangeConstructor((LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pCell);

    if (pCell == NULL) {
        TRACE("ERROR when QueryInterface\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Range_Initialize((I_Range*)pCell, (I_Range*)wsh->pAllRange, lefttop, rightbottom);
        if (FAILED(hres)){
            TRACE("ERROR when initialize\n");
            I_Range_Release(pCell);
            return hres;
        }
    *RHS = pCell;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Orientation(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Orientation(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_OutlineLevel(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_OutlineLevel(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_PageBreak(
        I_Range* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_PageBreak(
        I_Range* iface,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Parse(
        I_Range* iface,
        VARIANT ParseLine,
        VARIANT Destination,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range__PasteSpecial(
        I_Range* iface,
        XlPasteType Paste,
        XlPasteSpecialOperation Operation,
        VARIANT SkipBlanks,
        VARIANT Transpose,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_PivotField(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_PivotItem(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_PivotTable(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Precedents(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_PrefixCharacter(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Previous(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range__PrintOut(
        I_Range* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_PrintPreview(
        I_Range* iface,
        VARIANT EnableChanges,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_QueryTable(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Range(
        I_Range* iface,
        VARIANT Cell1,
        VARIANT Cell2,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_RemoveSubtotal(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Replace(
        I_Range* iface,
        VARIANT What,
        VARIANT Replacement,
        VARIANT LookAt,
        VARIANT SearchOrder,
        VARIANT MatchCase,
        VARIANT MatchByte,
        VARIANT SearchFormat,
        VARIANT ReplaceFormat,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Resize(
        I_Range* iface,
        VARIANT RowSize,
        VARIANT ColumnSize,
        IDispatch **RHS)
{
    RangeImpl* This = (RangeImpl*)iface;
    WorksheetImpl* wsh = (WorksheetImpl*)This->pwsheet;
    HRESULT hres;
    long left, top, right, bottom;
    long drow = 0, dcol = 0;
    struct CELL_COORD lefttop, rightbottom;
    IDispatch *pCell;
    IUnknown *punk;
    TRACE_IN;

    MSO_TO_OO_CorrectArg(RowSize, &RowSize);
    MSO_TO_OO_CorrectArg(ColumnSize, &ColumnSize);

    if (!Is_Variant_Null(RowSize)) {
         hres = VariantChangeTypeEx(&RowSize, &RowSize, 0, 0, VT_I4);
         if (FAILED(hres)) {
             TRACE("ERROR when VariantChangeTypeEx\n");
         }
         drow = V_I4(&RowSize);
    } else drow = 0;

    if (!Is_Variant_Null(ColumnSize)) {
         hres = VariantChangeTypeEx(&ColumnSize, &ColumnSize, 0, 0, VT_I4);
         if (FAILED(hres)) {
             TRACE("ERROR when VariantChangeTypeEx\n");
         }
         dcol = V_I4(&ColumnSize);
    } else dcol = 0;

    hres = MSO_TO_OO_GetRangeAddress(iface, &left, &top, &right, &bottom);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress \n");
        return E_FAIL;
    }
    TRACE("drow = %i , dcol = %i \n", drow, dcol);
    TRACE("StartRow = %i, StartColumn = %i, EndRow = %i, EndColumn = %i \n", left, top, right, bottom);
    /*В OpenOffice оси направлены наоборот*/
    if (drow) right = left + drow - 1;
    if (dcol) bottom = top + dcol - 1;

    TRACE("StartRow = %i, StartColumn = %i, EndRow = %i, EndColumn = %i \n", left, top, right, bottom);
    /*создаем новый */
    lefttop.x = top+1;
    lefttop.y = left+1;
    rightbottom.x = bottom+1;
    rightbottom.y = right+1;


    hres = _I_RangeConstructor((LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Range_QueryInterface(punk, &IID_I_Range, (void**) &pCell);

    if (pCell == NULL) {
        TRACE("ERROR when QueryInterface\n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Range_Initialize((I_Range*)pCell, (I_Range*)wsh->pAllRange, lefttop, rightbottom);
        if (FAILED(hres)){
            TRACE("ERROR when initialize\n");
            I_Range_Release(pCell);
            return hres;
        }
    *RHS = pCell;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_RowDifferences(
        I_Range* iface,
        VARIANT Comparison,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Rows(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return I_Range_get_EntireRow(iface, RHS);
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Run(
        I_Range* iface,
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

static HRESULT WINAPI MSO_TO_OO_I_Range_Show(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ShowDependents(
        I_Range* iface,
        VARIANT Remove,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_ShowDetail(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_ShowDetail(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ShowErrors(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ShowPrecedents(
        I_Range* iface,
        VARIANT Remove,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Sort(
        I_Range* iface,
        VARIANT Key1,
        XlSortOrder Order1,
        VARIANT Key2,
        VARIANT Type,
        XlSortOrder Order2,
        VARIANT Key3,
        XlSortOrder Order3,
        XlYesNoGuess Header,
        VARIANT OrderCustom,
        VARIANT MatchCase,
        XlSortOrientation Orientation,
        XlSortMethod SortMethod,
        XlSortDataOption DataOption1,
        XlSortDataOption DataOption2,
        XlSortDataOption DataOption3,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_SortSpecial(
        I_Range* iface,
        XlSortMethod SortMethod,
        VARIANT Key1,
        XlSortOrder Order1,
        VARIANT Type,
        VARIANT Key2,
        XlSortOrder Order2,
        VARIANT Key3,
        XlSortOrder Order3,
        XlYesNoGuess Header,
        VARIANT OrderCustom,
        VARIANT MatchCase,
        XlSortOrientation Orientation,
        XlSortDataOption DataOption1,
        XlSortDataOption DataOption2,
        XlSortDataOption DataOption3,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_SoundNote(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_SpecialCells(
        I_Range* iface,
        XlCellType Type,
        VARIANT Value,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Style(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Style(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_SubscribeTo(
        I_Range* iface,
        BSTR Edition,
        XlSubscribeToFormat Format,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Subtotal(
        I_Range* iface,
        long GroupBy,
        XlConsolidationFunction Function,
        VARIANT TotalList,
        VARIANT Replace,
        VARIANT PageBreaks,
        XlSummaryRow SummaryBelowData,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Summary(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Table(
        I_Range* iface,
        VARIANT RowInput,
        VARIANT ColumnInput,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Text(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_TextToColumns(
        I_Range* iface,
        VARIANT Destination,
        XlTextParsingType DataType,
        XlTextQualifier TextQualifier,
        VARIANT ConsecutiveDelimiter,
        VARIANT Tab,
        VARIANT Semicolon,
        VARIANT Comma,
        VARIANT Space,
        VARIANT Other,
        VARIANT OtherChar,
        VARIANT FieldInfo,
        VARIANT DecimalSeparator,
        VARIANT ThousandsSeparator,
        VARIANT TrailingMinusNumbers,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Ungroup(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_UseStandardHeight(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_UseStandardHeight(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_UseStandardWidth(
        I_Range* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_UseStandardWidth(
        I_Range* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Validation(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Value2(
        I_Range* iface,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Value2(
        I_Range* iface,
        LCID lcid,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_AddComment(
        I_Range* iface,
        VARIANT Text,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Comment(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Phonetic(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_FormatConditions(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_ReadingOrder(
        I_Range* iface,
        long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_ReadingOrder(
        I_Range* iface,
        long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Hyperlinks(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Phonetics(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_SetPhonetic(
        I_Range* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_ID(
        I_Range* iface,
        BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_ID(
        I_Range* iface,
        BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_PrintOut(
        I_Range* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        VARIANT PrToFileName,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_PivotCell(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Dirty(
        I_Range* iface)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Errors(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_SmartTags(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Speak(
        I_Range* iface,
        VARIANT SpeakDirection,
        VARIANT SpeakFormulas)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_PasteSpecial(
        I_Range* iface,
        XlPasteType Paste,
        XlPasteSpecialOperation Operation,
        VARIANT SkipBlanks,
        VARIANT Transpose,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_AllowEdit(
        I_Range* iface,
        VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_ListObject(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_XPath(
        I_Range* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Range_GetTypeInfoCount(
        I_Range* iface,
        UINT *pctinfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_GetTypeInfo(
        I_Range* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_GetIDsOfNames(
        I_Range* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_range(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Invoke(
        I_Range* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    RangeImpl *This = (RangeImpl*)iface;
    IDispatch *dret, *pretdisp;
    HRESULT hres;
    long lret, lval;
    VARIANT vNull;
    XlHAlign halign;
    XlVAlign valign;
    VARIANT_BOOL vbin;
    VARIANT vRet,vtmp;
    VARIANT var1,var2;
    ITypeInfo *typeinfo;

    VariantInit(&vtmp);
    VariantInit(&vRet);
    VariantInit(&vNull);
    VariantInit(&var1);
    VariantInit(&var2);

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    switch(dispIdMember)
    {
    case dispid_range_value:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if ((pDispParams->cArgs>2)||(pDispParams->cArgs==0)) return E_FAIL;
            if (pDispParams->cArgs==1) {
                hres = MSO_TO_OO_I_Range_put_Value(iface, vNull, 0,  pDispParams->rgvarg[0]);
            }
            if (pDispParams->cArgs==2) {
                hres = MSO_TO_OO_I_Range_put_Value(iface, pDispParams->rgvarg[1], 0, pDispParams->rgvarg[0]);
            }
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pDispParams->cArgs>1) return E_FAIL;
            if (pDispParams->cArgs==0) 
                hres = MSO_TO_OO_I_Range_get_Value(iface, vNull, 0,  pVarResult);
            if (pDispParams->cArgs==1) 
                hres = MSO_TO_OO_I_Range_get_Value(iface, pDispParams->rgvarg[0], 0, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        }
    case dispid_range_merge:
        if (pDispParams->cArgs==1) {
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = VariantChangeTypeEx(&vtmp, &var1, 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE("(case 12) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            hres = MSO_TO_OO_I_Range_Merge(iface, vbin);
        } else
            hres = MSO_TO_OO_I_Range_Merge(iface, VARIANT_FALSE);
        if (FAILED(hres)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hres;
        }
        return S_OK;
    case dispid_range_formular1c1:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if ((pDispParams->cArgs>2)||(pDispParams->cArgs==0)) return E_FAIL;
            if (pDispParams->cArgs==1) {
                hres = MSO_TO_OO_I_Range_put_FormulaR1C1(iface, 0,  pDispParams->rgvarg[0]);
            }
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Range_get_FormulaR1C1(iface, 0,  pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        }
    case dispid_range_formula:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if ((pDispParams->cArgs>1)||(pDispParams->cArgs==0)) return E_FAIL;
            if (pDispParams->cArgs==1) {
                hres = MSO_TO_OO_I_Range_put_Formula(iface, 0,  pDispParams->rgvarg[0]);
            }
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pDispParams->cArgs!=0) return E_FAIL;
            hres = MSO_TO_OO_I_Range_get_Formula(iface, 0,  pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        }
    case dispid_range_clearcontents:
        vRet = MSO_TO_OO_I_Range_ClearContents(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;
        return S_OK;
    case dispid_range_clear:
        vRet = MSO_TO_OO_I_Range_Clear(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;
        return S_OK;
    case dispid_range_clearcomments:
        return MSO_TO_OO_I_Range_ClearComments(iface);
    case dispid_range_clearformats:
        vRet = MSO_TO_OO_I_Range_ClearFormats(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;
        return S_OK;
    case dispid_range_clearnotes:
        vRet = MSO_TO_OO_I_Range_ClearNotes(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;
        return S_OK;
    case dispid_range_clearoutline:
        vRet = MSO_TO_OO_I_Range_ClearOutline(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;
        return S_OK;
    default:
        /*For default*/
        hres = get_typeinfo_range(&typeinfo);
        if(FAILED(hres))
            return hres;

        hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

        if (FAILED(hres)) {
            TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
        }
        return hres;
    }
    WTRACE(L" dispIdMember = %i NOT REALIZE\n",dispIdMember);
    return E_NOTIMPL;
}

const I_RangeVtbl MSO_TO_OO_I_RangeVtbl =
{
    MSO_TO_OO_I_Range_QueryInterface,
    MSO_TO_OO_I_Range_AddRef,
    MSO_TO_OO_I_Range_Release,
    MSO_TO_OO_I_Range_GetTypeInfoCount,
    MSO_TO_OO_I_Range_GetTypeInfo,
    MSO_TO_OO_I_Range_GetIDsOfNames,
    MSO_TO_OO_I_Range_Invoke,
    MSO_TO_OO_I_Range_get_Application,
    MSO_TO_OO_I_Range_get_Creator,
    MSO_TO_OO_I_Range_get_Parent,
    MSO_TO_OO_I_Range_Activate,
    MSO_TO_OO_I_Range_get_AddIndent,
    MSO_TO_OO_I_Range_put_AddIndent,
    MSO_TO_OO_I_Range_get_Address,
    MSO_TO_OO_I_Range_get_AddressLocal,
    MSO_TO_OO_I_Range_AdvancedFilter,
    MSO_TO_OO_I_Range_ApplyNames,
    MSO_TO_OO_I_Range_ApplyOutlineStyles,
    MSO_TO_OO_I_Range_get_Areas,
    MSO_TO_OO_I_Range_AutoComplete,
    MSO_TO_OO_I_Range_AutoFill,
    MSO_TO_OO_I_Range_AutoFilter,
    MSO_TO_OO_I_Range_AutoFit,
    MSO_TO_OO_I_Range_AutoFormat,
    MSO_TO_OO_I_Range_AutoOutline,
    MSO_TO_OO_I_Range_BorderAround,
    MSO_TO_OO_I_Range_get_Borders,
    MSO_TO_OO_I_Range_Calculate,
    MSO_TO_OO_I_Range_get_Cells,
    MSO_TO_OO_I_Range_get_Characters,
    MSO_TO_OO_I_Range_CheckSpelling,
    MSO_TO_OO_I_Range_Clear,
    MSO_TO_OO_I_Range_ClearContents,
    MSO_TO_OO_I_Range_ClearFormats,
    MSO_TO_OO_I_Range_ClearNotes,
    MSO_TO_OO_I_Range_ClearOutline,
    MSO_TO_OO_I_Range_get_Column,
    MSO_TO_OO_I_Range_ColumnDifferences,
    MSO_TO_OO_I_Range_get_Columns,
    MSO_TO_OO_I_Range_get_ColumnWidth,
    MSO_TO_OO_I_Range_put_ColumnWidth,
    MSO_TO_OO_I_Range_Consolidate,
    MSO_TO_OO_I_Range_Copy,
    MSO_TO_OO_I_Range_CopyFromRecordset,
    MSO_TO_OO_I_Range_CopyPicture,
    MSO_TO_OO_I_Range_get_Count,
    MSO_TO_OO_I_Range_CreateNames,
    MSO_TO_OO_I_Range_CreatePublisher,
    MSO_TO_OO_I_Range_get_CurrentArray,
    MSO_TO_OO_I_Range_get_CurrentRegion,
    MSO_TO_OO_I_Range_Cut,
    MSO_TO_OO_I_Range_DataSeries,
    MSO_TO_OO_I_Range_get__Default,
    MSO_TO_OO_I_Range_put__Default,
    MSO_TO_OO_I_Range_Delete,
    MSO_TO_OO_I_Range_get_Dependents,
    MSO_TO_OO_I_Range_DialogBox1,
    MSO_TO_OO_I_Range_get_DirectDependents,
    MSO_TO_OO_I_Range_EditionOptions,
    MSO_TO_OO_I_Range_get_End,
    MSO_TO_OO_I_Range_get_EntireColumn,
    MSO_TO_OO_I_Range_get_EntireRow,
    MSO_TO_OO_I_Range_FillDown,
    MSO_TO_OO_I_Range_FillLeft,
    MSO_TO_OO_I_Range_FillRight,
    MSO_TO_OO_I_Range_FillUp,
    MSO_TO_OO_I_Range_Find,
    MSO_TO_OO_I_Range_FindNext,
    MSO_TO_OO_I_Range_FindPrevious,
    MSO_TO_OO_I_Range_get_Font,
    MSO_TO_OO_I_Range_get_Formula,
    MSO_TO_OO_I_Range_put_Formula,
    MSO_TO_OO_I_Range_get_FormulaArray,
    MSO_TO_OO_I_Range_put_FormulaArray,
    MSO_TO_OO_I_Range_get_FormulaLabel,
    MSO_TO_OO_I_Range_put_FormulaLabel,
    MSO_TO_OO_I_Range_get_FormulaHidden,
    MSO_TO_OO_I_Range_put_FormulaHidden,
    MSO_TO_OO_I_Range_get_FormulaLocal,
    MSO_TO_OO_I_Range_put_FormulaLocal,
    MSO_TO_OO_I_Range_get_FormulaR1C1,
    MSO_TO_OO_I_Range_put_FormulaR1C1,
    MSO_TO_OO_I_Range_get_FormulaR1C1Local,
    MSO_TO_OO_I_Range_put_FormulaR1C1Local,
    MSO_TO_OO_I_Range_FunctionWizard,
    MSO_TO_OO_I_Range_GoalSeek,
    MSO_TO_OO_I_Range_Group,
    MSO_TO_OO_I_Range_get_HasArray,
    MSO_TO_OO_I_Range_get_HasFormula,
    MSO_TO_OO_I_Range_get_Height,
    MSO_TO_OO_I_Range_get_Hidden,
    MSO_TO_OO_I_Range_put_Hidden,
    MSO_TO_OO_I_Range_get_HorizontalAlignment,
    MSO_TO_OO_I_Range_put_HorizontalAlignment,
    MSO_TO_OO_I_Range_get_IndentLevel,
    MSO_TO_OO_I_Range_put_IndentLevel,
    MSO_TO_OO_I_Range_InsertIndent,
    MSO_TO_OO_I_Range_Insert,
    MSO_TO_OO_I_Range_get_Interior,
    MSO_TO_OO_I_Range_get_Item,
    MSO_TO_OO_I_Range_put_Item,
    MSO_TO_OO_I_Range_Justify,
    MSO_TO_OO_I_Range_get_Left,
    MSO_TO_OO_I_Range_get_ListHeaderRows,
    MSO_TO_OO_I_Range_ListNames,
    MSO_TO_OO_I_Range_get_LocationInTable,
    MSO_TO_OO_I_Range_get_Locked,
    MSO_TO_OO_I_Range_put_Locked,
    MSO_TO_OO_I_Range_Merge,
    MSO_TO_OO_I_Range_UnMerge,
    MSO_TO_OO_I_Range_get_MergeArea,
    MSO_TO_OO_I_Range_get_MergeCells,
    MSO_TO_OO_I_Range_put_MergeCells,
    MSO_TO_OO_I_Range_get_Name,
    MSO_TO_OO_I_Range_put_Name,
    MSO_TO_OO_I_Range_NavigateArrow,
    MSO_TO_OO_I_Range_get__NewEnum,
    MSO_TO_OO_I_Range_get_Next,
    MSO_TO_OO_I_Range_NoteText,
    MSO_TO_OO_I_Range_get_NumberFormat,
    MSO_TO_OO_I_Range_put_NumberFormat,
    MSO_TO_OO_I_Range_get_NumberFormatLocal,
    MSO_TO_OO_I_Range_put_NumberFormatLocal,
    MSO_TO_OO_I_Range_get_Offset,
    MSO_TO_OO_I_Range_get_Orientation,
    MSO_TO_OO_I_Range_put_Orientation,
    MSO_TO_OO_I_Range_get_OutlineLevel,
    MSO_TO_OO_I_Range_put_OutlineLevel,
    MSO_TO_OO_I_Range_get_PageBreak,
    MSO_TO_OO_I_Range_put_PageBreak,
    MSO_TO_OO_I_Range_Parse,
    MSO_TO_OO_I_Range__PasteSpecial,
    MSO_TO_OO_I_Range_get_PivotField,
    MSO_TO_OO_I_Range_get_PivotItem,
    MSO_TO_OO_I_Range_get_PivotTable,
    MSO_TO_OO_I_Range_get_Precedents,
    MSO_TO_OO_I_Range_get_PrefixCharacter,
    MSO_TO_OO_I_Range_get_Previous,
    MSO_TO_OO_I_Range__PrintOut,
    MSO_TO_OO_I_Range_PrintPreview,
    MSO_TO_OO_I_Range_get_QueryTable,
    MSO_TO_OO_I_Range_get_Range,
    MSO_TO_OO_I_Range_RemoveSubtotal,
    MSO_TO_OO_I_Range_Replace,
    MSO_TO_OO_I_Range_get_Resize,
    MSO_TO_OO_I_Range_get_Row,
    MSO_TO_OO_I_Range_RowDifferences,
    MSO_TO_OO_I_Range_get_RowHeight,
    MSO_TO_OO_I_Range_put_RowHeight,
    MSO_TO_OO_I_Range_get_Rows,
    MSO_TO_OO_I_Range_Run,
    MSO_TO_OO_I_Range_Select,
    MSO_TO_OO_I_Range_Show,
    MSO_TO_OO_I_Range_ShowDependents,
    MSO_TO_OO_I_Range_get_ShowDetail,
    MSO_TO_OO_I_Range_put_ShowDetail,
    MSO_TO_OO_I_Range_ShowErrors,
    MSO_TO_OO_I_Range_ShowPrecedents,
    MSO_TO_OO_I_Range_get_ShrinkToFit,
    MSO_TO_OO_I_Range_put_ShrinkToFit,
    MSO_TO_OO_I_Range_Sort,
    MSO_TO_OO_I_Range_SortSpecial,
    MSO_TO_OO_I_Range_get_SoundNote,
    MSO_TO_OO_I_Range_SpecialCells,
    MSO_TO_OO_I_Range_get_Style,
    MSO_TO_OO_I_Range_put_Style,
    MSO_TO_OO_I_Range_SubscribeTo,
    MSO_TO_OO_I_Range_Subtotal,
    MSO_TO_OO_I_Range_get_Summary,
    MSO_TO_OO_I_Range_Table,
    MSO_TO_OO_I_Range_get_Text,
    MSO_TO_OO_I_Range_TextToColumns,
    MSO_TO_OO_I_Range_get_Top,
    MSO_TO_OO_I_Range_Ungroup,
    MSO_TO_OO_I_Range_get_UseStandardHeight,
    MSO_TO_OO_I_Range_put_UseStandardHeight,
    MSO_TO_OO_I_Range_get_UseStandardWidth,
    MSO_TO_OO_I_Range_put_UseStandardWidth,
    MSO_TO_OO_I_Range_get_Validation,
    MSO_TO_OO_I_Range_get_Value,
    MSO_TO_OO_I_Range_put_Value,
    MSO_TO_OO_I_Range_get_Value2,
    MSO_TO_OO_I_Range_put_Value2,
    MSO_TO_OO_I_Range_get_VerticalAlignment,
    MSO_TO_OO_I_Range_put_VerticalAlignment,
    MSO_TO_OO_I_Range_get_Width,
    MSO_TO_OO_I_Range_get_Worksheet,
    MSO_TO_OO_I_Range_get_WrapText,
    MSO_TO_OO_I_Range_put_WrapText,
    MSO_TO_OO_I_Range_AddComment,
    MSO_TO_OO_I_Range_get_Comment,
    MSO_TO_OO_I_Range_ClearComments,
    MSO_TO_OO_I_Range_get_Phonetic,
    MSO_TO_OO_I_Range_get_FormatConditions,
    MSO_TO_OO_I_Range_get_ReadingOrder,
    MSO_TO_OO_I_Range_put_ReadingOrder,
    MSO_TO_OO_I_Range_get_Hyperlinks,
    MSO_TO_OO_I_Range_get_Phonetics,
    MSO_TO_OO_I_Range_SetPhonetic,
    MSO_TO_OO_I_Range_get_ID,
    MSO_TO_OO_I_Range_put_ID,
    MSO_TO_OO_I_Range_PrintOut,
    MSO_TO_OO_I_Range_get_PivotCell,
    MSO_TO_OO_I_Range_Dirty,
    MSO_TO_OO_I_Range_get_Errors,
    MSO_TO_OO_I_Range_get_SmartTags,
    MSO_TO_OO_I_Range_Speak,
    MSO_TO_OO_I_Range_PasteSpecial,
    MSO_TO_OO_I_Range_get_AllowEdit,
    MSO_TO_OO_I_Range_get_ListObject,
    MSO_TO_OO_I_Range_get_XPath
};

extern HRESULT _I_RangeConstructor(LPVOID *ppObj)
{
    RangeImpl *range;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    range = HeapAlloc(GetProcessHeap(), 0, sizeof(*range));
    if (!range)
    {
        return E_OUTOFMEMORY;
    }

    range->_rangeVtbl = &MSO_TO_OO_I_RangeVtbl;
    range->ref = 0;
    range->pOORange = NULL;
    range->pwsheet = NULL;
    range->pApplication = NULL;
    range->is_release = 1;

    *ppObj = &range->_rangeVtbl;
    TRACE_OUT;
    return S_OK;
}
