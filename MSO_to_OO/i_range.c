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

static WCHAR const str__Default[] = {
    '_','D','e','f','a','u','l','t',0};
static WCHAR const str_ColumnWidth[] = {
    'C','o','l','u','m','n','W','i','d','t','h',0};
static WCHAR const str_font[] = {
    'F','o','n','t',0};
static WCHAR const str_value[] = {
    'V','a','l','u','e',0};
static WCHAR const str_select[] = {
    'S','e','l','e','c','t',0};
static WCHAR const str_notetext[] = {
    'N','o','t','e','T','e','x','t',0};
static WCHAR const str_clearcontents[] = {
    'C','l','e','a','r','C','o','n','t','e','n','t','s',0};
static WCHAR const str_column[] = {
    'C','o','l','u','m','n',0};
static WCHAR const str_row[] = {
    'R','o','w',0};
static WCHAR const str_horisontalalign[] = {
    'H','o','r','i','z','o','n','t','a','l','A','l','i','g','n','m','e','n','t',0};
static WCHAR const str_verticalalign[] = {
    'V','e','r','t','i','c','a','l','A','l','i','g','n','m','e','n','t',0};
static WCHAR const str_merge[] = {
    'M','e','r','g','e',0};
static WCHAR const str_unmerge[] = {
    'U','n','M','e','r','g','e',0};
static WCHAR const str_wraptext[] = {
    'W','r','a','p','T','e','x','t',0};
static WCHAR const str_application[] = {
    'A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR const str_parent[] = {
    'P','a','r','e','n','t',0};
static WCHAR const str_worksheet[] = {
    'W','o','r','k','s','h','e','e','t',0};
static WCHAR const str_clear[] = {
    'C','l','e','a','r',0};
static WCHAR const str_clearcomments[] = {
    'C','l','e','a','r','C','o','m','m','e','n','t','s',0};
static WCHAR const str_clearformats[] = {
    'C','l','e','a','r','F','o','r','m','a','t','s',0};
static WCHAR const str_clearnotes[] = {
    'C','l','e','a','r','N','o','t','e','s',0};
static WCHAR const str_clearoutline[] = {
    'C','l','e','a','r','O','u','t','l','i','n','e',0};
static WCHAR const str_interior[] = {
    'I','n','t','e','r','i','o','r',0};
static WCHAR const str_borders[] = {
    'B','o','r','d','e','r','s',0};
static WCHAR const str_count[] = {
    'C','o','u','n','t',0};
static WCHAR const str_delete[] = {
    'D','e','l','e','t','e',0};
static WCHAR const str_rowheight[] = {
    'R','o','w','H','e','i','g','h','t',0};
static WCHAR const str_copy[] = {
    'C','o','p','y',0};
static WCHAR const str_numberformat[] = {
    'N','u','m','b','e','r','F','o','r','m','a','t',0};
static WCHAR const str_numberformatlocal[] = {
    'N','u','m','b','e','r','F','o','r','m','a','t','L','o','c','a','l',0};
static WCHAR const str_height[] = {
    'H','e','i','g','h','t',0};
static WCHAR const str_width[] = {
    'W','i','d','t','h',0};
static WCHAR const str_left[] = {
    'L','e','f','t',0};
static WCHAR const str_top[] = {
    'T','o','p',0};
static WCHAR const str_shrinktofit[] = {
    'S','h','r','i','n','k','T','o','F','i','t',0};
static WCHAR const str_mergecells[] = {
    'M','e','r','g','e','C','e','l','l','s',0};
static WCHAR const str_locked[] = {
    'L','o','c','k','e','d',0};
static WCHAR const str_hidden[] = {
    'H','i','d','d','e','n',0};
static WCHAR const str_mergearea[] = {
    'M','e','r','g','e','A','r','e','a',0};
static WCHAR const str_autofit[] = {
    'A','u','t','o','F','i','t',0};

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

    TRACE("\n");

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
        if (This->pwsheet != NULL) {
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
    LPUNKNOWN pUnkOuter = NULL;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *ppObject = NULL;

    if (V_VT(&varRowIndex)==VT_BSTR) {

        hres = _I_RangeConstructor(pUnkOuter, (LPVOID*) &punk);
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
        return S_OK;
    } else {
        /*преобразовываем любой тип к I4*/

        hres = VariantChangeTypeEx(&varRowIndex, &varRowIndex, 0, 0, VT_I4);

        hres = VariantChangeTypeEx(&varColumnIndex, &varColumnIndex, 0, 0, VT_I4);

        if ((V_VT(&varRowIndex) != VT_I4) || (V_VT(&varColumnIndex) != VT_I4))
            return E_FAIL;

        /*Создаем новый объект I_Range that occupy one cell*/
        struct CELL_COORD cellCoord;

        hres = _I_RangeConstructor(pUnkOuter, (LPVOID*) &punk);

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

        return S_OK;
    }
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_ColumnWidth(
        I_Range* iface,
        long *pnColumnWidth)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_ColumnWidth(
        I_Range* iface,
        long nColumnWidth)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Font(
        I_Range* iface,
        IDispatch **ppFont)
{
    RangeImpl *This = (RangeImpl*)iface;

    HRESULT hr;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    IDispatch *pFont;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *ppFont = NULL;

    hr = _I_FontConstructor(pUnkOuter, (LPVOID*) &punk);

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Value(
        I_Range* iface,
        VARIANT varRangeValueDataType,
        VARIANT *pvarValue)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    hres = AutoWrap(DISPATCH_METHOD, pvarValue, V_DISPATCH(&resultCell), L"getFormula", 0);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_Value(
        I_Range* iface,
        VARIANT varRangeValueDataType,
        VARIANT varValue)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE(" \n");

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

                    hres = MSO_TO_OO_I_Range_put_Value(temp_range, vNull, pvar[j*maxi+i]);
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
            hres = AutoWrap(DISPATCH_METHOD, &res, V_DISPATCH(&resultCell), L"setFormula", 1, varValue);
            if (hres != S_OK) {
                TRACE("ERROR when setFormula \n");
                TRACE("    VT = %i \n",V_VT(&varValue));
            }
            return hres;
        }
    }
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Select(
        I_Range* iface,
        VARIANT *pvarResult)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    TRACE(" \n");

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
    return E_NOTIMPL;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearContents(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    VARIANT result,param,vRes;
    VariantInit(&result);
    VariantInit(&param);
    VariantInit(&vRes);
    V_VT(&result) = VT_NULL;
    HRESULT hres;

    if (This == NULL) return result;

    V_VT(&param) = VT_I4;
    V_I4(&param) = VALUE + DATETIME + STRING + FORMULA + OBJECTS;

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"clearContents", 1, param);

    if (hres != S_OK) {
        TRACE("ERROR when clearContents \n");
        return result;
    }

    return result;
}

static long WINAPI MSO_TO_OO_I_Range_get_Column(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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
    return lres + 1;
}

static long WINAPI MSO_TO_OO_I_Range_get_Row(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return lres + 1;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_HorizontalAlignment(
        I_Range* iface,
        XlHAlign *halign)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE(" \n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_HorizontalAlignment(
        I_Range* iface,
        XlHAlign halign)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_VerticalAlignment(
        I_Range* iface,
        XlVAlign *valign)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_VerticalAlignment(
        I_Range* iface,
        XlVAlign valign)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE(" \n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_Merge(
        I_Range* iface,
        VARIANT_BOOL flag)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes, vraddr;
    HRESULT hres;
    VariantInit(&param);
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = VARIANT_TRUE;

    if (flag == VARIANT_FALSE) {
        hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"Merge", 1, param);
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
    return S_OK;
    }
}

static HRESULT WINAPI MSO_TO_OO_I_Range_UnMerge(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes;
    HRESULT hres;
    VariantInit(&param);
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = VARIANT_FALSE; /*т.к. мы разбиваем ячейки*/

    hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"Merge", 1, param);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_WrapText(
        I_Range* iface,
        VARIANT_BOOL *pvbwraptext)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VARIANT vRes;
    HRESULT hres;

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, This->pOORange, L"isTextWrapped", 0);

    if (FAILED(hres)) {
        return hres;
    }
    *pvbwraptext = V_BOOL(&vRes);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_WrapText(
        I_Range* iface,
        VARIANT_BOOL pvbwraptext)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    VARIANT param;
    VARIANT vRes;
    HRESULT hres;
    VariantInit(&param);
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = pvbwraptext;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"isTextWrapped", 1, param);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Application(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *value = This->pApplication;
    I_ApplicationExcel_AddRef(This->pApplication);

    if (value==NULL)
        return E_POINTER;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Parent(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *value = This->pwsheet;
    I_Worksheet_AddRef(This->pwsheet);

    if (value==NULL)
        return E_POINTER;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Worksheet(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *value = This->pwsheet;
    I_Worksheet_AddRef(This->pwsheet);

    if (value==NULL)
        return E_POINTER;

    return S_OK;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_Clear(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return result;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_ClearComments(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return S_OK;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearFormats(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return result;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearNotes(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return result;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_ClearOutline(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return result;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Interior(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;

    HRESULT hr;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    IDispatch *pInterior;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    *value = NULL;

    hr = _I_InteriorConstructor(pUnkOuter, (LPVOID*) &punk);

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Borders(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;

    HRESULT hr;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    IDispatch *pBorders;

    TRACE(" \n");

    if (This == NULL) return E_POINTER;

    *value = NULL;

    hr = _I_BordersConstructor(pUnkOuter, (LPVOID*) &punk);

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Count(
        I_Range* iface,
        long *lcount)
{
    HRESULT hres;
    long startrow, startcolumn, endrow, endcolumn,w,h;

    TRACE("\n");
    hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress\n");
        return hres;
    }
    w = endcolumn - startcolumn + 1;
    h = endrow - startrow + 1;

    *lcount = w*h;

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

    TRACE("\n");

    if (This == NULL) return E_POINTER;
    *value = NULL;

    hres = MSO_TO_OO_GetRangeAddress(iface, &startrow, &startcolumn, &endrow, &endcolumn);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress\n");
        return hres;
    }

    if ((V_VT(&param)==VT_NULL)||(V_VT(&param)==VT_EMPTY)) {
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
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_RowHeight(
        I_Range* iface,
        long *pnrowheight)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_RowHeight(
        I_Range* iface,
        long nrowheight)
{
    RangeImpl *This = (RangeImpl*)iface;

    TRACE("\n");

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

    TRACE("\n");

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

    if ((V_VT(&RangeTo)==VT_NULL)||(V_VT(&RangeTo)==VT_EMPTY)) {
        TRACE("(To Clipboard)\n");
        *value = (IDispatch*)iface;
        IDispatch_AddRef(*value);
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
        return S_OK;
    }

    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_NumberFormat(
        I_Range* iface,
        VARIANT *pnumbformat)
{
/*TODO*/
    TRACE(" \n");
    V_VT(pnumbformat) = VT_BSTR;
    V_BSTR(pnumbformat) = SysAllocString(L"");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_NumberFormat(
        I_Range* iface,
        VARIANT numbformat)
{
/*TODO*/
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_NumberFormatLocal(
        I_Range* iface,
        VARIANT *pnumbformat)
{
/*TODO*/
    TRACE("\n");
    V_VT(pnumbformat) = VT_BSTR;
    V_BSTR(pnumbformat) = SysAllocString(L"");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_NumberFormatLocal(
        I_Range* iface,
        VARIANT numbformat)
{
/*TODO*/
    TRACE("\n");
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Height(
        I_Range* iface,
        VARIANT *value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    VARIANT vsize;

    VariantInit(&vsize);

    TRACE("\n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Width(
        I_Range* iface,
        VARIANT *value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    VARIANT vsize;

    VariantInit(&vsize);

    TRACE("\n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_Left(
        I_Range* iface,
        VARIANT *value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    VARIANT vsize;

    VariantInit(&vsize);

    TRACE(" \n");

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

    TRACE("\n");

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

    TRACE(" \n");

    if (This == NULL) return E_POINTER;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vRes, This->pOORange, L"ShrinkToFit", 0);

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    *pparam = vtmp;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_ShrinkToFit(
        I_Range* iface,
        VARIANT param)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT vRes;
    HRESULT hres;

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vRes, This->pOORange, L"ShrinkToFit", 1, param);
    if (FAILED(hres))  {
       TRACE("ERROR when ShrinkToFit \n");
       return hres;
    }

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

    TRACE(" \n");

    if (This == NULL) return E_POINTER;

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &vRes, This->pOORange, L"getIsMerged", 0);

    hres = VariantChangeTypeEx(&vtmp, &vRes,0,0,VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
        return E_FAIL;
    }
    *pparam = vtmp;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_put_MergeCells(
        I_Range* iface,
        VARIANT param)
{
    HRESULT hres;
    VARIANT vtmp;

    VariantInit(&vtmp);

    hres = VariantChangeTypeEx(&vtmp, &param, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR when VariantChangeTypeEx\n");
    }

    if (V_BOOL(&vtmp)==VARIANT_FALSE) {
        hres = MSO_TO_OO_I_Range_UnMerge(iface);
    } else {
        hres = MSO_TO_OO_I_Range_Merge(iface, VARIANT_TRUE);
    }

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

    TRACE(" \n");

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

    TRACE("\n");

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

    TRACE(" \n");

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

    TRACE("\n");

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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_get_MergeArea(
        I_Range* iface,
        IDispatch **value)
{
    RangeImpl *This = (RangeImpl*)iface;
    HRESULT hres;
    long left, right,top,bottom;

    TRACE(" \n");

    hres = MSO_TO_OO_GetRangeAddress(iface, &left, &top, &right, &bottom);
    if (FAILED(hres)) {
        TRACE("ERROR when GetRangeAddress\n");
    }
    /*Если не 1 ячейка, значит ошибка*/
    if ((left!=right)||(top!=bottom)) return E_FAIL;
    /*Получить объединенную область или ячейку*/

    *value = (IDispatch*)This;
    IDispatch_AddRef(*value);

    return S_OK;
}

static VARIANT WINAPI MSO_TO_OO_I_Range_AutoFit(
        I_Range* iface)
{
    RangeImpl *This = (RangeImpl*)iface;
    VARIANT result;
    VARIANT res;
    VARIANT columns;
    VARIANT param;
    HRESULT hres;

    TRACE("\n");

    VariantInit(&result);
    VariantInit(&res);
    VariantInit(&columns);
    VariantInit(&param);

    V_VT(&result) = VT_NULL;

    if (This == NULL) return result;

    hres = AutoWrap(DISPATCH_METHOD, &columns, This->pOORange, L"getColumns", 0);
    if (hres != S_OK) {
        TRACE("Error when getColumns\n");
        return result;
    }

    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = VARIANT_TRUE;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, V_DISPATCH(&columns), L"OptimalWidth", 1, param);
    if (FAILED(hres)) TRACE("ERROR when OptimalWidth\n");
    IDispatch_Release(V_DISPATCH(&columns));

    VariantClear(&res);
    VariantClear(&param);

    return result;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Range_GetTypeInfoCount(
        I_Range* iface,
        UINT *pctinfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Range_GetTypeInfo(
        I_Range* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE(" \n");
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
    if (!lstrcmpiW(*rgszNames, str__Default)) {
        *rgDispId = 1;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_ColumnWidth)) {
        *rgDispId = 2;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_font)) {
        *rgDispId = 3;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_value)) {
        *rgDispId = 4;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_select)) {
        *rgDispId = 5;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_notetext)) {
        *rgDispId = 6;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_clearcontents)) {
        *rgDispId = 7;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_column)) {
        *rgDispId = 8;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_row)) {
        *rgDispId = 9;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_horisontalalign)) {
        *rgDispId = 10;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_verticalalign)) {
        *rgDispId = 11;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_merge)) {
        *rgDispId = 12;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_unmerge)) {
        *rgDispId = 13;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_wraptext)) {
        *rgDispId = 14;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = 15;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_parent)) {
        *rgDispId = 16;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_worksheet)) {
        *rgDispId = 17;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_clear)) {
        *rgDispId = 18;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_clearcomments)) {
        *rgDispId = 19;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_clearformats)) {
        *rgDispId = 20;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_clearnotes)) {
        *rgDispId = 21;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_clearoutline)) {
        *rgDispId = 22;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_interior)) {
        *rgDispId = 23;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_borders)) {
        *rgDispId = 24;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_count)) {
        *rgDispId = 25;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_delete)) {
        *rgDispId = 26;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_rowheight)) {
        *rgDispId = 27;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_copy)) {
        *rgDispId = 28;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_numberformat)) {
        *rgDispId = 29;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_numberformatlocal)) {
        *rgDispId = 30;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_height)) {
        *rgDispId = 31;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_width)) {
        *rgDispId = 32;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_left)) {
        *rgDispId = 33;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_top)) {
        *rgDispId = 34;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_shrinktofit)) {
        *rgDispId = 35;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_mergecells)) {
        *rgDispId = 36;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_locked)) {
        *rgDispId = 37;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_hidden)) {
        *rgDispId = 38;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_mergearea)) {
        *rgDispId = 39;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_autofit)) {
        *rgDispId = 40;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L"%s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
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

    VariantInit(&vtmp);
    VariantInit(&vRet);
    VariantInit(&vNull);
    VariantInit(&var1);
    VariantInit(&var2);

    TRACE("\n");

    if (This == NULL) return E_POINTER;

    switch(dispIdMember)
    {
    case 1:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &var2);
            if (pDispParams->cArgs!=2) return E_FAIL;
            hres = MSO_TO_OO_I_Range_get__Default(iface, var2, var1, &dret);
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
    case 2:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            lret=1;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            /*преобразовываем любой тип к I4*/
            hres = VariantChangeTypeEx(&vtmp, &var1, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("(case 2) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            lret = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Range_put_ColumnWidth(iface, lret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Range_get_ColumnWidth(iface, &lret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lret;
            }
            return S_OK;
        }
    case 3:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            if (pDispParams->cArgs!=0) return E_FAIL;

            hres = MSO_TO_OO_I_Range_get_Font(iface, &dret);
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
    case 4:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if ((pDispParams->cArgs>2)||(pDispParams->cArgs==0)) return E_FAIL;
            if (pDispParams->cArgs==1) {
TRACE("VT = %i\n",V_VT(&(pDispParams->rgvarg[0])));
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var2);
TRACE("VT = %i\n",V_VT(&var2));
                hres = MSO_TO_OO_I_Range_put_Value(iface, vNull, var2);
            }
            if (pDispParams->cArgs==2) {
                /*Привести параметры к типу VARIANT если они переданы по ссылке*/
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[1], &var1);
                MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var2);

                hres = MSO_TO_OO_I_Range_put_Value(iface, var1, var2);
            }
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pDispParams->cArgs>1) return E_FAIL;
            if (pDispParams->cArgs==0) 
                hres = MSO_TO_OO_I_Range_get_Value(iface, vNull, pVarResult);
            if (pDispParams->cArgs==1) 
                hres = MSO_TO_OO_I_Range_get_Value(iface, pDispParams->rgvarg[0],  pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        }
    case 5:
        hres = MSO_TO_OO_I_Range_Select(iface, &vRet);
        if (FAILED(hres)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hres;
        }
        if (pVarResult!=NULL) 
            *pVarResult = vRet;

        return S_OK;
    case 6:
        /*MSO_TO_OO_I_Range_NoteText*/
        return E_NOTIMPL;
    case 7:
        vRet = MSO_TO_OO_I_Range_ClearContents(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;

        return S_OK;
    case 8:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            lret = MSO_TO_OO_I_Range_get_Column(iface);
            if (lret<1) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return E_FAIL;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lret;
            }
            return S_OK;
        }
    case 9:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            lret = MSO_TO_OO_I_Range_get_Row(iface);
            if (lret<1) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return E_FAIL;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lret;
            }
            return S_OK;
        }
    case 10:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = VariantChangeTypeEx(&vtmp, &var1, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("(case 10) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            lret = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Range_put_HorizontalAlignment(iface, lret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Range_get_HorizontalAlignment(iface, &halign);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = halign;
            }
            return S_OK;
        }
    case 11:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = VariantChangeTypeEx(&vtmp, &var1, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("(case 11) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            lret = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Range_put_VerticalAlignment(iface, lret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Range_get_VerticalAlignment(iface, &valign);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = valign;
            }
            return S_OK;
        }
    case 12:
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
    case 13:
        hres = MSO_TO_OO_I_Range_UnMerge(iface);
        if (FAILED(hres)) {
            pExcepInfo->bstrDescription=SysAllocString(str_error);
            return hres;
        }
        return S_OK;
    case 14:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = VariantChangeTypeEx(&vtmp, &var1, 0, 0, VT_BOOL);
            if (FAILED(hres)) {
                TRACE("(case 14) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            hres = MSO_TO_OO_I_Range_put_WrapText(iface, vbin);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Range_get_WrapText(iface, &vbin);
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
    case 15:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Application(iface,&dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case 16:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Parent(iface,&dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case 17:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Worksheet(iface,&dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case 18:
        vRet = MSO_TO_OO_I_Range_Clear(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;

        return S_OK;
    case 19:
        return MSO_TO_OO_I_Range_ClearComments(iface);
    case 20:
        vRet = MSO_TO_OO_I_Range_ClearFormats(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;

        return S_OK;
    case 21:
        vRet = MSO_TO_OO_I_Range_ClearNotes(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;

        return S_OK;
    case 22:
        vRet = MSO_TO_OO_I_Range_ClearOutline(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;

        return S_OK;
    case 23:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Interior(iface,&dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case 24:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Borders(iface,&dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pDispParams->cArgs!=1) {
                if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=dret;
                } else {
                    IDispatch_Release(dret);
                }
            } else {
            /*Привести параметры к типу VARIANT если они переданы по ссылке*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("(24) ERROR when VariantChangeType \n");
                return hres;
            }
            lval = V_I4(&vtmp);
            hres = I_Borders_get__Default((I_Borders*)dret,lval, &pretdisp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                    V_VT(pVarResult)=VT_DISPATCH;
                    V_DISPATCH(pVarResult)=pretdisp;
                    IDispatch_Release(dret);
                } else {
                    IDispatch_Release(dret);
                    IDispatch_Release(pretdisp);
                }
            }
            return S_OK;
        }
    case 25:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Count(iface,&lret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_I4;
                V_I4(pVarResult)=lret;
            }
            return S_OK;
        }
    case 26://Delete
        switch (pDispParams->cArgs) {
        case 0:
            TRACE("DELETE - NUMBER OF PARAMETERS 0\n");
            VariantInit(&var1);
            VariantClear(&var1);
            V_VT(&var1) = VT_EMPTY;
            hres = MSO_TO_OO_I_Range_Delete(iface, var1, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        case 1:
            VariantInit(&var1);
            TRACE("DELETE - NUMBER OF PARAMETERS 1\n");
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_Delete(iface, var1, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        default:
            TRACE(" (case 26) ERROR Parameters\n");
            return E_FAIL;
        }
        return S_OK;
    case 27:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            lret=1;
            /*преобразовываем любой тип к I4*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = VariantChangeTypeEx(&vtmp, &var1, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("(case 27) ERROR when VariantChangeTypeEx\n");
              return E_FAIL;
            }
            lret = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Range_put_RowHeight(iface, lret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Range_get_RowHeight(iface, &lret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lret;
            }
            return S_OK;
        }
    case 28:
        switch (pDispParams->cArgs) {
        case 0:
TRACE("Parametr 0\n");
            VariantInit(&var1);
            VariantClear(&var1);
            hres = MSO_TO_OO_I_Range_Copy(iface, var1, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        case 1:
TRACE("Parametr 1\n");
            VariantInit(&var1);
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_Copy(iface, var1, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        default:
            TRACE("(case 28) ERROR Parameters\n");
            return E_FAIL;
        }
        return S_OK;
    case 29://NumberFormat
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_put_NumberFormat(iface, var1);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pDispParams->cArgs!=0) return E_FAIL;

            hres = MSO_TO_OO_I_Range_get_NumberFormat(iface, &vRet);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BSTR;
                V_BSTR(pVarResult) = SysAllocString(V_BSTR(&vRet));
            }
            VariantClear(&vRet);
            return S_OK;
        }
    case 30://NumberFormatLocal
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_put_NumberFormatLocal(iface, var1);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pDispParams->cArgs!=0) return E_FAIL;

            hres = MSO_TO_OO_I_Range_get_NumberFormatLocal(iface, &vRet);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BSTR;
                V_BSTR(pVarResult) = SysAllocString(V_BSTR(&vRet));
            }
            VariantClear(&vRet);
            return S_OK;
        }
    case 31://height
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Height(iface, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return hres;
        }
    case 32://width
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Width(iface, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return hres;
        }
    case 33://left
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Left(iface, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return hres;
        }
    case 34://top
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_Top(iface, pVarResult);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return hres;
        }
    case 35://ShrinkToFit
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("ERROR INVALID ARGUMENT \n");
                return E_INVALIDARG;
            }
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_put_ShrinkToFit(iface, var1);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pVarResult!=NULL){
                hres = MSO_TO_OO_I_Range_get_ShrinkToFit(iface, pVarResult);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                return hres;
            }
            TRACE("pVarResult = NULL \n");
            return E_FAIL;
        }
    case 36://MergeCells
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("ERROR INVALID ARGUMENT \n");
                return E_INVALIDARG;
            }
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_put_MergeCells(iface, var1);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pVarResult!=NULL){
                hres = MSO_TO_OO_I_Range_get_MergeCells(iface, pVarResult);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                return hres;
            }
            TRACE("pVarResult = NULL \n");
            return E_FAIL;
        }
    case 37://Locked
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("ERROR INVALID ARGUMENT \n");
                return E_INVALIDARG;
            }
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_put_Locked(iface, var1);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pVarResult!=NULL){
                hres = MSO_TO_OO_I_Range_get_Locked(iface, pVarResult);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                return hres;
            }
            TRACE("pVarResult = NULL \n");
            return E_FAIL;
        }
    case 38://Hidden
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("ERROR INVALID ARGUMENT \n");
                return E_INVALIDARG;
            }
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &var1);
            hres = MSO_TO_OO_I_Range_put_Hidden(iface, var1);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            if (pVarResult!=NULL){
                hres = MSO_TO_OO_I_Range_get_Hidden(iface, pVarResult);
                if (FAILED(hres)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hres;
                }
                return hres;
            }
            TRACE("pVarResult = NULL \n");
            return E_FAIL;
        }
    case 39://MergeArea
        if (wFlags==DISPATCH_PROPERTYPUT) {
            TRACE("\n");
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Range_get_MergeArea(iface, &dret);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
                return hres;
            } else {
                IDispatch_Release(dret);
            }
            TRACE("pVarResult = NULL \n");
            return E_FAIL;
        }
    case 40:
        vRet = MSO_TO_OO_I_Range_AutoFit(iface);
        if (pVarResult!=NULL)
            *pVarResult = vRet;

        return S_OK;
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
    MSO_TO_OO_I_Range_get__Default,
    MSO_TO_OO_I_Range_get_ColumnWidth,
    MSO_TO_OO_I_Range_put_ColumnWidth,
    MSO_TO_OO_I_Range_get_Font,
    MSO_TO_OO_I_Range_get_Value,
    MSO_TO_OO_I_Range_put_Value,
    MSO_TO_OO_I_Range_Select,
    MSO_TO_OO_I_Range_NoteText,
    MSO_TO_OO_I_Range_ClearContents,
    MSO_TO_OO_I_Range_get_Column,
    MSO_TO_OO_I_Range_get_Row,
    MSO_TO_OO_I_Range_get_HorizontalAlignment,
    MSO_TO_OO_I_Range_put_HorizontalAlignment,
    MSO_TO_OO_I_Range_get_VerticalAlignment,
    MSO_TO_OO_I_Range_put_VerticalAlignment,
    MSO_TO_OO_I_Range_Merge,
    MSO_TO_OO_I_Range_UnMerge,
    MSO_TO_OO_I_Range_get_WrapText,
    MSO_TO_OO_I_Range_put_WrapText,
    MSO_TO_OO_I_Range_get_Application,
    MSO_TO_OO_I_Range_get_Parent,
    MSO_TO_OO_I_Range_get_Worksheet,
    MSO_TO_OO_I_Range_Clear,
    MSO_TO_OO_I_Range_ClearComments,
    MSO_TO_OO_I_Range_ClearFormats,
    MSO_TO_OO_I_Range_ClearNotes,
    MSO_TO_OO_I_Range_ClearOutline,
    MSO_TO_OO_I_Range_get_Interior,
    MSO_TO_OO_I_Range_get_Borders,
    MSO_TO_OO_I_Range_get_Count,
    MSO_TO_OO_I_Range_Delete,
    MSO_TO_OO_I_Range_get_RowHeight,
    MSO_TO_OO_I_Range_put_RowHeight,
    MSO_TO_OO_I_Range_Copy,
    MSO_TO_OO_I_Range_get_NumberFormat,
    MSO_TO_OO_I_Range_put_NumberFormat,
    MSO_TO_OO_I_Range_get_NumberFormatLocal,
    MSO_TO_OO_I_Range_put_NumberFormatLocal,
    MSO_TO_OO_I_Range_get_Height,
    MSO_TO_OO_I_Range_get_Width,
    MSO_TO_OO_I_Range_get_Left,
    MSO_TO_OO_I_Range_get_Top,
    MSO_TO_OO_I_Range_get_ShrinkToFit,
    MSO_TO_OO_I_Range_put_ShrinkToFit,
    MSO_TO_OO_I_Range_get_MergeCells,
    MSO_TO_OO_I_Range_put_MergeCells,
    MSO_TO_OO_I_Range_get_Locked,
    MSO_TO_OO_I_Range_put_Locked,
    MSO_TO_OO_I_Range_get_Hidden,
    MSO_TO_OO_I_Range_put_Hidden,
    MSO_TO_OO_I_Range_get_MergeArea,
    MSO_TO_OO_I_Range_AutoFit
};


RangeImpl MSO_TO_OO_Range =
{
    &MSO_TO_OO_I_RangeVtbl,
    0,
    NULL,
    NULL,
    NULL
};

extern HRESULT _I_RangeConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    RangeImpl *range;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

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

    *ppObj = &range->_rangeVtbl;

    return S_OK;
}
