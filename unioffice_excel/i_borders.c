/*
 * IBorders interface functions
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

/* Enumeration membership excludes the four legacy aliases accepted by Item. */
static const XlBordersIndex border_indices[] = {
    xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop,
    xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal
};
enum { BORDER_COUNT = sizeof(border_indices) / sizeof(border_indices[0]) };

ITypeInfo *ti_borders = NULL;
ITypeInfo *ti_borders_dispatch = NULL;

HRESULT get_typeinfo_borders(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_borders) {
        *typeinfo = ti_borders;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        ERR("LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Borders, &ti_borders);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_borders;

    return hres;
}

static HRESULT get_borders_dispatch_typeinfo(ITypeInfo **result)
{
    ITypeInfo *legacy;
    ITypeLib *typelib;
    UINT index;
    HRESULT hr;
    if (!ti_borders_dispatch) {
        hr = get_typeinfo_borders(&legacy);
        if (FAILED(hr)) return hr;
        hr = ITypeInfo_GetContainingTypeLib(legacy, &typelib, &index);
        if (FAILED(hr)) return hr;
        hr = ITypeLib_GetTypeInfoOfGuid(typelib, &DIID_D_Borders, &ti_borders_dispatch);
        ITypeLib_Release(typelib);
        if (FAILED(hr)) return hr;
    }
    *result = ti_borders_dispatch;
    return S_OK;
}

/*IBorders interface*/

#define BORDERS_THIS(iface) DEFINE_THIS(BordersImpl, borders, iface)

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Borders_AddRef(
        I_Borders* iface)
{
    BordersImpl *This = BORDERS_THIS(iface);
    ULONG ref;
    TRACE("REF = %i \n", This->ref);

    if (!This) {
        ERR("objectis NULL \n");
        return E_POINTER;
    }

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_QueryInterface(
        I_Borders* iface,
        REFIID riid,
        void **ppvObject)
{
    BordersImpl *This = BORDERS_THIS(iface);
    WCHAR str_clsid[39];
    
    if (!ppvObject) return E_POINTER;
    *ppvObject = NULL;

    if (!This) {
        ERR("object is NULL \n");
        return E_POINTER;
    }

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Borders) ||
            IsEqualGUID(riid, &DIID_D_Borders)) {
        *ppvObject = BORDERS_BORDERS(This);
        I_Borders_AddRef((I_Borders*)(*ppvObject));
        return S_OK;
    }
    if (IsEqualGUID(riid, &IID_IEnumVARIANT)) {
        *ppvObject = BORDERS_ENUM(This);
        IUnknown_AddRef((IUnknown*)(*ppvObject));
        return S_OK;
    }
    
    StringFromGUID2(riid, str_clsid, 39);
    WERR(L"(%s) not supported \n", str_clsid);
    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Borders_Release(
        I_Borders* iface)
{
    BordersImpl *This = BORDERS_THIS(iface);
    ULONG ref;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) {
        ERR("object is NULL \n");
        return E_POINTER;
    }

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pRange) {
            I_Range_Release(This->pRange);
            This->pRange = NULL;
        }
        if (This->pOORange) {
            IDispatch_Release(This->pOORange);
            This->pOORange = NULL;
        }      
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
        DELETE_OBJECT;
    }
    return ref;
}

/*** I_Borders methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Application(
        I_Borders* iface,
        IDispatch **value)
{
    BordersImpl *This = BORDERS_THIS(iface);
    HRESULT hres;
    TRACE_IN;

    *value = NULL;

    if (!This) {
        ERR("object is NULL \n");
        return E_POINTER;
    }
    
    if (!(This->pRange)) {
        ERR("object pRange is NULL \n");                     
        return E_POINTER;
    }

    hres = I_Range_get_Application(This->pRange, value);
    if (FAILED(hres)) {
        ERR("get_Application \n");
        return E_FAIL;                  
    }

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Parent(
        I_Borders* iface,
        IDispatch **value)
{
    BordersImpl *This = BORDERS_THIS(iface);
    TRACE_IN;

    *value = NULL;

    if (!This) {
        ERR("object is NULL \n");
        return E_POINTER;
    }

    *value = (IDispatch*)This->pRange;
    if (!(*value)) {
        ERR("object pRange is NULL \n");
        return E_FAIL;
    }
    I_Range_AddRef(This->pRange);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Color(
        I_Borders* iface,
        LONG *plcolor)
{
    BordersImpl *This = BORDERS_THIS(iface);
    HRESULT hres;
    I_Border *border_tmp;
    TRACE_IN;

    I_Borders_get_Item(iface, xlEdgeTop, (IDispatch**) &border_tmp);
    I_Border_get_Color(border_tmp, plcolor);
    I_Border_Release(border_tmp);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Color(
        I_Borders* iface,
        LONG lcolor)
{
    BordersImpl *This = BORDERS_THIS(iface);
    HRESULT hres;
    I_Border *border_tmp;
    int i;
    TRACE_IN;
    TRACE(" lcolor = %i\n", lcolor);

    if (!This) {
        ERR("object is NULL \n");
        return E_POINTER;
    }

    for (i = xlEdgeLeft; i <= xlInsideHorizontal; i++) {
        I_Borders_get_Item(iface, i,(IDispatch**) &border_tmp);
        I_Border_put_Color(border_tmp, lcolor);
        I_Border_Release(border_tmp);
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_ColorIndex(
        I_Borders* iface,
        LONG *plcolorindex)
{
    BordersImpl *This = BORDERS_THIS(iface);
    LONG tmpcolor;
    int i;
    HRESULT hres;
    TRACE_IN;

    if (!This) {
        ERR("object is NULL \n");
        return E_POINTER;
    }

    hres = I_Borders_get_Color(iface,&tmpcolor);
    if (FAILED(hres)) {
        return hres;
    }
    for (i=0;i<56;i++)
        if (color[i]==tmpcolor) {
            *plcolorindex = i+1;
            TRACE_OUT;
            return S_OK;
        }

    ERR("Color don`t have colorindex \n");
    *plcolorindex = 1;/*белый цвет*/
    /*Отправляем что все хорошо, на всякий случай*/
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_ColorIndex(
        I_Borders* iface,
        LONG lcolorindex)
{
    BordersImpl *This = BORDERS_THIS(iface);
    LONG tmpcolor;
    TRACE_IN;

    if (!This) {
        ERR("object is NULL \n");
        return E_POINTER;
    }

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;
    TRACE_OUT;
    if ((lcolorindex<1)||(lcolorindex>56)) {
        ERR("Incorrect colorindex %i\n", lcolorindex);
        return S_OK;
    } else 
        return I_Borders_put_Color(iface,color[lcolorindex-1]);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Creator(
        I_Borders* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_LineStyle(
        I_Borders* iface, XlLineStyle *plinestyle)
{
    I_Border *border;
    HRESULT hr;
    if (!plinestyle) return E_POINTER;
    hr = I_Borders_get_Item(iface, xlEdgeTop, (IDispatch**)&border);
    if (FAILED(hr)) return hr;
    hr = I_Border_get_LineStyle(border, plinestyle);
    I_Border_Release(border);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_LineStyle(
        I_Borders* iface, XlLineStyle linestyle)
{
    I_Border *border;
    HRESULT hr;
    int key;
    for (key = xlEdgeLeft; key <= xlInsideHorizontal; ++key) {
        hr = I_Borders_get_Item(iface, key, (IDispatch**)&border);
        if (FAILED(hr)) return hr;
        hr = I_Border_put_LineStyle(border, linestyle);
        I_Border_Release(border);
        if (FAILED(hr)) return hr;
    }
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Weight(
        I_Borders* iface, XlBorderWeight *pweight)
{
    I_Border *border;
    HRESULT hr;
    if (!pweight) return E_POINTER;
    hr = I_Borders_get_Item(iface, xlEdgeTop, (IDispatch**)&border);
    if (FAILED(hr)) return hr;
    hr = I_Border_get_Weight(border, pweight);
    I_Border_Release(border);
    return hr;
}

static BOOL borders_match_weight(XlBorderWeight weight, BOOL *have_weight,
        XlBorderWeight *common)
{
    if (!*have_weight) {
        *common = weight;
        *have_weight = TRUE;
    }
    return *common == weight;
}

/* Only ambiguous lines need per-cell inspection; colors/styles alone are not mixed weights. */
static HRESULT borders_weight_variant(I_Borders *iface, VARIANT *result)
{
    BordersImpl *This = BORDERS_THIS(iface);
    LONG first_row, first_column, last_row, last_column, rows, columns;
    LONG x0, x1, y0, y1, x, y;
    XlBordersIndex key, cell_key;
    XlBorderWeight weight, common = xlThin;
    BOOL uniform, have_weight = FALSE;
    VARIANT cell, row, column;
    HRESULT hr;
    VariantInit(result);
    hr = MSO_TO_OO_GetRangeAddress(This->pRange, &first_row, &first_column, &last_row, &last_column);
    if (FAILED(hr)) return hr;
    rows = last_row - first_row + 1;
    columns = last_column - first_column + 1;
    if (rows <= 0 || columns <= 0) return E_UNEXPECTED;
    for (key = xlEdgeLeft; key <= xlInsideHorizontal; ++key) {
        if ((key == xlInsideVertical && columns == 1) ||
                (key == xlInsideHorizontal && rows == 1)) continue;
        hr = MSO_TO_OO_BorderWeight(This->pOORange, key, &weight, &uniform);
        if (FAILED(hr)) return hr;
        if (uniform) {
            if (!borders_match_weight(weight, &have_weight, &common)) goto mixed;
            continue;
        }
        x0 = y0 = 0;
        x1 = columns;
        y1 = rows;
        cell_key = key;
        switch (key) {
        case xlEdgeLeft: x1 = 1; break;
        case xlEdgeRight: x0 = columns - 1; break;
        case xlEdgeTop: y1 = 1; break;
        case xlEdgeBottom: y0 = rows - 1; break;
        case xlInsideVertical: x0 = 1; cell_key = xlEdgeLeft; break;
        case xlInsideHorizontal: y0 = 1; cell_key = xlEdgeTop; break;
        default: return E_UNEXPECTED;
        }
        VariantInit(&row);
        VariantInit(&column);
        V_VT(&row) = V_VT(&column) = VT_I4;
        for (y = y0; y < y1; ++y) {
            for (x = x0; x < x1; ++x) {
                V_I4(&row) = y;
                V_I4(&column) = x;
                VariantInit(&cell);
                hr = AutoWrap(DISPATCH_METHOD, &cell, This->pOORange, L"getCellByPosition", 2, row, column);
                if (SUCCEEDED(hr) && (V_VT(&cell) != VT_DISPATCH || !V_DISPATCH(&cell))) hr = E_UNEXPECTED;
                if (SUCCEEDED(hr)) hr = MSO_TO_OO_BorderWeight(V_DISPATCH(&cell), cell_key, &weight, NULL);
                VariantClear(&cell);
                if (FAILED(hr)) return hr;
                if (!borders_match_weight(weight, &have_weight, &common)) goto mixed;
            }
        }
    }
    V_VT(result) = VT_I4;
    V_I4(result) = common;
    return S_OK;
mixed:
    V_VT(result) = VT_NULL;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Weight(
        I_Borders* iface, XlBorderWeight weight)
{
    I_Border *border;
    HRESULT hr;
    int key;
    for (key = xlEdgeLeft; key <= xlInsideHorizontal; ++key) {
        hr = I_Borders_get_Item(iface, key, (IDispatch**)&border);
        if (FAILED(hr)) return hr;
        hr = I_Border_put_Weight(border, weight);
        I_Border_Release(border);
        if (FAILED(hr)) return hr;
    }
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get__Default(
        I_Borders* iface,
        XlBordersIndex key,
        IDispatch **ppObject)
{
    BordersImpl *This = BORDERS_THIS(iface);
    IUnknown *punk = NULL;
    I_Border *pborder;
    HRESULT hres;
    TRACE_IN;
    TRACE("key=%08x\n",key);

    if (!ppObject) return E_POINTER;
    *ppObject = NULL;
    if (!This) return E_POINTER;
    if (key < xlLeft || key > xlInsideHorizontal) return DISP_E_BADINDEX;
    /*Создаем объект Border*/
    hres = _I_BorderConstructor((LPVOID*) &punk);

    if (FAILED(hres)) {
        ERR("create IBorder object\n");
        return E_NOINTERFACE;
    }

    hres = I_Border_QueryInterface(punk, &IID_I_Border, (void**) &pborder);
    if (FAILED(hres)) {
        ERR("QueryInterface \n");
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Border_Initialize(pborder, iface, key);

    if (FAILED(hres)) {
        ERR("Initialized \n");
        I_Border_Release(pborder);
        return hres;
    }

    *ppObject = (IDispatch*)pborder;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Item(
        I_Borders* iface,
        XlBordersIndex key,
        IDispatch **ppObject)
{
    TRACE("  ----> get__Default \n");
    return MSO_TO_OO_I_Borders_get__Default(iface, key, ppObject);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Value(
        I_Borders* iface,
        XlLineStyle *plinestyle)
{
    TRACE(" ----> get_LineStyle \n");
    return MSO_TO_OO_I_Borders_get_LineStyle(iface, plinestyle);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Value(
        I_Borders* iface,
        XlLineStyle linestyle)
{
    TRACE(" ----> put_LineStyle \n");
    return MSO_TO_OO_I_Borders_put_LineStyle(iface, linestyle);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Count(
        I_Borders* iface,
        LONG *pretval)
{
    if (!pretval) return E_POINTER;
    *pretval = BORDER_COUNT;
    return S_OK;
}

/* Each enumeration owns its cursor and retains the same range/backend. */
static HRESULT copy_borders_enumerator(BordersImpl *source, int position,
        IEnumVARIANT **result)
{
    I_Borders *borders;
    BordersImpl *copy;
    HRESULT hres;

    if (!result) return E_POINTER;
    *result = NULL;
    hres = _I_BordersConstructor((LPVOID*)&borders);
    if (FAILED(hres)) return hres;
    copy = (BordersImpl*)borders;
    copy->pRange = source->pRange;
    if (copy->pRange) I_Range_AddRef(copy->pRange);
    copy->pOORange = source->pOORange;
    if (copy->pOORange) IDispatch_AddRef(copy->pOORange);
    copy->enum_position = position;
    *result = BORDERS_ENUM(copy);
    IEnumVARIANT_AddRef(*result);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetEnumerator(
        I_Borders* iface,
        IUnknown **pdretval)
{
    BordersImpl *This = BORDERS_THIS(iface);
    IEnumVARIANT *enumerator;
    HRESULT hres;

    if (!pdretval) return E_POINTER;
    *pdretval = NULL;
    hres = copy_borders_enumerator(This, 0, &enumerator);
    if (SUCCEEDED(hres)) *pdretval = (IUnknown*)enumerator;
    return hres;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Borders_GetTypeInfoCount(
        I_Borders* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    if (!pctinfo) return E_POINTER;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetTypeInfo(
        I_Borders* iface, UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    HRESULT hr;
    if (!ppTInfo) return E_POINTER;
    *ppTInfo = NULL;
    if (iTInfo) return DISP_E_BADINDEX;
    hr = get_borders_dispatch_typeinfo(ppTInfo);
    if (SUCCEEDED(hr)) ITypeInfo_AddRef(*ppTInfo);
    return hr;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetIDsOfNames(
        I_Borders* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_borders_dispatch_typeinfo(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WERR(L"name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_Invoke(
        I_Borders* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    if (!IsEqualGUID(riid, &IID_NULL)) return DISP_E_UNKNOWNINTERFACE;
    if (!pDispParams) return E_POINTER;
    if (dispIdMember == dispid_borders_weight && (wFlags & DISPATCH_PROPERTYGET)) {
        VARIANT result;
        if (pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
        hres = borders_weight_variant(iface, &result);
        if (SUCCEEDED(hres) && pVarResult) {
            VariantInit(pVarResult);
            V_VT(pVarResult) = V_VT(&result);
            if (V_VT(&result) == VT_I4) V_I4(pVarResult) = V_I4(&result);
        }
        VariantClear(&result);
        return hres;
    }
    /* Legacy metadata describes the unchanged native vtable used for other calls. */
    hres = get_typeinfo_borders(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
    if (FAILED(hres)) {
        ERR("wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
    }
    TRACE_OUT;
    return hres;
}

const I_BordersVtbl MSO_TO_OO_I_Borders_Vtbl =
{
    MSO_TO_OO_I_Borders_QueryInterface,
    MSO_TO_OO_I_Borders_AddRef,
    MSO_TO_OO_I_Borders_Release,
    MSO_TO_OO_I_Borders_GetTypeInfoCount,
    MSO_TO_OO_I_Borders_GetTypeInfo,
    MSO_TO_OO_I_Borders_GetIDsOfNames,
    MSO_TO_OO_I_Borders_Invoke,
    MSO_TO_OO_I_Borders_get_Application,
    MSO_TO_OO_I_Borders_get_Creator,
    MSO_TO_OO_I_Borders_get_Parent,
    MSO_TO_OO_I_Borders_get_Color,
    MSO_TO_OO_I_Borders_put_Color,
    MSO_TO_OO_I_Borders_get_ColorIndex,
    MSO_TO_OO_I_Borders_put_ColorIndex,
    MSO_TO_OO_I_Borders_get_Count,
    MSO_TO_OO_I_Borders_get_Item,
    MSO_TO_OO_I_Borders_get_LineStyle,
    MSO_TO_OO_I_Borders_put_LineStyle,
    MSO_TO_OO_I_Borders_GetEnumerator,
    MSO_TO_OO_I_Borders_get_Value,
    MSO_TO_OO_I_Borders_put_Value,
    MSO_TO_OO_I_Borders_get_Weight,
    MSO_TO_OO_I_Borders_put_Weight,
    MSO_TO_OO_I_Borders_get__Default
};

#undef BORDERS_THIS

/*IEnumVARIANT interface*/

#define ENUMVAR_THIS(iface) DEFINE_THIS(BordersImpl, enumerator, iface);

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Borders_EnumVAR_AddRef(
        IEnumVARIANT* iface)
{
    BordersImpl *This = ENUMVAR_THIS(iface);
    return I_Borders_AddRef(BORDERS_BORDERS(This));
}


static HRESULT WINAPI MSO_TO_OO_I_Borders_EnumVAR_QueryInterface(
        IEnumVARIANT* iface,
        REFIID riid,
        void **ppvObject)
{
    BordersImpl *This = ENUMVAR_THIS(iface);
    return I_Borders_QueryInterface(BORDERS_BORDERS(This), riid, ppvObject);
}

static ULONG WINAPI MSO_TO_OO_I_Borders_EnumVAR_Release(
        IEnumVARIANT* iface)
{
    BordersImpl *This = ENUMVAR_THIS(iface);
    return I_Borders_Release(BORDERS_BORDERS(This));
}

/*** IEnumVARIANT methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Borders_EnumVAR_Next(
        IEnumVARIANT* iface,
        ULONG celt,
        VARIANT *rgVar,
        ULONG *pCeltFetched)
{
    BordersImpl *This = ENUMVAR_THIS(iface);
    HRESULT hres;
    ULONG l;
    LONG l1;
    LONG count;
    ULONG l2;
    IDispatch *dret;

    if (pCeltFetched != NULL)
       *pCeltFetched = 0;

    if (rgVar == NULL)
       return E_INVALIDARG;

    /*Init Array*/
    for (l=0; l<celt; l++)
       VariantInit(&rgVar[l]);

    hres = I_Borders_get_Count(BORDERS_BORDERS(This), &count);
    if (FAILED(hres)) goto error;

    for (l1=This->enum_position, l2=0; l1<count && l2<celt; l1++, l2++) {
      hres = I_Borders_get_Item(BORDERS_BORDERS(This), border_indices[l1], &dret);
      if (FAILED(hres)) goto error;
      V_VT(&rgVar[l2]) = VT_DISPATCH;
      V_DISPATCH(&rgVar[l2]) = dret;
    }

    if (pCeltFetched != NULL)
       *pCeltFetched = l2;

   This->enum_position = l1;

   return  (l2 < celt) ? S_FALSE : S_OK;

error:
   for (l=0; l<celt; l++)
      VariantClear(&rgVar[l]);
   return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_EnumVAR_Skip(
        IEnumVARIANT* iface,
        ULONG celt)
{
    BordersImpl *This = ENUMVAR_THIS(iface);
    LONG count;
    ULONG remaining;
    HRESULT hres;

    hres = I_Borders_get_Count(BORDERS_BORDERS(This), &count);
    if (FAILED(hres)) return hres;
    remaining = count - This->enum_position;
    if (celt > remaining) {
        This->enum_position = count;
        return S_FALSE;
    }
    This->enum_position += celt;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_EnumVAR_Reset(
        IEnumVARIANT* iface)
{
    BordersImpl *This = ENUMVAR_THIS(iface);
    This->enum_position = 0;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_EnumVAR_Clone(
        IEnumVARIANT* iface,
        IEnumVARIANT **ppEnum)
{
    BordersImpl *This = ENUMVAR_THIS(iface);
    return copy_borders_enumerator(This, This->enum_position, ppEnum);
}

#undef ENUMVAR_THIS

const IEnumVARIANTVtbl MSO_TO_OO_I_Borders_enumvarVtbl =
{
    MSO_TO_OO_I_Borders_EnumVAR_QueryInterface,
    MSO_TO_OO_I_Borders_EnumVAR_AddRef,
    MSO_TO_OO_I_Borders_EnumVAR_Release,
    MSO_TO_OO_I_Borders_EnumVAR_Next,
    MSO_TO_OO_I_Borders_EnumVAR_Skip,
    MSO_TO_OO_I_Borders_EnumVAR_Reset,
    MSO_TO_OO_I_Borders_EnumVAR_Clone
};


extern HRESULT _I_BordersConstructor(LPVOID *ppObj)
{
    BordersImpl *borders;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    borders = HeapAlloc(GetProcessHeap(), 0, sizeof(*borders));
    if (!borders)
    {
        return E_OUTOFMEMORY;
    }

    borders->pbordersVtbl = &MSO_TO_OO_I_Borders_Vtbl;
    borders->penumeratorVtbl = &MSO_TO_OO_I_Borders_enumvarVtbl;
    borders->ref = 0;
    borders->pRange = NULL;
    borders->enum_position = 0;
    borders->pOORange = NULL;
             
    *ppObj = BORDERS_BORDERS(borders);
    
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}
