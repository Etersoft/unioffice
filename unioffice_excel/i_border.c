/*
 * IBorder interface functions
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

ITypeInfo *ti_border = NULL;

HRESULT get_typeinfo_border(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_border) {
        *typeinfo = ti_border;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Border, &ti_border);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_border;
    return hres;
}


    /*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Border_AddRef(
        I_Border* iface)
{
    BorderImpl *This = (BorderImpl*)iface;
    ULONG ref;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_QueryInterface(
        I_Border* iface,
        REFIID riid,
        void **ppvObject)
{
    BorderImpl *This = (BorderImpl*)iface;

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Border)) {
        *ppvObject = &This->_borderVtbl;
        MSO_TO_OO_I_Border_AddRef(iface);

        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Border_Release(
        I_Border* iface)
{
    BorderImpl *This = (BorderImpl*)iface;
    ULONG ref;
    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->pborders!=NULL) {
            IDispatch_Release(This->pborders);
            This->pborders = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*** I_Border methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Border_get_Application(
        I_Border* iface,
        IDispatch **value)
{
    BorderImpl *This = (BorderImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    BordersImpl *borders = (BordersImpl*)(This->pborders);
    TRACE("\n");

    if (borders==NULL) return E_POINTER;
    if (borders->prange==NULL) return E_POINTER;
    TRACE_OUT;
    return I_Range_get_Application((I_Range*)(borders->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_Parent(
        I_Border* iface,
        IDispatch **value)
{
    BorderImpl *This = (BorderImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    *value = This->pborders;
    I_Range_AddRef(This->pborders);

    if (value==NULL)
        return E_POINTER;

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_Color(
        I_Border* iface,
        long *plcolor)
{
    BorderImpl *This = (BorderImpl*)iface;
    HRESULT hres;
    VARIANT vret;
    VARIANT X,Y;
    BordersImpl *borders = (BordersImpl*)(This->pborders);
    RangeImpl *cur_range = (RangeImpl*)borders->prange;
    TRACE_IN;

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &X, cur_range->pOORange, L"TableBorder", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get TableBorder\n");
        return E_FAIL;
    }

    switch(This->key) {
/*    case xlLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"LeftBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftBorder\n");
            return E_FAIL;
        }
        break;
    case xlRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"RightBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightBorder\n");
            return E_FAIL;
        }
        break;
    case xlTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"TopBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopBorder\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"BottomBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomBorder\n");
            return E_FAIL;
        }
        break;*/
    case xlLeft:
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE(" ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlRight:
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlTop:
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
    case xlEdgeBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"BottomLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideHorizontal:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"HorizontalLine", 0);
        if (FAILED(hres)) {
            TRACE(" ERROR when get HorizontalLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideVertical:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"VerticalLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get VerticalLine\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalDown:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalTLBR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalTLBR\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalUp:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalBLTR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalBLTR\n");
            return E_FAIL;
        }
        break;
    default :
        TRACE("NOT REALIZE %i \n",This->key);
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vret, V_DISPATCH(&Y), L"Color", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when Color \n");
        return E_FAIL;
    }
    *plcolor=V_I4(&vret);

    VariantClear(&vret);
    VariantClear(&X);
    VariantClear(&Y);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_put_Color(
        I_Border* iface,
        long lcolor)
{
    BorderImpl *This = (BorderImpl*)iface;
    HRESULT hres;
    VARIANT vret;
    VARIANT X,Y,param1;
    BordersImpl *borders = (BordersImpl*)(This->pborders);
    RangeImpl *cur_range = (RangeImpl*)borders->prange;
    TRACE_IN;

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);
    VariantInit(&param1);

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }
    /*Делаем границу видимой*/
    I_Border_put_LineStyle(iface, xlContinuous);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &X, cur_range->pOORange, L"TableBorder", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get TableBorder\n");
        return E_FAIL;
    }

    switch(This->key) {
 /*   case xlLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"LeftBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftBorder\n");
            return E_FAIL;
        }
        break;
    case xlRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"RightBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightBorder\n");
            return E_FAIL;
        }
        break;
    case xlTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"TopBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopBorder\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"BottomBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomBorder\n");
            return E_FAIL;
        }
        break;*/
    case xlLeft:
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlRight:
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlTop:
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
    case xlEdgeBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"BottomLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideHorizontal:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"HorizontalLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get HorizontalLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideVertical:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"VerticalLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get VerticalLine\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalDown:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalTLBR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalTLBR\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalUp:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalBLTR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalBLTR\n");
            return E_FAIL;
        }
        break;
    default :
        TRACE("NOT REALIZE %i \n",This->key);
    }

    V_VT(&param1) = VT_I4;
    V_I4(&param1) = lcolor;
    VariantClear(&vret);

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&Y), L"Color", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when Color \n");
        return E_FAIL;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = V_DISPATCH(&Y);
    IDispatch_AddRef(V_DISPATCH(&param1));

    switch(This->key) {
/*    case xlLeft:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"LeftBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put LeftBorder\n");
            return E_FAIL;
        }
        break;
    case xlRight:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"RightBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put RightBorder\n");
            return E_FAIL;
        }
        break;
    case xlTop:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"TopBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put TopBorder\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"BottomBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put BottomBorder\n");
            return E_FAIL;
        }
        break;*/
    case xlLeft:
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"LeftLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlRight:
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"RightLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlTop:
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"TopLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
    case xlEdgeBottom:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"BottomLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideHorizontal:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"HorizontalLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get HorizontalLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideVertical:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"VerticalLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get VerticalLine\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalDown:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"DiagonalTLBR", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put DiagonalTLBR\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalUp:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"DiagonalBLTR", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put DiagonalBLTR\n");
            return E_FAIL;
        }
        break;
    default :
        TRACE("NOT REALIZE %i \n",This->key);
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = V_DISPATCH(&X);
    IDispatch_AddRef(V_DISPATCH(&param1));
    VariantClear(&vret);

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"TableBorder", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when get TableBorder\n");
        return E_FAIL;
    }

    VariantClear(&vret);
    VariantClear(&X);
    VariantClear(&Y);
    VariantClear(&param1);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_ColorIndex(
        I_Border* iface,
        long *plcolorindex)
{
    BorderImpl *This = (BorderImpl*)iface;
    long tmpcolor;
    int i;
    HRESULT hres;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    hres = MSO_TO_OO_I_Border_get_Color(iface,&tmpcolor);
    if (FAILED(hres)) {
        return hres;
    }
    for (i=0;i<56;i++)
        if (color[i]==tmpcolor) {
            *plcolorindex = i+1;
            return S_OK;
        }

    TRACE("ERROR Color don`t have colorindex \n");
    *plcolorindex = 1;/*белый цвет*/
    /*Отправляем что все хорошо, на всякий случай*/
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_put_ColorIndex(
        I_Border* iface,
        long lcolorindex)
{
    BorderImpl *This = (BorderImpl*)iface;
    long tmpcolor;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;

    if ((lcolorindex<1)||(lcolorindex>56)) {
        TRACE("ERROR Incorrect colorindex %i\n", lcolorindex);
        TRACE_OUT;
        return S_OK;
    } else {
        TRACE_OUT;
        return MSO_TO_OO_I_Border_put_Color(iface,color[lcolorindex-1]);
    }
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_Creator(
        I_Border* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_LineStyle(
        I_Border* iface,
        XlLineStyle *plinestyle)
{
    BorderImpl *This = (BorderImpl*)iface;
    HRESULT hres;
    VARIANT vret;
    VARIANT X,Y;
    long in,out,l;
    BordersImpl *borders = (BordersImpl*)(This->pborders);
    RangeImpl *cur_range = (RangeImpl*)borders->prange;
    TRACE_IN;

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &X, cur_range->pOORange, L"TableBorder", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get TableBorder\n");
        return E_FAIL;
    }

    switch(This->key) {
/*    case xlLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"LeftBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftBorder\n");
            return E_FAIL;
        }
        break;
    case xlRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"RightBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightBorder\n");
            return E_FAIL;
        }
        break;
    case xlTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"TopBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopBorder\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"BottomBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomBorder\n");
            return E_FAIL;
        }
        break;*/
    case xlLeft:
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlRight:
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlTop:
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
    case xlEdgeBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"BottomLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideHorizontal:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"HorizontalLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get HorizontalLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideVertical:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"VerticalLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get VerticalLine\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalDown:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalTLBR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalTLBR\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalUp:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalBLTR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalBLTR\n");
            return E_FAIL;
        }
        break;
    default :
        TRACE("NOT REALIZE %i \n",This->key);
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vret, V_DISPATCH(&Y), L"InnerLineWidth", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when InnerLineWidth \n");
        return E_FAIL;
    }
    hres = VariantChangeTypeEx(&vret, &vret, 0, 0, VT_I4);
    in=V_I4(&vret);
    VariantClear(&vret);
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vret, V_DISPATCH(&Y), L"OuterLineWidth", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when OuterLineWidth \n");
        return E_FAIL;
    }
    hres = VariantChangeTypeEx(&vret, &vret, 0, 0, VT_I4);
    out=V_I4(&vret);
    VariantClear(&vret);
    hres = AutoWrap(DISPATCH_PROPERTYGET, &vret, V_DISPATCH(&Y), L"LineDistance", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when LineDistance \n");
        return E_FAIL;
    }
    hres = VariantChangeTypeEx(&vret, &vret, 0, 0, VT_I4);
    l=V_I4(&vret);

    if ((l==0)&&(out==0)&&(in==0)) *plinestyle=xlLineStyleNone;
        else if ((l==0)&&(out==0)) *plinestyle=xlContinuous;
            else *plinestyle=xlDouble;

    VariantClear(&vret);
    VariantClear(&X);
    VariantClear(&Y);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_put_LineStyle(
        I_Border* iface,
        XlLineStyle linestyle)
{
    BorderImpl *This = (BorderImpl*)iface;
    HRESULT hres;
    VARIANT vret;
    VARIANT X,Y,param1;
    BordersImpl *borders = (BordersImpl*)(This->pborders);
    RangeImpl *cur_range = (RangeImpl*)borders->prange;
    TRACE_IN;

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);
    VariantInit(&param1);
    long in,out,l;

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    switch(linestyle) {
    case xlContinuous:
        in = 10;
        out = 0;
        l = 0;
        break;
    case xlDouble:
        in = 10;
        out = 10;
        l = 10;
        break;
    case xlLineStyleNone:
        in = 0;
        out = 0;
        l = 0;
        break;
    case xlDash:
    case xlDashDot:
    case xlDashDotDot:
    case xlDot:
    case xlSlantDashDot:
        TRACE("NOT IMPLEMENTED\n");
        return E_NOTIMPL;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &X, cur_range->pOORange, L"TableBorder", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get TableBorder\n");
        return E_FAIL;
    }

    switch(This->key) {
 /*   case xlLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"LeftBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftBorder\n");
            return E_FAIL;
        }
        break;
    case xlRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"RightBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightBorder\n");
            return E_FAIL;
        }
        break;
    case xlTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"TopBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopBorder\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"BottomBorder", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomBorder\n");
            return E_FAIL;
        }
        break;*/
    case xlLeft:
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlRight:
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlTop:
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
    case xlEdgeBottom:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"BottomLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideHorizontal:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"HorizontalLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get HorizontalLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideVertical:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"VerticalLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get VerticalLine\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalDown:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalTLBR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalTLBR\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalUp:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, cur_range->pOORange, L"DiagonalBLTR", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get DiagonalBLTR\n");
            return E_FAIL;
        }
        break;
    default :
        TRACE("NOT REALIZE %i \n",This->key);
    }

    V_VT(&param1) = VT_I4;
    V_I4(&param1) = in;
    VariantClear(&vret);

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&Y), L"InnerLineWidth", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when InnerLineWidth \n");
        return E_FAIL;
    }
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = out;
    VariantClear(&vret);

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&Y), L"OuterLineWidth", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when OuterLineWidth \n");
        return E_FAIL;
    }
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = out;
    VariantClear(&vret);

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&Y), L"LineDistance", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when LineDistance \n");
        return E_FAIL;
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = V_DISPATCH(&Y);
    IDispatch_AddRef(V_DISPATCH(&param1));

    switch(This->key) {
/*    case xlLeft:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"LeftBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put LeftBorder\n");
            return E_FAIL;
        }
        break;
    case xlRight:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"RightBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put RightBorder\n");
            return E_FAIL;
        }
        break;
    case xlTop:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"TopBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put TopBorder\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"BottomBorder", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put BottomBorder\n");
            return E_FAIL;
        }
        break;*/
    case xlLeft:
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"LeftLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlRight:
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"RightLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlTop:
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"TopLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
    case xlBottom:
    case xlEdgeBottom:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"BottomLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get BottomLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideHorizontal:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"HorizontalLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get HorizontalLine\n");
            return E_FAIL;
        }
        break;
    case xlInsideVertical:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"VerticalLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get VerticalLine\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalDown:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"DiagonalTLBR", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put DiagonalTLBR\n");
            return E_FAIL;
        }
        break;
    case xlDiagonalUp:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"DiagonalBLTR", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when put DiagonalBLTR\n");
            return E_FAIL;
        }
        break;
    default :
        TRACE("NOT REALIZE %i \n",This->key);
    }

    VariantClear(&param1);
    V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = V_DISPATCH(&X);
    IDispatch_AddRef(V_DISPATCH(&param1));
    VariantClear(&vret);

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"TableBorder", 1, param1);
    if (FAILED(hres)) {
        TRACE("ERROR when get TableBorder\n");
        return E_FAIL;
    }

    VariantClear(&vret);
    VariantClear(&X);
    VariantClear(&Y);
    VariantClear(&param1);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_Weight(
        I_Border* iface,
        XlBorderWeight *pweight)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_put_Weight(
        I_Border* iface,
        XlBorderWeight weight)
{
    TRACE_NOTIMPL;
    /*FIXME to realize*/
//    I_Border_put_LineStyle(iface, xlContinuous);
    return S_OK;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Border_GetTypeInfoCount(
        I_Border* iface,
        UINT *pctinfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_GetTypeInfo(
        I_Border* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_GetIDsOfNames(
        I_Border* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_border(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_Invoke(
        I_Border* iface,
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

    hres = get_typeinfo_border(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams,
                            pVarResult, pExcepInfo, puArgErr);
    if (FAILED(hres)) {
        TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
    }

    TRACE_OUT;
    return hres;
}


const I_BorderVtbl MSO_TO_OO_I_Border_Vtbl =
{
    MSO_TO_OO_I_Border_QueryInterface,
    MSO_TO_OO_I_Border_AddRef,
    MSO_TO_OO_I_Border_Release,
    MSO_TO_OO_I_Border_GetTypeInfoCount,
    MSO_TO_OO_I_Border_GetTypeInfo,
    MSO_TO_OO_I_Border_GetIDsOfNames,
    MSO_TO_OO_I_Border_Invoke,
    MSO_TO_OO_I_Border_get_Application,
    MSO_TO_OO_I_Border_get_Creator,
    MSO_TO_OO_I_Border_get_Parent,
    MSO_TO_OO_I_Border_get_Color,
    MSO_TO_OO_I_Border_put_Color,
    MSO_TO_OO_I_Border_get_ColorIndex,
    MSO_TO_OO_I_Border_put_ColorIndex,
    MSO_TO_OO_I_Border_get_LineStyle,
    MSO_TO_OO_I_Border_put_LineStyle,
    MSO_TO_OO_I_Border_get_Weight,
    MSO_TO_OO_I_Border_put_Weight
};

extern HRESULT _I_BorderConstructor(LPVOID *ppObj)
{
    BorderImpl *border;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    border = HeapAlloc(GetProcessHeap(), 0, sizeof(*border));
    if (!border)
    {
        return E_OUTOFMEMORY;
    }

    border->_borderVtbl = &MSO_TO_OO_I_Border_Vtbl;
    border->ref = 0;
    border->key = 0;
    border->pborders = NULL;

    *ppObj = &border->_borderVtbl;
    TRACE_OUT;
    return S_OK;
}

