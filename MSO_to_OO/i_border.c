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

static WCHAR const str_application[] = {
    'A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR const str_parent[] = {
    'P','a','r','e','n','t',0};
static WCHAR const str_color[] = {
    'C','o','l','o','r',0};
static WCHAR const str_colorindex[] = {
    'C','o','l','o','r','I','n','d','e','x',0};
static WCHAR const str_creator[] = {
    'C','r','e','a','t','o','r',0};
static WCHAR const str_linestyle[] = {
    'L','i','n','e','S','t','y','l','e',0};
static WCHAR const str_weight[] = {
    'W','e','i','g','h','t',0};


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

    TRACE("\n");

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

    if (This==NULL) return E_POINTER;

    BordersImpl *borders = (BordersImpl*)(This->pborders);
    TRACE("\n");

    if (borders==NULL) return E_POINTER;
    if (borders->prange==NULL) return E_POINTER;

    return I_Range_get_Application((I_Range*)(borders->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_Parent(
        I_Border* iface,
        IDispatch **value)
{
    BorderImpl *This = (BorderImpl*)iface;

    TRACE(" \n");

    if (This==NULL) return E_POINTER;

    *value = This->pborders;
    I_Range_AddRef(This->pborders);

    if (value==NULL)
        return E_POINTER;

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

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);

    TRACE(" \n");

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
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE(" ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
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

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);
    VariantInit(&param1);

    TRACE("\n");

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }
    /*������ ������� �������*/
    I_Border_put_LineStyle(iface, xlContinuous);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &X, cur_range->pOORange, L"TableBorder", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when get TableBorder\n");
        return E_FAIL;
    }

    switch(This->key) {
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
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
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"LeftLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"RightLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"TopLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
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
    TRACE("\n");

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
    *plcolorindex = 1;/*����� ����*/
    /*���������� ��� ��� ������, �� ������ ������*/
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_put_ColorIndex(
        I_Border* iface,
        long lcolorindex)
{
    BorderImpl *This = (BorderImpl*)iface;
    long tmpcolor;
    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;

    if ((lcolorindex<1)||(lcolorindex>56)) {
        TRACE("ERROR Incorrect colorindex %i\n", lcolorindex);
        return S_OK;
    } else 
        return MSO_TO_OO_I_Border_put_Color(iface,color[lcolorindex-1]);
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_Creator(
        I_Border* iface,
        VARIANT *result)
{
    TRACE(" \n");
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
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

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);

    TRACE("\n");

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
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
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

    VariantInit(&vret);
    VariantInit(&X);
    VariantInit(&Y);
    VariantInit(&param1);
    long in,out,l;

    TRACE(" \n");

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
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"LeftLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"RightLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYGET, &Y, V_DISPATCH(&X), L"TopLine", 0);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
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
    case xlEdgeLeft:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"LeftLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get LeftLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeRight:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"RightLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get RightLine\n");
            return E_FAIL;
        }
        break;
    case xlEdgeTop:
        hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, V_DISPATCH(&X), L"TopLine", 1, param1);
        if (FAILED(hres)) {
            TRACE("ERROR when get TopLine\n");
            return E_FAIL;
        }
        break;
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

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_get_Weight(
        I_Border* iface,
        XlBorderWeight *pweight)
{
    TRACE("NOT IMPL\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_put_Weight(
        I_Border* iface,
        XlBorderWeight weight)
{
    TRACE("NOT IMPL \n");
    return E_NOTIMPL;
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Border_GetTypeInfoCount(
        I_Border* iface,
        UINT *pctinfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Border_GetTypeInfo(
        I_Border* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE(" \n");
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
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = 1;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_parent)) {
        *rgDispId = 2;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_color)) {
        *rgDispId = 3;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_colorindex)) {
        *rgDispId = 4;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_creator)) {
        *rgDispId = 5;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_linestyle)) {
        *rgDispId = 6;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_weight)) {
        *rgDispId = 7;
        return S_OK;
    }
    /*������� �������� ������ ��� ��������,
    ����� ����� ���� �� �������.*/
    WTRACE(L"%s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
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
    HRESULT hres;
    IDispatch *dret;
    VARIANT vval;
    long ltmp;
    TRACE(" dispIdMember = %i\n", dispIdMember);

    VariantInit(&vval);

    if (iface==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }

    switch(dispIdMember)
    {
    case 1://Application
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Border_get_Application(iface,&dret);
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
    case 2://Parent
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Border_get_Parent(iface,&dret);
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
    case 3://Color
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vval, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);

            if (FAILED(hres)) {
                TRACE(" (case 3) ERROR VariantChangeTypeEx   %08x   VT = %i\n",hres,V_VT(&(pDispParams->rgvarg[0])));
                return E_FAIL;
            }
            ltmp = V_I4(&vval);
            hres = MSO_TO_OO_I_Border_put_Color(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Border_get_Color(iface,&ltmp);
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
    case 4://ColorIndex
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vval, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (case 4) ERROR VariantChangeTypeEx   %08x\n",hres);
                return E_FAIL;
            }
            ltmp = V_I4(&vval);
            hres = MSO_TO_OO_I_Border_put_ColorIndex(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Border_get_ColorIndex(iface,&ltmp);
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
    case 5://creator
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Border_get_Creator(iface, &vval);
            if (pVarResult!=NULL){
                *pVarResult = vval;
            }
            return hres;
        }
    case 6://linestyle
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vval, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);

            if (FAILED(hres)) {
                TRACE(" (case 3) ERROR VariantChangeTypeEx   %08x   VT = %i\n",hres,V_VT(&(pDispParams->rgvarg[0])));
                return E_FAIL;
            }
            ltmp = V_I4(&vval);
            hres = MSO_TO_OO_I_Border_put_LineStyle(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Border_get_LineStyle(iface,(XlLineStyle*)&ltmp);
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

    case 7://weight
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vval, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);

            if (FAILED(hres)) {
                TRACE(" (case 3) ERROR VariantChangeTypeEx   %08x   VT = %i\n",hres,V_VT(&(pDispParams->rgvarg[0])));
                return E_FAIL;
            }
            ltmp = V_I4(&vval);
            hres = MSO_TO_OO_I_Border_put_Weight(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Border_get_Weight(iface,(XlBorderWeight*)&ltmp);
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

    }

    TRACE(" dispIdMember = %i NOT REALIZE\n",dispIdMember);
    return E_NOTIMPL;
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
    MSO_TO_OO_I_Border_get_Parent,
    MSO_TO_OO_I_Border_get_Color,
    MSO_TO_OO_I_Border_put_Color,
    MSO_TO_OO_I_Border_get_ColorIndex,
    MSO_TO_OO_I_Border_put_ColorIndex,
    MSO_TO_OO_I_Border_get_Creator,
    MSO_TO_OO_I_Border_get_LineStyle,
    MSO_TO_OO_I_Border_put_LineStyle,
    MSO_TO_OO_I_Border_get_Weight,
    MSO_TO_OO_I_Border_put_Weight
};

BorderImpl MSO_TO_OO_Border =
{
    &MSO_TO_OO_I_Border_Vtbl,
    0,
    NULL,
    0
};

extern HRESULT _I_BorderConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    BorderImpl *border;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);

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

    return S_OK;
}

