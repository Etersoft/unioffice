/*
 * IFont interface functions
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
#include <oleauto.h>


static WCHAR const str_bold[] = {
    'B','o','l','d',0};
static WCHAR const str_italic[] = {
    'I','t','a','l','i','c',0};
static WCHAR const str_underline[] = {
    'U','n','d','e','r','l','i','n','e',0};
static WCHAR const str_size[] = {
    'S','i','z','e',0};
static WCHAR const str_strikethrought[] = {
    'S','t','r','i','k','e','t','h','r','o','u','g','h',0};
static WCHAR const str_name[] = {
    'N','a','m','e',0};
static WCHAR const str_colorindex[] = {
    'C','o','l','o','r','I','n','d','e','x',0};
static WCHAR const str_color[] = {
    'C','o','l','o','r',0};
static WCHAR const str_application[] = {
    'A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR const str_parent[] = {
    'P','a','r','e','n','t',0};
static WCHAR const str_creator[] = {
    'C','r','e','a','t','o','r',0};
static WCHAR const str_shadow[] = {
    'S','h','a','d','o','w',0};
static WCHAR const str_subscript[] = {
    'S','u','b','s','c','r','i','p','t',0};
static WCHAR const str_superscript[] = {
    'S','u','p','e','r','s','c','r','i','p','t',0};

#define usNONE 0
#define usSINGLE 1
#define usDOUBLE 2
#define usDOTTED 3
#define usDONTKNOW 4
#define usDASH 5
#define usLONGDASH 6
#define usDASHDOT 7
#define usDASHDOTDOT 8
#define usSMALLWAVE 9
#define usWAVE 10
#define usDOUBLEWAVE 11
#define usBOLD 12
#define usBOLDDOTTED 13
#define usBOLDDASH 14
#define usBOLDLONGDASH 15
#define usBOLDDASHDOT 16
#define usBOLDDASHDOTDOT 17
#define usBOLDWAVE 18

/*IUnknown*/
static ULONG WINAPI MSO_TO_OO_I_Font_AddRef(
        I_Font* iface)
{
    _FontImpl *This = (_FontImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_QueryInterface(
        I_Font* iface,
        REFIID riid,
        void **ppvObject)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Font)) {
        *ppvObject = &This->_ifontVtbl;
        MSO_TO_OO_I_Font_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}


static ULONG WINAPI MSO_TO_OO_I_Font_Release(
        I_Font* iface)
{
    _FontImpl *This = (_FontImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
        if (This->prange!=NULL) {
            IDispatch_Release(This->prange);
            This->prange = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
    }
    return ref;
}

/*I_Font methods*/

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Bold(
        I_Font* iface,
        VARIANT_BOOL *pvbBold)
{
    _FontImpl *This = (_FontImpl*)iface;

    /*In OO bold is specified as weight of the character*/
    VARIANT vBoldState;
    VariantInit (&vBoldState);

    TRACE("\n");

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vBoldState, range->pOORange, L"CharWeight", 0);
    if (hres != S_OK)  {
        *pvbBold = VARIANT_FALSE;
        return hres;
    }
    if (V_R4(&vBoldState) > 140)
        *pvbBold = VARIANT_TRUE;
    else
        *pvbBold = VARIANT_FALSE;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Bold(
        I_Font* iface,
        VARIANT_BOOL vbBold)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT vBoldState;
    VariantInit (&vBoldState);
    V_VT(&vBoldState) = VT_R4;

    if (vbBold == VARIANT_TRUE)
        V_R4(&vBoldState) = 200;
    else
        V_R4(&vBoldState) = 100;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharWeight", 1, vBoldState);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Italic(
        I_Font* iface,
        VARIANT_BOOL *pvbItalic)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT vItalicState;
    VariantInit (&vItalicState);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vItalicState, range->pOORange, L"CharPosture", 0);
    if (hres != S_OK)
        return hres;
    if (V_I2(&vItalicState) != 0)  /* !!!!! V_INT */
        *pvbItalic = VARIANT_TRUE;
    else
        *pvbItalic = VARIANT_FALSE;
    return S_OK;

}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Italic(
        I_Font* iface,
        VARIANT_BOOL vbItalic)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT vItalicState;
    VariantInit (&vItalicState);
    V_VT(&vItalicState) = VT_I2; /* !!!!! VT_INT */

    if (vbItalic == VARIANT_TRUE)
        V_I2(&vItalicState) = 1;
    else
        V_I2(&vItalicState) = 0;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharPosture", 1, vItalicState);

    return S_OK;
}
/* TODO 1 - нет подчеркивания 2-есть подчеркивание*/
static HRESULT WINAPI MSO_TO_OO_I_Font_get_Underline(
        I_Font* iface,
        VARIANT *pvbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, range->pOORange, L"CharUnderline", 0);

    if (hres != S_OK)
       return hres;

    V_VT(pvbUnderline) = VT_I4;

    hres = VariantChangeTypeEx(&vUnderlineState, &vUnderlineState, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }

    switch(V_I4(&vUnderlineState)) {
        case usSINGLE:
            V_I4(pvbUnderline) = xlUnderlineStyleSingle;
            break;
        case usDOUBLE:
            V_I4(pvbUnderline) = xlUnderlineStyleDouble;
            break;
        case usNONE:
            V_I4(pvbUnderline) = xlUnderlineStyleNone;
            break;
        default:
            TRACE("ERROR CharUnderline \n");
            return E_FAIL;
    }

    return S_OK;
}
/* TODO 1 - нет подчеркивания 2-есть подчеркивание*/
static HRESULT WINAPI MSO_TO_OO_I_Font_put_Underline(
        I_Font* iface,
        VARIANT vbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;

    TRACE("\n");

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    hres = VariantChangeTypeEx(&vbUnderline, &vbUnderline, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }

    V_VT(&vUnderlineState) = VT_I4;

    switch (V_I4(&vbUnderline)) {
        case xlUnderlineStyleDouble:
        case xlUnderlineStyleDoubleAccounting:
        V_I4(&vUnderlineState) = usDOUBLE;
        break;
        case xlUnderlineStyleNone:
        V_I4(&vUnderlineState) = usNONE;
        break;
        case xlUnderlineStyleSingle:
        case xlUnderlineStyleSingleAccounting:
        V_I4(&vUnderlineState) = usSINGLE;
        break;
    default :
      TRACE("ERROR parameters \n");
      return E_FAIL;
    }

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharUnderline", 1, vUnderlineState);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Size(
        I_Font* iface,
        long *plsize)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT vsize;
    VariantInit (&vsize);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vsize, range->pOORange, L"CharHeight", 0);

    if (hres != S_OK)
        return hres;

    hres = VariantChangeTypeEx(&vsize,&vsize,0,0,VT_I4);
    if (FAILED(hres)) {
        TRACE("Error when VariantChangeTypeEx\n");
    }
    *plsize = V_I4(&vsize);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Size(
        I_Font* iface,
        long lsize)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE(" %i \n",lsize);

    VARIANT vsize;
    VariantInit (&vsize);

    V_VT(&vsize) = VT_I4;
    V_I4(&vsize) = lsize;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharHeight", 1, vsize);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Strikethrough(
        I_Font* iface,
        VARIANT_BOOL *pvbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, range->pOORange, L"CharStrikeout", 0);

    if (hres != S_OK)
        return hres;
    *pvbUnderline = V_BOOL(&vUnderlineState);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Strikethrough(
        I_Font* iface,
        VARIANT_BOOL vbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE(" \n");

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    V_VT(&vUnderlineState) = VT_BOOL;
    V_BOOL(&vUnderlineState) = vbUnderline;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharStrikeout", 1, vUnderlineState);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Name(
        I_Font* iface,
        VARIANT *vName)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, vName, range->pOORange, L"CharFontName", 0);

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Name(
        I_Font* iface,
        VARIANT vName)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharFontName", 1, vName);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Color(
        I_Font* iface,
        long *plcolor)
{
    InteriorImpl *This = (InteriorImpl*)iface;
    HRESULT hres;
    VARIANT vret;
    VariantInit(&vret);

    TRACE("\n");

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    RangeImpl *cur_range = (RangeImpl*)(I_Range*)(This->prange);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vret, cur_range->pOORange, L"CharColor", 0);

    if (FAILED(hres)) {
        TRACE("ERROR when CharColor");
    }

    hres = VariantChangeTypeEx(&vret, &vret, 0, 0, VT_I4);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
    return E_FAIL;
    }
    *plcolor = V_I4(&vret);
    TRACE(" lcolor=%i\n",*plcolor);
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Color(
        I_Font* iface,
        long lcolor)
{
    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;
    VARIANT vret,param1;

    TRACE(" lcolor = %i\n",lcolor);

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    RangeImpl *cur_range = (RangeImpl*)((I_Range*)(This->prange));

    VariantInit(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = lcolor;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"CharColor", 1, param1);

    if (FAILED(hres)) TRACE("ERROR when CharColor");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_ColorIndex(
        I_Font* iface,
        long *plcolorindex)
{
    _FontImpl *This = (_FontImpl*)iface;
    long tmpcolor;
    int i;
    HRESULT hres;
    TRACE("\n");

    if (This==NULL) return E_POINTER;


    hres = MSO_TO_OO_I_Font_get_Color(iface,&tmpcolor);
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
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_ColorIndex(
        I_Font* iface,
        long lcolorindex)
{
    _FontImpl *This = (_FontImpl*)iface;
    long tmpcolor;
    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;

    if ((lcolorindex<1)||(lcolorindex>56)) {
        TRACE("ERROR Incorrect colorindex \n");
        return S_OK;
    } else 
        return MSO_TO_OO_I_Font_put_Color(iface,color[lcolorindex-1]);
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Application(
        I_Font* iface,
        IDispatch **value)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    return I_Range_get_Application((I_Range*)(This->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Parent(
        I_Font* iface,
        IDispatch **value)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if (value==NULL)
        return E_POINTER;

    *value = This->prange;
    I_Range_AddRef(This->prange);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Creator(
        I_Font* iface,
        VARIANT *result)
{
    TRACE("\n");
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Shadow(
        I_Font* iface,
        VARIANT_BOOL *pvbshadow)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE("\n");

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, range->pOORange, L"CharShadowed", 0);

    if (hres != S_OK)
        return hres;
    *pvbshadow = V_BOOL(&vUnderlineState);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Shadow(
        I_Font* iface,
        VARIANT_BOOL vbshadow)
{
    _FontImpl *This = (_FontImpl*)iface;

    TRACE(" \n");

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    V_VT(&vUnderlineState) = VT_BOOL;
    V_BOOL(&vUnderlineState) = vbshadow;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharShadowed", 1, vUnderlineState);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Background(
        I_Font* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Background(
        I_Font* iface,
        VARIANT RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_FontStyle(
        I_Font* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_FontStyle(
        I_Font* iface,
        VARIANT RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_OutlineFont(
        I_Font* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_OutlineFont(
        I_Font* iface,
        VARIANT RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Subscript(
        I_Font* iface,
        VARIANT *RHS)
{
/*    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;
    VARIANT res;
    RangeImpl *range = (RangeImpl*)This->prange;
    VARIANT vsubscript;

    TRACE(" \n");

    VariantInit (&vsubscript);
    VariantInit (&res);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &res, range->pOORange, L"CharEscapement", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when put CharEscapement\n");
        return E_FAIL;
    }

    hres = VariantChangeTypeEx(&res, &res, 0, 0, VT_I2);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }
    if (V_I2(&res)<0) {
        V_VT(RHS) = VT_BOOL;
        V_BOOL(RHS) = VARIANT_TRUE;
    } else {
        V_VT(RHS) = VT_BOOL;
        V_BOOL(RHS) = VARIANT_FALSE;
    }

    return S_OK;*/
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Subscript(
        I_Font* iface,
        VARIANT RHS)
{
    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;
    VARIANT res;
    RangeImpl *range = (RangeImpl*)This->prange;
    VARIANT vsubscript;

    TRACE(" \n");

    VariantInit (&vsubscript);
    VariantInit (&res);

    hres = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }
    if (V_BOOL(&RHS)==VARIANT_TRUE) {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = -50;
    } else {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = 0;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharEscapement", 1, vsubscript);
    if (FAILED(hres)) {
        TRACE("ERROR when put CharEscapement\n");
        return E_FAIL;
    }

    return S_OK;
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Superscript(
        I_Font* iface,
        VARIANT *RHS)
{
/*    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;
    VARIANT res;
    RangeImpl *range = (RangeImpl*)This->prange;
    VARIANT vsubscript;

    TRACE(" \n");

    VariantInit (&vsubscript);
    VariantInit (&res);

    hres = AutoWrap(DISPATCH_PROPERTYGET, &res, range->pOORange, L"CharEscapement", 0);
    if (FAILED(hres)) {
        TRACE("ERROR when put CharEscapement\n");
        return E_FAIL;
    }

    hres = VariantChangeTypeEx(&res, &res, 0, 0, VT_I2);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }
    if (V_I2(&res)>0) {
        V_VT(RHS) = VT_BOOL;
        V_BOOL(RHS) = VARIANT_TRUE;
    } else {
        V_VT(RHS) = VT_BOOL;
        V_BOOL(RHS) = VARIANT_FALSE;
    }

    return S_OK;*/
    TRACE("\n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Superscript(
        I_Font* iface,
        VARIANT RHS)
{
    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;
    VARIANT res;
    RangeImpl *range = (RangeImpl*)This->prange;
    VARIANT vsubscript;

    TRACE(" \n");

    VariantInit (&vsubscript);
    VariantInit (&res);

    hres = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        TRACE("ERROR VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }
    if (V_BOOL(&RHS)==VARIANT_TRUE) {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = 50;
    } else {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = 0;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharEscapement", 1, vsubscript);
    if (FAILED(hres)) {
        TRACE("ERROR when put CharEscapement\n");
        return E_FAIL;
    }

    return S_OK;
    TRACE("\n");
    return E_NOTIMPL;
}

/*IDispatch methods*/

static HRESULT WINAPI MSO_TO_OO_I_Font_GetTypeInfoCount(
        I_Font* iface,
        UINT *pctinfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_GetTypeInfo(
        I_Font* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_GetIDsOfNames(
        I_Font* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    if (!lstrcmpiW(*rgszNames, str_bold)) {
        *rgDispId = dispid_font_bold;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_italic)) {
        *rgDispId = dispid_font_italic;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_underline)) {
        *rgDispId = dispid_font_underline;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_size)) {
        *rgDispId = dispid_font_size;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_strikethrought)) {
        *rgDispId = dispid_font_strikethrough;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_name)) {
        *rgDispId = dispid_font_name;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_colorindex)) {
        *rgDispId = dispid_font_colorindex;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_color)) {
        *rgDispId = dispid_font_color;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_application)) {
        *rgDispId = dispid_font_application;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_parent)) {
        *rgDispId = dispid_font_parent;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_creator)) {
        *rgDispId = dispid_font_creator;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_shadow)) {
        *rgDispId = dispid_font_shadow;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_subscript)) {
        *rgDispId = dispid_font_subscript;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_superscript)) {
        *rgDispId = dispid_font_superscript;
        return S_OK;
    }
    /*Выводим название метода или свойства,
    чтобы знать чего не хватает.*/
    WTRACE(L" %s NOT REALIZE \n",*rgszNames);
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_Invoke(
        I_Font* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hr;
    VARIANT_BOOL vbin;
    VARIANT vtmp;
    VARIANT_BOOL ret;
    long lret;
    IDispatch *dret;
    VARIANT vresult;
    long ltmp;

    VariantInit(&vresult);
    VariantInit(&vtmp);
    /* http://www.spravkaweb.ru/php/sovet/calccom/formatfont */
    TRACE(" \n");

    if (This == NULL) return E_POINTER;

    switch(dispIdMember) 
    {
    case dispid_font_bold:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            hr = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE("ERROR VariantChangeTypeEx   %08x\n",hr);
                return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_Font_put_Bold(iface, vbin);
        } else {
            hr = MSO_TO_OO_I_Font_get_Bold(iface, &ret);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = ret;
            }
            return S_OK;
        }
    case dispid_font_italic:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            hr = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE("ERROR VariantChangeTypeEx   %08x\n",hr);
                return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_Font_put_Italic(iface, vbin);
        } else {
            hr = MSO_TO_OO_I_Font_get_Italic(iface, &ret);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = ret;
            }
            return S_OK;
        }
    case dispid_font_underline:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            return MSO_TO_OO_I_Font_put_Underline(iface, vtmp);
        } else {
            if (pVarResult!=NULL){
                hr = MSO_TO_OO_I_Font_get_Underline(iface, pVarResult);
                if (FAILED(hr)) {
                    pExcepInfo->bstrDescription=SysAllocString(str_error);
                    return hr;
                }
            }
            return S_OK;
        }
    case dispid_font_size:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            hr = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hr)) {
                TRACE("(case 4) ERROR VariantChangeTypeEx   %08x\n",hr);
                return E_FAIL;
            }
            lret = V_I4(&vtmp);
            return MSO_TO_OO_I_Font_put_Size(iface, lret);
        } else {
            hr = MSO_TO_OO_I_Font_get_Size(iface, &lret);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lret;
            }
            return S_OK;
        }
    case dispid_font_strikethrough:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            hr = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 5) ERROR VariantChangeTypeEx   %08x\n",hr);
                return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_Font_put_Strikethrough(iface, vbin);
        } else {
            hr = MSO_TO_OO_I_Font_get_Strikethrough(iface, &ret);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = ret;
            }
            return S_OK;
        }
    case dispid_font_name:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            return MSO_TO_OO_I_Font_put_Name(iface, vtmp);
        } else {
            hr = MSO_TO_OO_I_Font_get_Name(iface, pVarResult);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            return S_OK;
        }
    case dispid_font_colorindex:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            hr = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hr)) {
                TRACE(" (case 7) ERROR VariantChangeTypeEx   %08x\n",hr);
                return E_FAIL;
            }
            ltmp = V_I4(&vtmp);
            hr = MSO_TO_OO_I_Font_put_ColorIndex(iface,ltmp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            return S_OK;
        } else {
            hr = MSO_TO_OO_I_Font_get_ColorIndex(iface,&ltmp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = ltmp;
            }
            return S_OK;
        }
    case dispid_font_color:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            hr = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hr)) {
                TRACE("(case 8) ERROR VariantChangeTypeEx   %08x\n",hr);
                return E_FAIL;
            }
            ltmp = V_I4(&vtmp);
            hr = MSO_TO_OO_I_Font_put_Color(iface,ltmp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            return S_OK;
        } else {
            hr = MSO_TO_OO_I_Font_get_Color(iface,&ltmp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = ltmp;
            }
            return S_OK;
        }
    case dispid_font_application:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_Font_get_Application(iface,&dret);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case dispid_font_parent:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_Font_get_Parent(iface,&dret);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult)=VT_DISPATCH;
                V_DISPATCH(pVarResult)=dret;
            } else {
                IDispatch_Release(dret);
            }
            return S_OK;
        }
    case dispid_font_creator:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hr = MSO_TO_OO_I_Font_get_Creator(iface, &vresult);
            if (pVarResult!=NULL){
                *pVarResult = vresult;
            }
            return hr;
        }
    case dispid_font_shadow:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            hr = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_BOOL);
            if (FAILED(hr)) {
                TRACE(" (case 5) ERROR VariantChangeTypeEx   %08x\n",hr);
                return E_FAIL;
            }
            vbin = V_BOOL(&vtmp);
            return MSO_TO_OO_I_Font_put_Shadow(iface, vbin);
        } else {
            hr = MSO_TO_OO_I_Font_get_Shadow(iface, &ret);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = ret;
            }
            return S_OK;
        }
    case dispid_font_subscript:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            return MSO_TO_OO_I_Font_put_Subscript(iface, vtmp);
        } else {
            hr = MSO_TO_OO_I_Font_get_Subscript(iface, &vtmp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = V_BOOL(&vtmp);
            }
            return S_OK;
        }
    case dispid_font_superscript:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);
            return MSO_TO_OO_I_Font_put_Superscript(iface, vtmp);
        } else {
            hr = MSO_TO_OO_I_Font_get_Superscript(iface, &vtmp);
            if (FAILED(hr)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hr;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_BOOL;
                V_BOOL(pVarResult) = V_BOOL(&vtmp);
            }
            return S_OK;
        }


    }
    return E_NOTIMPL;
}


const I_FontVtbl MSO_TO_OO_I_Font_Vtbl =
{
    MSO_TO_OO_I_Font_QueryInterface,
    MSO_TO_OO_I_Font_AddRef,
    MSO_TO_OO_I_Font_Release ,
    MSO_TO_OO_I_Font_GetTypeInfoCount,
    MSO_TO_OO_I_Font_GetTypeInfo,
    MSO_TO_OO_I_Font_GetIDsOfNames,
    MSO_TO_OO_I_Font_Invoke,
    MSO_TO_OO_I_Font_get_Application,
    MSO_TO_OO_I_Font_get_Creator,
    MSO_TO_OO_I_Font_get_Parent,
    MSO_TO_OO_I_Font_get_Background,
    MSO_TO_OO_I_Font_put_Background,
    MSO_TO_OO_I_Font_get_Bold,
    MSO_TO_OO_I_Font_put_Bold,
    MSO_TO_OO_I_Font_get_Color,
    MSO_TO_OO_I_Font_put_Color,
    MSO_TO_OO_I_Font_get_ColorIndex,
    MSO_TO_OO_I_Font_put_ColorIndex,
    MSO_TO_OO_I_Font_get_FontStyle,
    MSO_TO_OO_I_Font_put_FontStyle,
    MSO_TO_OO_I_Font_get_Italic,
    MSO_TO_OO_I_Font_put_Italic,
    MSO_TO_OO_I_Font_get_Name,
    MSO_TO_OO_I_Font_put_Name,
    MSO_TO_OO_I_Font_get_OutlineFont,
    MSO_TO_OO_I_Font_put_OutlineFont,
    MSO_TO_OO_I_Font_get_Shadow,
    MSO_TO_OO_I_Font_put_Shadow,
    MSO_TO_OO_I_Font_get_Size,
    MSO_TO_OO_I_Font_put_Size,
    MSO_TO_OO_I_Font_get_Strikethrough,
    MSO_TO_OO_I_Font_put_Strikethrough,
    MSO_TO_OO_I_Font_get_Subscript,
    MSO_TO_OO_I_Font_put_Subscript,
    MSO_TO_OO_I_Font_get_Superscript,
    MSO_TO_OO_I_Font_put_Superscript,
    MSO_TO_OO_I_Font_get_Underline,
    MSO_TO_OO_I_Font_put_Underline
};

HRESULT _I_FontConstructor(LPVOID *ppObj)
{
    _FontImpl *_font;

    TRACE("(%p)\n", ppObj);
    
    _font = HeapAlloc(GetProcessHeap(), 0, sizeof(*_font));
    if (!_font)
    {
        return E_OUTOFMEMORY;
    }

    _font->_ifontVtbl = &MSO_TO_OO_I_Font_Vtbl;
    _font->ref = 0;
    _font->prange = NULL;

    *ppObj = &_font->_ifontVtbl;

    return S_OK;
}


