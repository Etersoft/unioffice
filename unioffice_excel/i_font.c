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

ITypeInfo *ti_font = NULL;

HRESULT get_typeinfo_font(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if(ti_font) {
        *typeinfo = ti_font;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Font, &ti_font);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_font;
    return hres;
}


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

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Font)) {
        *ppvObject = &This->_ifontVtbl;
        MSO_TO_OO_I_Font_AddRef(iface);
    }
    TRACE_OUT;
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
    TRACE_IN;

    VariantInit (&vBoldState);

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Bold(
        I_Font* iface,
        VARIANT_BOOL vbBold)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

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

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Italic(
        I_Font* iface,
        VARIANT_BOOL *pvbItalic)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;

}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Italic(
        I_Font* iface,
        VARIANT_BOOL vbItalic)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}
/* TODO 1 - нет подчеркивания 2-есть подчеркивание*/
static HRESULT WINAPI MSO_TO_OO_I_Font_get_Underline(
        I_Font* iface,
        VARIANT *pvbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}
/* TODO 1 - нет подчеркивания 2-есть подчеркивание*/
static HRESULT WINAPI MSO_TO_OO_I_Font_put_Underline(
        I_Font* iface,
        VARIANT vbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    MSO_TO_OO_CorrectArg(vbUnderline, &vbUnderline);

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

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Size(
        I_Font* iface,
        long *plsize)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Size(
        I_Font* iface,
        long lsize)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;
    TRACE(" %i \n",lsize);

    VARIANT vsize;
    VariantInit (&vsize);

    V_VT(&vsize) = VT_I4;
    V_I4(&vsize) = lsize;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharHeight", 1, vsize);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Strikethrough(
        I_Font* iface,
        VARIANT_BOOL *pvbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, range->pOORange, L"CharStrikeout", 0);

    if (hres != S_OK)
        return hres;
    *pvbUnderline = V_BOOL(&vUnderlineState);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Strikethrough(
        I_Font* iface,
        VARIANT_BOOL vbUnderline)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    V_VT(&vUnderlineState) = VT_BOOL;
    V_BOOL(&vUnderlineState) = vbUnderline;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharStrikeout", 1, vUnderlineState);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Name(
        I_Font* iface,
        VARIANT *vName)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, vName, range->pOORange, L"CharFontName", 0);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Name(
        I_Font* iface,
        VARIANT vName)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

    VARIANT res;

    MSO_TO_OO_CorrectArg(vName, &vName);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharFontName", 1, vName);

    TRACE_OUT;
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
    TRACE_IN;

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

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Color(
        I_Font* iface,
        long lcolor)
{
    _FontImpl *This = (_FontImpl*)iface;
    HRESULT hres;
    VARIANT vret,param1;
    TRACE_IN;
    TRACE(" lcolor = %i\n",lcolor);

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    RangeImpl *cur_range = (RangeImpl*)((I_Range*)(This->prange));

    VariantInit(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = lcolor;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, cur_range->pOORange, L"CharColor", 1, param1);

    if (FAILED(hres)) TRACE("ERROR when CharColor");
    TRACE_OUT;
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
    TRACE_IN;

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
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_ColorIndex(
        I_Font* iface,
        long lcolorindex)
{
    _FontImpl *This = (_FontImpl*)iface;
    long tmpcolor;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;

    TRACE_OUT;
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
    TRACE_IN;

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;
    TRACE_OUT;
    return I_Range_get_Application((I_Range*)(This->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Parent(
        I_Font* iface,
        IDispatch **value)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

    if (This==NULL) return E_POINTER;

    if (value==NULL)
        return E_POINTER;

    *value = This->prange;
    I_Range_AddRef(This->prange);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Creator(
        I_Font* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Shadow(
        I_Font* iface,
        VARIANT_BOOL *pvbshadow)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, range->pOORange, L"CharShadowed", 0);

    if (hres != S_OK)
        return hres;
    *pvbshadow = V_BOOL(&vUnderlineState);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Shadow(
        I_Font* iface,
        VARIANT_BOOL vbshadow)
{
    _FontImpl *This = (_FontImpl*)iface;
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    V_VT(&vUnderlineState) = VT_BOOL;
    V_BOOL(&vUnderlineState) = vbshadow;

    VARIANT res;

    RangeImpl *range = (RangeImpl*)This->prange;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, range->pOORange, L"CharShadowed", 1, vUnderlineState);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Background(
        I_Font* iface,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Background(
        I_Font* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_FontStyle(
        I_Font* iface,
        VARIANT *RHS)
{
    TRACE_IN;
    WCHAR str[200];
    VARIANT_BOOL tmp;
    int pusto = 1;
    HRESULT hres;

    tmp = VARIANT_FALSE;
    hres = I_Font_get_Bold(iface, &tmp);
    if (FAILED(hres)) {
        TRACE("ERROR when get_Bold");
    }
    if (tmp==VARIANT_TRUE) {
        if (pusto) swprintf(str, L"%s", L"bold");
            else swprintf(str, L"%s %s", str, L"bold");
        pusto = 0;
    }

    tmp = VARIANT_FALSE;
    hres = I_Font_get_Italic(iface, &tmp);
    if (FAILED(hres)) {
        TRACE("ERROR when get_Italic");
    }
    if (tmp==VARIANT_TRUE) {
        if (pusto) swprintf(str, L"%s", L"italic");
            else swprintf(str, L"%s %s", str, L"italic");
        pusto = 0;
    }

    if (pusto) swprintf(str, L"%s", L"normal");

    V_VT(RHS) = VT_BSTR;
    V_BSTR(RHS) = SysAllocString(str);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_FontStyle(
        I_Font* iface,
        VARIANT RHS)
{
    TRACE_IN;
    static WCHAR str_bold_en[] = {
        'b','o','l','d',0};
    static WCHAR str_italic_en[] = {
        'i','t','a','l','i','c',0};
    static WCHAR str_bold_ru[] = {
        0x0436,0x0438,0x0440,0x043d,0x044b, 0x0439,0};
    static WCHAR str_italic_ru[] = {
        0x043a,0x0443,0x0440,0x0441,0x0438, 0x0432,0};
    static WCHAR str_bold2_ru[] = {
        0x043f, 0x043e, 0x043b, 0x0443,0x0436,0x0438,0x0440,0x043d,0x044b, 0x0439,0};

    int i = 0;
    WCHAR str[100];

    MSO_TO_OO_CorrectArg(RHS, &RHS);

    if (V_VT(&RHS)!=VT_BSTR) {
        TRACE("ERROR parameter not BSTR");
        return E_FAIL;
    }

    str[0] = 0;
    while (*(V_BSTR(&RHS)+i)) {
        if (*(V_BSTR(&RHS)+i)==L' ') {
            if ((!lstrcmpiW(str, str_bold_en)) ||
                (!lstrcmpiW(str, str_bold_ru)) ||
                (!lstrcmpiW(str, str_bold2_ru))) {
                 I_Font_put_Bold(iface, VARIANT_TRUE);
            }
            if ((!lstrcmpiW(str, str_italic_en)) ||
                (!lstrcmpiW(str, str_italic_ru))) {
                 I_Font_put_Italic(iface, VARIANT_TRUE);
            }
            str[0] = 0;
        } else {
            swprintf(str, L"%s%c",str, *(V_BSTR(&RHS)+i));
        }
        i++;
    }
    if ((!lstrcmpiW(str, str_bold_en)) ||
        (!lstrcmpiW(str, str_bold_ru)) ||
        (!lstrcmpiW(str, str_bold2_ru)))  {
         I_Font_put_Bold(iface, VARIANT_TRUE);
    }
    if ((!lstrcmpiW(str, str_italic_en)) ||
        (!lstrcmpiW(str, str_italic_ru))) {
         I_Font_put_Italic(iface, VARIANT_TRUE);
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_OutlineFont(
        I_Font* iface,
        VARIANT *RHS)
{
    TRACE_IN;
    V_VT(RHS) = VT_BOOL;
    V_BOOL(RHS) = VARIANT_FALSE;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_OutlineFont(
        I_Font* iface,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return S_OK;
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
    TRACE_NOTIMPL;
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

    MSO_TO_OO_CorrectArg(RHS, &RHS);

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
/*    TRACE_NOTIMPL;
    return E_NOTIMPL;*/
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
    TRACE_NOTIMPL;
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

    MSO_TO_OO_CorrectArg(RHS, &RHS);

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
/*    TRACE_NOTIMPL;
    return E_NOTIMPL;*/
}

/*IDispatch methods*/

static HRESULT WINAPI MSO_TO_OO_I_Font_GetTypeInfoCount(
        I_Font* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_GetTypeInfo(
        I_Font* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_font(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_GetIDsOfNames(
        I_Font* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_font(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
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
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;

    hres = get_typeinfo_font(&typeinfo);
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
    TRACE_IN;
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
    TRACE_OUT;
    return S_OK;
}


