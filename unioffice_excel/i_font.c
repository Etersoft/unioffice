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

#define FONT_THIS(iface) DEFINE_THIS(FontImpl, font, iface)

/*IUnknown*/
static ULONG WINAPI MSO_TO_OO_I_Font_AddRef(
        I_Font* iface)
{
    FontImpl *This = FONT_THIS(iface);
    ULONG ref;
    TRACE("REF = %i \n", This->ref);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

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
    FontImpl *This = FONT_THIS(iface);

    *ppvObject = NULL;

    if ((!This) || (!ppvObject)) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Font)) {
        *ppvObject = FONT_FONT(This);
        I_Font_AddRef(FONT_FONT(This));
    }
    TRACE_OUT;
    return E_NOINTERFACE;
}


static ULONG WINAPI MSO_TO_OO_I_Font_Release(
        I_Font* iface)
{
    FontImpl *This = FONT_THIS(iface);
    ULONG ref;
    TRACE("REF = %i \n", This->ref);

    if (!This) {
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

/*I_Font methods*/

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Bold(
        I_Font* iface,
        VARIANT_BOOL *pvbBold)
{
    FontImpl *This = FONT_THIS(iface);

    /*In OO bold is specified as weight of the character*/
    VARIANT vBoldState;
    TRACE_IN;

    VariantInit (&vBoldState);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vBoldState, This->pOORange, L"CharWeight", 0);
    if (FAILED(hres))  {
        ERR("CharWeight \n");
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
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vBoldState;
    VariantInit (&vBoldState);
    V_VT(&vBoldState) = VT_R4;

    if (vbBold == VARIANT_TRUE)
        V_R4(&vBoldState) = 150;
    else
        V_R4(&vBoldState) = 100;

    VARIANT res;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharWeight", 1, vBoldState);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Italic(
        I_Font* iface,
        VARIANT_BOOL *pvbItalic)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vItalicState;
    VariantInit (&vItalicState);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vItalicState, This->pOORange, L"CharPosture", 0);
    if (FAILED(hres)) {
        ERR("CharPosture \n");
        return hres;
    }
        
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
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vItalicState;
    VariantInit (&vItalicState);
    V_VT(&vItalicState) = VT_I2; /* !!!!! VT_INT */

    if (vbItalic == VARIANT_TRUE)
        V_I2(&vItalicState) = 1;
    else
        V_I2(&vItalicState) = 0;

    VARIANT res;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharPosture", 1, vItalicState);

    TRACE_OUT;
    return S_OK;
}

/* TODO 1 - ��� ������������� 2-���� �������������*/
static HRESULT WINAPI MSO_TO_OO_I_Font_get_Underline(
        I_Font* iface,
        VARIANT *pvbUnderline)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, This->pOORange, L"CharUnderline", 0);

    if (FAILED(hres)) {
       ERR("CharUnderline \n");
       return hres;
    }

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
/* TODO 1 - ��� ������������� 2-���� �������������*/
static HRESULT WINAPI MSO_TO_OO_I_Font_put_Underline(
        I_Font* iface,
        VARIANT vbUnderline)
{
    FontImpl *This = FONT_THIS(iface);
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

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharUnderline", 1, vUnderlineState);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Size(
        I_Font* iface,
        long *plsize)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vsize;
    VariantInit (&vsize);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vsize, This->pOORange, L"CharHeight", 0);

    if (FAILED(hres)) {
        ERR("CharHeight \n");
        return hres;
    }

    hres = VariantChangeTypeEx(&vsize, &vsize, 0, 0, VT_I4);
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
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;
    TRACE(" %i \n",lsize);

    VARIANT vsize;
    VariantInit (&vsize);

    V_VT(&vsize) = VT_I4;
    V_I4(&vsize) = lsize;

    VARIANT res;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharHeight", 1, vsize);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Strikethrough(
        I_Font* iface,
        VARIANT_BOOL *pvbUnderline)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, This->pOORange, L"CharStrikeout", 0);

    if (FAILED(hres)) {
        ERR("CharStrikeout \n");
        return hres;
    }
    
    *pvbUnderline = V_BOOL(&vUnderlineState);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Strikethrough(
        I_Font* iface,
        VARIANT_BOOL vbUnderline)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    V_VT(&vUnderlineState) = VT_BOOL;
    V_BOOL(&vUnderlineState) = vbUnderline;

    VARIANT res;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharStrikeout", 1, vUnderlineState);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Name(
        I_Font* iface,
        VARIANT *vName)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, vName, This->pOORange, L"CharFontName", 0);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Name(
        I_Font* iface,
        VARIANT vName)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT res;

    MSO_TO_OO_CorrectArg(vName, &vName);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharFontName", 1, vName);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Color(
        I_Font* iface,
        long *plcolor)
{
    FontImpl *This = FONT_THIS(iface);
    HRESULT hres;
    VARIANT vret;
    VariantInit(&vret);
    TRACE_IN;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    hres = AutoWrap(DISPATCH_PROPERTYGET, &vret, This->pOORange, L"CharColor", 0);

    if (FAILED(hres)) {
        ERR("CharColor \n");
    }

    hres = VariantChangeTypeEx(&vret, &vret, 0, 0, VT_I4);
    if (FAILED(hres)) {
        ERR("VariantChangeTypeEx   %08x\n",hres);
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
    FontImpl *This = FONT_THIS(iface);
    HRESULT hres;
    VARIANT vret,param1;
    TRACE_IN;
    TRACE(" lcolor = %i\n",lcolor);

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }


    VariantInit(&param1);
    V_VT(&param1) = VT_I4;
    V_I4(&param1) = lcolor;

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &vret, This->pOORange, L"CharColor", 1, param1);

    if (FAILED(hres)) 
        ERR("CharColor");
        
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_ColorIndex(
        I_Font* iface,
        long *plcolorindex)
{
    FontImpl *This = FONT_THIS(iface);
    long tmpcolor;
    int i;
    HRESULT hres;
    TRACE_IN;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    hres = I_Font_get_Color(iface, &tmpcolor);
    if (FAILED(hres)) {
        ERR("FONT_THIS \n");              
        return hres;
    }
    
    for (i=0;i<56;i++)
        if (color[i]==tmpcolor) {
            *plcolorindex = i+1;
            return S_OK;
        }

    ERR("Color don`t have colorindex \n");
    *plcolorindex = 1;/*����� ����*/
    /*���������� ��� ��� ������, �� ������ ������*/
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_ColorIndex(
        I_Font* iface,
        long lcolorindex)
{
    FontImpl *This = FONT_THIS(iface);
    long tmpcolor;
    TRACE_IN;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;

    TRACE_OUT;
    if ((lcolorindex<1)||(lcolorindex>56)) {
        ERR("Incorrect colorindex \n");
        return S_OK;
    } else 
        return I_Font_put_Color(iface, color[lcolorindex-1]);
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Application(
        I_Font* iface,
        IDispatch **value)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;
    
    *value = NULL;
    
    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    TRACE_OUT;
    return I_Range_get_Application(This->pRange, value);
}

static HRESULT WINAPI MSO_TO_OO_I_Font_get_Parent(
        I_Font* iface,
        IDispatch **value)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    *value = NULL;

    if (!This) {
        ERR("Object is NULL \n");
        return E_POINTER;
    }

    *value = (IDispatch*)(This->pRange);
    I_Range_AddRef(This->pRange);

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
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYGET, &vUnderlineState, This->pOORange, L"CharShadowed", 0);

    if (FAILED(hres)) {
        ERR("CharShadowed \n");                  
        return hres;
    }
    
    *pvbshadow = V_BOOL(&vUnderlineState);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Font_put_Shadow(
        I_Font* iface,
        VARIANT_BOOL vbshadow)
{
    FontImpl *This = FONT_THIS(iface);
    TRACE_IN;

    VARIANT vUnderlineState;
    VariantInit (&vUnderlineState);

    V_VT(&vUnderlineState) = VT_BOOL;
    V_BOOL(&vUnderlineState) = vbshadow;

    VARIANT res;

    HRESULT hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharShadowed", 1, vUnderlineState);

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
        ERR("when get_Bold");
    }
    if (tmp==VARIANT_TRUE) {
        if (pusto) swprintf(str, L"%s", L"bold");
            else swprintf(str, L"%s %s", str, L"bold");
        pusto = 0;
    }

    tmp = VARIANT_FALSE;
    hres = I_Font_get_Italic(iface, &tmp);
    if (FAILED(hres)) {
        ERR("when get_Italic");
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
        ERR("parameter not BSTR");
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
    FontImpl *This = FONT_THIS(iface);
    HRESULT hres;
    VARIANT res;
    VARIANT vsubscript;

    TRACE(" \n");

    MSO_TO_OO_CorrectArg(RHS, &RHS);

    VariantInit (&vsubscript);
    VariantInit (&res);

    hres = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        ERR("VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }
    if (V_BOOL(&RHS)==VARIANT_TRUE) {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = -50;
    } else {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = 0;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharEscapement", 1, vsubscript);
    if (FAILED(hres)) {
        ERR("when put CharEscapement\n");
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
    FontImpl *This = FONT_THIS(iface);
    HRESULT hres;
    VARIANT res;
    VARIANT vsubscript;

    TRACE(" \n");

    MSO_TO_OO_CorrectArg(RHS, &RHS);

    VariantInit (&vsubscript);
    VariantInit (&res);

    hres = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    if (FAILED(hres)) {
        ERR("VariantChangeTypeEx   %08x\n",hres);
        return E_FAIL;
    }
    if (V_BOOL(&RHS)==VARIANT_TRUE) {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = 50;
    } else {
        V_VT(&vsubscript) = VT_I4;
        V_BOOL(&vsubscript) = 0;
    }

    hres = AutoWrap(DISPATCH_PROPERTYPUT, &res, This->pOORange, L"CharEscapement", 1, vsubscript);
    if (FAILED(hres)) {
        ERR("when put CharEscapement\n");
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

#undef FONT_THIS

HRESULT _I_FontConstructor(LPVOID *ppObj)
{
    FontImpl *font;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    font = HeapAlloc(GetProcessHeap(), 0, sizeof(*font));
    if (!font)
    {
        return E_OUTOFMEMORY;
    }

    font->pfontVtbl = &MSO_TO_OO_I_Font_Vtbl;
    font->ref = 0;
    font->pRange = NULL;
    font->pOORange = NULL;
    *ppObj = FONT_FONT(font);
    
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}


