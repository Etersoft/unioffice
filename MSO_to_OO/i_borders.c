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
static WCHAR const str__default[] = {
    '_','D','e','f','a','u','l','t',0};
static WCHAR const str_item[] = {
    'I','t','e','m',0};
static WCHAR const str_value[] = {
    'V','a','l','u','e',0};

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Borders_AddRef(
        I_Borders* iface)
{
    BordersImpl *This = (BordersImpl*)iface;
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

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
    BordersImpl *This = (BordersImpl*)iface;

    TRACE("\n");

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Borders)) {
        *ppvObject = &This->_bordersVtbl;
        MSO_TO_OO_I_Borders_AddRef(iface);
        return S_OK;
    }

    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Borders_Release(
        I_Borders* iface)
{
    BordersImpl *This = (BordersImpl*)iface;
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

/*** I_Borders methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Application(
        I_Borders* iface,
        IDispatch **value)
{
    BordersImpl *This = (BordersImpl*)iface;

    TRACE(" \n");

    if (This==NULL) return E_POINTER;
    if (This->prange==NULL) return E_POINTER;

    return I_Range_get_Application((I_Range*)(This->prange),value);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Parent(
        I_Borders* iface,
        IDispatch **value)
{
    BordersImpl *This = (BordersImpl*)iface;

    TRACE(" \n");

    if (This==NULL) return E_POINTER;

    *value = This->prange;
    I_Range_AddRef(This->prange);

    if (value==NULL)
        return E_POINTER;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Color(
        I_Borders* iface,
        long *plcolor)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;

    TRACE(" \n");

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_get_Color((I_Border*)border_tmp, plcolor);
    IDispatch_Release(border_tmp);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Color(
        I_Borders* iface,
        long lcolor)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;

    TRACE(" lcolor = %i\n",lcolor);

    if (This==NULL) return E_POINTER;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_put_Color((I_Border*)border_tmp, lcolor);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeBottom, &border_tmp);
    I_Border_put_Color((I_Border*)border_tmp, lcolor);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeLeft, &border_tmp);
    I_Border_put_Color((I_Border*)border_tmp, lcolor);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeRight, &border_tmp);
    I_Border_put_Color((I_Border*)border_tmp, lcolor);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlInsideHorizontal, &border_tmp);
    I_Border_put_Color((I_Border*)border_tmp, lcolor);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlInsideVertical, &border_tmp);
    I_Border_put_Color((I_Border*)border_tmp, lcolor);
    IDispatch_Release(border_tmp);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_ColorIndex(
        I_Borders* iface,
        long *plcolorindex)
{
    BordersImpl *This = (BordersImpl*)iface;
    long tmpcolor;
    int i;
    HRESULT hres;
    TRACE(" \n");

    if (This==NULL) return E_POINTER;

    hres = MSO_TO_OO_I_Borders_get_Color(iface,&tmpcolor);
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

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_ColorIndex(
        I_Borders* iface,
        long lcolorindex)
{
    BordersImpl *This = (BordersImpl*)iface;
    long tmpcolor;
    TRACE("\n");

    if (This==NULL) return E_POINTER;

    if (lcolorindex==xlColorIndexNone) lcolorindex = 2;
    if (lcolorindex==xlColorIndexAutomatic) lcolorindex = 1;

    if ((lcolorindex<1)||(lcolorindex>56)) {
        TRACE(" ERROR Incorrect colorindex %i\n", lcolorindex);
        return S_OK;
    } else 
        return MSO_TO_OO_I_Borders_put_Color(iface,color[lcolorindex-1]);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Creator(
        I_Borders* iface,
        VARIANT *result)
{
    TRACE("\n");
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_LineStyle(
        I_Borders* iface,
        XlLineStyle *plinestyle)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_get_LineStyle((I_Border*)border_tmp, plinestyle);
    IDispatch_Release(border_tmp);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_LineStyle(
        I_Borders* iface,
        XlLineStyle linestyle)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_put_LineStyle((I_Border*)border_tmp, linestyle);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeBottom, &border_tmp);
    I_Border_put_LineStyle((I_Border*)border_tmp, linestyle);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeLeft, &border_tmp);
    I_Border_put_LineStyle((I_Border*)border_tmp, linestyle);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeRight, &border_tmp);
    I_Border_put_LineStyle((I_Border*)border_tmp, linestyle);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlInsideHorizontal, &border_tmp);
    I_Border_put_LineStyle((I_Border*)border_tmp, linestyle);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlInsideVertical, &border_tmp);
    I_Border_put_LineStyle((I_Border*)border_tmp, linestyle);
    IDispatch_Release(border_tmp);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Weight(
        I_Borders* iface,
        XlBorderWeight *pweight)
{
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_get_Weight((I_Border*)border_tmp, pweight);
    IDispatch_Release(border_tmp);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Weight(
        I_Borders* iface,
        XlBorderWeight weight)
{
    TRACE(" \n");
    BordersImpl *This = (BordersImpl*)iface;
    HRESULT hres;
    IDispatch *border_tmp;

    TRACE("\n");

    if (This==NULL) return E_POINTER;

    I_Borders_get_Item(iface, xlEdgeTop, &border_tmp);
    I_Border_put_Weight((I_Border*)border_tmp, weight);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeBottom, &border_tmp);
    I_Border_put_Weight((I_Border*)border_tmp, weight);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeLeft, &border_tmp);
    I_Border_put_Weight((I_Border*)border_tmp, weight);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlEdgeRight, &border_tmp);
    I_Border_put_Weight((I_Border*)border_tmp, weight);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlInsideHorizontal, &border_tmp);
    I_Border_put_Weight((I_Border*)border_tmp, weight);
    IDispatch_Release(border_tmp);
    I_Borders_get_Item(iface, xlInsideVertical, &border_tmp);
    I_Border_put_Weight((I_Border*)border_tmp, weight);
    IDispatch_Release(border_tmp);

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get__Default(
        I_Borders* iface,
        XlBordersIndex key,
        IDispatch **ppObject)
{
    BordersImpl *This = (BordersImpl*)iface;
    IUnknown *punk = NULL;
    LPUNKNOWN pUnkOuter = NULL;
    IDispatch *pborder;
    HRESULT hres;

    TRACE("key=%08x\n",key);

    if (This==NULL) {
        TRACE("ERROR Object is NULL \n");
        return E_FAIL;
    }
    switch (key) {
    case xlDiagonalUp:
        TRACE("ERROR xlDiagonalUp\n");
        return E_NOTIMPL;
        break;
    case xlDiagonalDown:
        TRACE("ERROR xlDiagonalDown\n");
        return E_NOTIMPL;
        break;
    }
    /*������� ������ Border*/
    *ppObject = NULL;

    hres = _I_BorderConstructor(pUnkOuter, (LPVOID*) &punk);

    if (FAILED(hres)) return E_NOINTERFACE;

    hres = I_Border_QueryInterface(punk, &IID_I_Border, (void**) &pborder);
    if (pborder == NULL) {
        return E_FAIL;
    }

    hres = MSO_TO_OO_I_Border_Initialize((I_Border*)pborder, iface, key);

    if (FAILED(hres)) {
        IDispatch_Release(pborder);
        return hres;
    }

    *ppObject = pborder;

    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Item(
        I_Borders* iface,
        XlBordersIndex key,
        IDispatch **ppObject)
{
    TRACE("\n");
    return MSO_TO_OO_I_Borders_get__Default(iface, key, ppObject);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_get_Value(
        I_Borders* iface,
        XlLineStyle *plinestyle)
{
    TRACE("\n");
    return MSO_TO_OO_I_Borders_get_LineStyle(iface, plinestyle);
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_put_Value(
        I_Borders* iface,
        XlLineStyle linestyle)
{
    TRACE("\n");
    return MSO_TO_OO_I_Borders_put_LineStyle(iface, linestyle);
}

/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Borders_GetTypeInfoCount(
        I_Borders* iface,
        UINT *pctinfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetTypeInfo(
        I_Borders* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Borders_GetIDsOfNames(
        I_Borders* iface,
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
    if (!lstrcmpiW(*rgszNames, str__default)) {
        *rgDispId = 8;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_item)) {
        *rgDispId = 9;
        return S_OK;
    }
    if (!lstrcmpiW(*rgszNames, str_value)) {
        *rgDispId = 10;
        return S_OK;
    }

    /*������� �������� ������ ��� ��������,
    ����� ����� ���� �� �������.*/
    WTRACE(L"%s NOT REALIZE\n",*rgszNames);
    return E_NOTIMPL;
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
    HRESULT hres;
    VARIANT vresult,vtmp;
    IDispatch *dret;
    long ltmp,lval;
    double dtmp;

    VariantInit(&vtmp);
    VariantInit(&vresult);

    switch(dispIdMember)
    {
    case 1:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Borders_get_Application(iface,&dret);
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
    case 2:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Borders_get_Parent(iface,&dret);
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
    case 3:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);

            if (FAILED(hres)) {
                TRACE(" (case 3) ERROR VariantChangeTypeEx   %08x   VT = %i\n",hres,V_VT(&(pDispParams->rgvarg[0])));
                return E_FAIL;
            }
            ltmp = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Borders_put_Color(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Borders_get_Color(iface,&ltmp);
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
    case 4:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) return E_FAIL;

            hres = VariantChangeTypeEx(&vtmp, &(pDispParams->rgvarg[0]), 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (case 1) ERROR VariantChangeTypeEx   %08x\n",hres);
                return E_FAIL;
            }
            ltmp = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Borders_put_ColorIndex(iface,ltmp);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            return S_OK;
        } else {
            hres = MSO_TO_OO_I_Borders_get_ColorIndex(iface,&ltmp);
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
    case 5:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            hres = MSO_TO_OO_I_Borders_get_Creator(iface, &vresult);
            if (pVarResult!=NULL){
                *pVarResult = vresult;
            }
            return hres;
        }
    case 6:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (6) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*�������� ��������� � ���� VARIANT ���� ��� �������� �� ������*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (6) ERROR when VariantChangeType \n");
                return hres;
            }
            lval = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Borders_put_LineStyle(iface, lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_Borders_get_LineStyle(iface,(XlLineStyle*) &lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lval;
            }
            return hres;
        }
    case 7://Weight
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE(" (7) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*�������� ��������� � ���� VARIANT ���� ��� �������� �� ������*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (7) ERROR when VariantChangeType \n");
                return hres;
            }
            lval = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Borders_put_Weight(iface, lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_Borders_get_Weight(iface,(XlBorderWeight*) &lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lval;
            }
            return hres;
        }
    case 8://Default
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            if (pDispParams->cArgs!=1) {
                TRACE(" (8) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*�������� ��������� � ���� VARIANT ���� ��� �������� �� ������*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE("(8) ERROR when VariantChangeType \n");
                return hres;
            }
            lval = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Borders_get__Default(iface, lval, &dret);
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
            return hres;
        }
    case 9:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            return E_NOTIMPL;
        } else {
            if (pDispParams->cArgs!=1) {
                TRACE("(9) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*�������� ��������� � ���� VARIANT ���� ��� �������� �� ������*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (9) ERROR when VariantChangeType \n");
                return hres;
            }
            lval = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Borders_get_Item(iface, lval, &dret);
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
            return hres;
        }
    case 10:
        if (wFlags==DISPATCH_PROPERTYPUT) {
            if (pDispParams->cArgs!=1) {
                TRACE("(10) ERROR Number of parameters \n");
                return E_FAIL;
            }
            /*�������� ��������� � ���� VARIANT ���� ��� �������� �� ������*/
            MSO_TO_OO_CorrectArg(pDispParams->rgvarg[0], &vtmp);

            hres = VariantChangeTypeEx(&vtmp, &vtmp, 0, 0, VT_I4);
            if (FAILED(hres)) {
                TRACE(" (10) ERROR when VariantChangeType \n");
                return hres;
            }
            lval = V_I4(&vtmp);
            hres = MSO_TO_OO_I_Borders_put_Value(iface, lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
            }
            return hres;
        } else {
            hres = MSO_TO_OO_I_Borders_get_Value(iface,(XlLineStyle*) &lval);
            if (FAILED(hres)) {
                pExcepInfo->bstrDescription=SysAllocString(str_error);
                return hres;
            }
            if (pVarResult!=NULL){
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = lval;
            }
            return hres;
        }
    }

    TRACE(" dispIdMember = %i NOT REALIZE\n",dispIdMember);
    return E_NOTIMPL;
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
    MSO_TO_OO_I_Borders_get_Parent,
    MSO_TO_OO_I_Borders_get_Color,
    MSO_TO_OO_I_Borders_put_Color,
    MSO_TO_OO_I_Borders_get_ColorIndex,
    MSO_TO_OO_I_Borders_put_ColorIndex,
    MSO_TO_OO_I_Borders_get_Creator,
    MSO_TO_OO_I_Borders_get_LineStyle,
    MSO_TO_OO_I_Borders_put_LineStyle,
    MSO_TO_OO_I_Borders_get_Weight,
    MSO_TO_OO_I_Borders_put_Weight,
    MSO_TO_OO_I_Borders_get__Default,
    MSO_TO_OO_I_Borders_get_Item,
    MSO_TO_OO_I_Borders_get_Value,
    MSO_TO_OO_I_Borders_put_Value
};

BordersImpl MSO_TO_OO_Borders =
{
    &MSO_TO_OO_I_Borders_Vtbl,
    0,
    NULL
};

extern HRESULT _I_BordersConstructor(IUnknown *pUnkOuter, LPVOID *ppObj)
{
    BordersImpl *borders;

    TRACE("(%p,%p)\n", pUnkOuter, ppObj);
    
    borders = HeapAlloc(GetProcessHeap(), 0, sizeof(*borders));
    if (!borders)
    {
        return E_OUTOFMEMORY;
    }

    borders->_bordersVtbl = &MSO_TO_OO_I_Borders_Vtbl;
    borders->ref = 0;
    borders->prange = NULL;

    *ppObj = &borders->_bordersVtbl;

    return S_OK;
}
