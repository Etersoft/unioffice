/*
 * ISheets interface functions
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

ITypeInfo *ti_sheets = NULL;

HRESULT get_typeinfo_sheets(ITypeInfo **typeinfo)
{
    ITypeLib *typelib;
    HRESULT hres;
    WCHAR file_name[]= {'u','n','i','o','f','f','i','c','e','_','e','x','c','e','l','.','t','l','b',0};

    if (ti_sheets) {
        *typeinfo = ti_sheets;
        return S_OK;
    }

    hres = LoadTypeLib(file_name, &typelib);
    if(FAILED(hres)) {
        TRACE("ERROR: LoadTypeLib hres = %08x \n", hres);
        return hres;
    }

    hres = typelib->lpVtbl->GetTypeInfoOfGuid(typelib, &IID_I_Sheets, &ti_sheets);
    typelib->lpVtbl->Release(typelib);

    *typeinfo = ti_sheets;
    return hres;
}


/*ISheets interface*/

#define SHEETS_THIS(iface) DEFINE_THIS(SheetsImpl, sheets, iface);

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Sheets_AddRef(
        I_Sheets* iface)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedIncrement(&This->ref);
    if (ref == 1) {
        InterlockedIncrement(&dll_ref);
    }
    return ref;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_QueryInterface(
        I_Sheets* iface,
        REFIID riid,
        void **ppvObject)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    WCHAR str_clsid[39];

    if (This == NULL || ppvObject == NULL) return E_POINTER;

    if (IsEqualGUID(riid, &IID_IDispatch) ||
            IsEqualGUID(riid, &IID_IUnknown) ||
            IsEqualGUID(riid, &IID_I_Sheets)) {
        *ppvObject = SHEETS_SHEETS(This);
        I_Sheets_AddRef((I_Sheets*)(*ppvObject));
        return S_OK;
    }
    if (IsEqualGUID(riid, &IID_IEnumVARIANT)) {
        *ppvObject = SHEETS_ENUM(This);
        IUnknown_AddRef((IUnknown*)(*ppvObject));
        return S_OK;
    }
    StringFromGUID2(riid, str_clsid, 39);
    WTRACE(L"(%s) not supported \n", str_clsid);
    return E_NOINTERFACE;
}

static ULONG WINAPI MSO_TO_OO_I_Sheets_Release(
        I_Sheets* iface)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    ULONG ref;

    TRACE("REF = %i \n", This->ref);

    if (This == NULL) return E_POINTER;

    ref = InterlockedDecrement(&This->ref);
    if (ref == 0) {
/*        if (This->pwb != NULL) {
            IDispatch_Release(This->pwb);*/
            This->pwb = NULL;
/*        }*/
        if (This->pOOSheets != NULL) {
            IDispatch_Release(This->pOOSheets);
            This->pOOSheets = NULL;
        }
        InterlockedDecrement(&dll_ref);
        HeapFree(GetProcessHeap(), 0, This);
        DELETE_OBJECT;
    }
    return ref;
}

/*** I_Sheets methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Sheets_get__Default(
        I_Sheets* iface,
        VARIANT varIndex,
        IDispatch **ppSheet)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    TRACE_IN;

    MSO_TO_OO_CorrectArg(varIndex, &varIndex);

    if (This == NULL) return E_POINTER;

    if ((This->pwb == NULL) && (This->pOOSheets == NULL)){
        return E_FAIL;
    }

    VARIANT resultSheet;
    I_Worksheet *pSheet = NULL;
    HRESULT hres;
    IUnknown *punk = NULL;

    /*преобразовываем любой тип к I4*/
    hres = VariantChangeTypeEx(&varIndex, &varIndex, 0, 0, VT_I4);

    if (V_VT(&varIndex) == VT_I4) {
        V_I4(&varIndex)--;

        hres = AutoWrap (DISPATCH_METHOD, &resultSheet, This->pOOSheets, L"getByIndex", 1, varIndex);
        if (hres!=S_OK)
            return hres;

        hres = _I_WorksheetConstructor((LPVOID*) &punk);
        if (FAILED(hres)) return E_NOINTERFACE;

        hres = I_Worksheet_QueryInterface(punk, &IID_I_Worksheet, (void**) &(pSheet));
        if (FAILED(hres)) return E_NOINTERFACE;


        MSO_TO_OO_I_Worksheet_Initialize(pSheet,(I_Workbook*)(This->pwb), V_DISPATCH(&resultSheet));

        *ppSheet = (IDispatch*)pSheet;
        TRACE_OUT;
        return S_OK;
    } else 
        if (V_VT(&varIndex) == VT_BSTR) {
            hres = AutoWrap (DISPATCH_METHOD, &resultSheet, This->pOOSheets, L"getByName", 1, varIndex);
            if (hres!=S_OK)
                return hres;

            hres = _I_WorksheetConstructor((LPVOID*) &punk);
            if (FAILED(hres)) return E_NOINTERFACE;

            hres = I_Worksheet_QueryInterface(punk, &IID_I_Worksheet, (void**) &(pSheet));
            if (FAILED(hres)) return E_NOINTERFACE;

            MSO_TO_OO_I_Worksheet_Initialize(pSheet,(I_Workbook*)(This->pwb), V_DISPATCH(&resultSheet));

            *ppSheet = (IDispatch*)pSheet;
            TRACE_OUT;
            return S_OK;
        } else {
            *ppSheet = NULL;
            return E_FAIL;
        }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_Count(
        I_Sheets* iface,
        int *count)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    VARIANT res;
    VariantInit(&res);
    TRACE_IN;

    if (This == NULL) return E_POINTER;

    if (This->pOOSheets == NULL) return E_POINTER;

    HRESULT hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheets, L"getCount", 0);
    if (hres!=S_OK) {
        TRACE("ERROR when getCount \n");
        return hres;
    }
    *count = V_I4(&res);
    TRACE("return = %i \n",*count);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_Application(
        I_Sheets* iface,
        IDispatch **value)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    TRACE_IN;

    if (This == NULL) {
        TRACE("ERROR: This Object is NULL\n");
        return E_POINTER;
    }
    if (This->pwb == NULL){
        TRACE("ERROR: pwb Object is NULL\n");
        return E_POINTER;
    }
    TRACE_OUT;
    return I_Workbook_get_Application((I_Workbook*)(This->pwb), value);
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_Parent(
        I_Sheets* iface,
        IDispatch **value)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    TRACE_IN;

    if (This == NULL) {
        TRACE("ERROR: This Object is NULL\n");
        return E_POINTER;
    }
    if (This->pwb == NULL){
        TRACE("ERROR: pwb Object is NULL\n");
        return E_POINTER;
    }

    *value = This->pwb;
    IDispatch_AddRef(This->pwb);

    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_Item(
        I_Sheets* iface,
        VARIANT varIndex,
        IDispatch **ppSheet)
{
    TRACE(" ----> get__Default");
    return MSO_TO_OO_I_Sheets_get__Default(iface,varIndex,ppSheet);
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_Creator(
        I_Sheets* iface,
        VARIANT *result)
{
    TRACE_IN;
    V_VT(result) = VT_I4;
    V_I4(result) = 1480803660;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_Add(
        I_Sheets* iface,
        VARIANT Before,
        VARIANT After,
        VARIANT Count,
        VARIANT Type,
        IDispatch **value)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    int ftype_add = 0,i, j;
    int count;
    HRESULT hres;
    VARIANT par1,par2,res;
    BSTR tmp;
    IDispatch *wsh = NULL;
    TRACE_IN;

    VariantInit(&par1);
    VariantInit(&par2);
    VariantInit(&res);

    MSO_TO_OO_CorrectArg(Before, &Before);
    MSO_TO_OO_CorrectArg(After, &After);
    MSO_TO_OO_CorrectArg(Count, &Count);
    MSO_TO_OO_CorrectArg(Type, &Type);

    if (This == NULL) {
        TRACE("ERROR: This Object is NULL\n");
        return E_POINTER;
    }
    /*Приводим все значения к необходимому виду.*/
    if (Is_Variant_Null(Before)) {
        VariantClear(&Before);
    } else {
        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&Before, &Before, 0, 0, VT_I4);
        /*или останется текст*/
        ftype_add = 1;
    }
    if (Is_Variant_Null(After)) {
        VariantClear(&After);
    } else {
        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&After, &After, 0, 0, VT_I4);
        /*или останется текст*/
        ftype_add = 2;
    }
    if (Is_Variant_Null(Count)) {
        VariantClear(&Count);
        V_VT(&Count) = VT_I4;
        V_I4(&Count) = 1;
    } else {
        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&Count, &Count, 0, 0, VT_I4);
        if (FAILED(hres)) {
            TRACE("ERROR when VariantChangeTypeEx -Count-\n");
        }
    }
    if (Is_Variant_Null(Type)) {
        VariantClear(&Type);
        V_VT(&Type) = VT_I4;
        V_I4(&Type) = xlWorksheet;
    } else {
        /*преобразовываем любой тип к I4*/
        hres = VariantChangeTypeEx(&Type, &Type, 0, 0, VT_I4);
        if (FAILED(hres)) {
            TRACE("ERROR when VariantChangeTypeEx -Type-\n");
        }
        /*Поддерживается только xlWorksheet*/
        switch (V_I4(&Type)) {
        case xlWorksheet:break;
        default :
            TRACE("ERROR: This Type not implemented type = %i \n", V_I4(&Type));
            return E_FAIL;
        }
    }
    /*Должна быть разная реализация в зависимости от параметров*/
    VariantClear(&par2);
    V_VT(&par2) = VT_I4;
    /*Получаем общее кол-во таблиц*/
    MSO_TO_OO_I_Sheets_get_Count(iface, &count);
    V_I4(&par2) = 0;

    switch (ftype_add) {
    case 1: //перед указанным элементом
        WTRACE(L" before element %s\n",V_BSTR(&Before));
        if (V_VT(&Before) == VT_I4) {
            /*Если нам повезло и прислан индекс, то*/
            V_I4(&par2) = V_I4(&Before) - 1;
        } else {
            /*Если нет, то ищем по имени */
            i = MSO_TO_OO_FindIndexWorksheetByName(iface, V_BSTR(&Before));
            if (i>=0) V_I4(&par2) = i; else V_I4(&par2) = 0;
        }
        break;
    case 2: //после указанного элемента
        WTRACE(L"after element %s\n",V_BSTR(&After));
        if (V_VT(&After) == VT_I4) {
            /*Если нам повезло и прислан индекс, то*/
            V_I4(&par2) = V_I4(&After);
        } else {
            /*Если нет, то ищем по имени */
            i = MSO_TO_OO_FindIndexWorksheetByName(iface, V_BSTR(&After));
            if (i>=0) V_I4(&par2) = i+1; else V_I4(&par2) = 0;
        }
        break;
    case 0: //в начало списка
    default:
        TRACE(" to the begining of the list \n");
        V_I4(&par2) = 0;
    }

    for (i=V_I4(&Count);i>0;i--) {
        j=0;
        do {
            VariantClear(&par1);
            V_VT(&par1) = VT_BSTR;
            V_BSTR(&par1) = SysAllocString(L"Sheet");
            hres = VarBstrFromI4(count+i+j, 0, 0, &tmp);
            if (FAILED(hres)) {
                TRACE("ERROR when VarBSTRFromI4\n");
                tmp = SysAllocString(L"4");
            }
            VarBstrCat(V_BSTR(&par1), tmp, &(V_BSTR(&par1)));
            j++;
            hres = I_Sheets_get__Default(iface, par1, &wsh);
            if (wsh!=NULL) {
                IDispatch_Release(wsh);
                wsh = NULL;
                }
        } while (!FAILED(hres));

        hres = AutoWrap(DISPATCH_METHOD, &res, This->pOOSheets, L"insertNewByName", 2,par2,par1);
        if (FAILED(hres)) {
            TRACE("ERROR when insertNewByName\n");
            SysFreeString(tmp);
            return hres;
        }
        SysFreeString(tmp);
    }
    V_I4(&par2)++;
    hres = MSO_TO_OO_I_Sheets_get__Default(iface, par2,value);
    I_Worksheet_Activate((I_Worksheet*)(*value), 0);

    VariantClear(&par1);
    VariantClear(&par2);

    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_Copy(
        I_Sheets* iface,
        VARIANT Before,
        VARIANT After,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_Delete(
        I_Sheets* iface,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_FillAcrossSheets(
        I_Sheets* iface,
        IDispatch *IRange,
        XlFillWith Type,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_Move(
        I_Sheets* iface,
        VARIANT Before,
        VARIANT After,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get__NewEnum(
        I_Sheets* iface,
        IUnknown **RHS)
{
    SheetsImpl *This = SHEETS_THIS(iface);
    TRACE_IN;
    *RHS = (IUnknown*)SHEETS_ENUM(This);
    IUnknown_AddRef(*RHS);
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets__PrintOut(
        I_Sheets* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_PrintPreview(
        I_Sheets* iface,
        VARIANT EnableChanges,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_Select(
        I_Sheets* iface,
        VARIANT Replace,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_HPageBreaks(
        I_Sheets* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_VPageBreaks(
        I_Sheets* iface,
        IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_get_Visible(
        I_Sheets* iface,
        LCID lcid,
        VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_put_Visible(
        I_Sheets* iface,
        LCID lcid,
        VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_PrintOut(
        I_Sheets* iface,
        VARIANT From,
        VARIANT To,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        VARIANT PrToFileName,
        LCID lcid)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}


/*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Sheets_GetTypeInfoCount(
        I_Sheets* iface,
        UINT *pctinfo)
{
    TRACE_IN;
    *pctinfo = 1;
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_GetTypeInfo(
        I_Sheets* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    HRESULT hres = get_typeinfo_sheets(ppTInfo);
    TRACE("\n");
    if (FAILED(hres))
        TRACE("Error when GetTypeInfo");

    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_GetIDsOfNames(
        I_Sheets* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    ITypeInfo *typeinfo;
    HRESULT hres;
    TRACE_IN;
    hres = get_typeinfo_sheets(&typeinfo);
    if(FAILED(hres))
        return hres;

    hres = typeinfo->lpVtbl->GetIDsOfNames(typeinfo,rgszNames, cNames, rgDispId);
    if (FAILED(hres)) {
        WTRACE(L"ERROR name = %s \n", *rgszNames);
    }
    TRACE_OUT;
    return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_Invoke(
        I_Sheets* iface,
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
    ITypeInfo *typeinfo;

    if (iface == NULL) return E_POINTER;

    switch (dispIdMember)
    {
    case dispid_sheets_visible:
    case dispid_sheets__printout:
        break;
    default:
        /*For default*/
        hres = get_typeinfo_sheets(&typeinfo);
        if(FAILED(hres))
            return hres;

        hres = typeinfo->lpVtbl->Invoke(typeinfo, iface, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
        if (FAILED(hres)) {
            TRACE("ERROR wFlags = %i, cArgs = %i, dispIdMember = %i \n", wFlags,pDispParams->cArgs, dispIdMember);
        }
        return hres;
    }
    TRACE("NOT IMPL \n");
    return E_NOTIMPL;
}

const I_SheetsVtbl MSO_TO_OO_I_SheetsVtbl =
{
    MSO_TO_OO_I_Sheets_QueryInterface,
    MSO_TO_OO_I_Sheets_AddRef,
    MSO_TO_OO_I_Sheets_Release,
    MSO_TO_OO_I_Sheets_GetTypeInfoCount,
    MSO_TO_OO_I_Sheets_GetTypeInfo,
    MSO_TO_OO_I_Sheets_GetIDsOfNames,
    MSO_TO_OO_I_Sheets_Invoke,
    MSO_TO_OO_I_Sheets_get_Application,
    MSO_TO_OO_I_Sheets_get_Creator,
    MSO_TO_OO_I_Sheets_get_Parent,
    MSO_TO_OO_I_Sheets_Add,
    MSO_TO_OO_I_Sheets_Copy,
    MSO_TO_OO_I_Sheets_get_Count,
    MSO_TO_OO_I_Sheets_Delete,
    MSO_TO_OO_I_Sheets_FillAcrossSheets,
    MSO_TO_OO_I_Sheets_get_Item,
    MSO_TO_OO_I_Sheets_Move,
    MSO_TO_OO_I_Sheets_get__NewEnum,
    MSO_TO_OO_I_Sheets__PrintOut,
    MSO_TO_OO_I_Sheets_PrintPreview,
    MSO_TO_OO_I_Sheets_Select,
    MSO_TO_OO_I_Sheets_get_HPageBreaks,
    MSO_TO_OO_I_Sheets_get_VPageBreaks,
    MSO_TO_OO_I_Sheets_get_Visible,
    MSO_TO_OO_I_Sheets_put_Visible,
    MSO_TO_OO_I_Sheets_get__Default,
    MSO_TO_OO_I_Sheets_PrintOut
};

#undef SHEETS_THIS

/*IEnumVARIANT interface*/

#define ENUMVAR_THIS(iface) DEFINE_THIS(SheetsImpl, enumerator, iface);

/*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Sheets_EnumVAR_AddRef(
        IEnumVARIANT* iface)
{
    SheetsImpl *This = ENUMVAR_THIS(iface);
    return I_Sheets_AddRef(SHEETS_SHEETS(This));
}


static HRESULT WINAPI MSO_TO_OO_I_Sheets_EnumVAR_QueryInterface(
        IEnumVARIANT* iface,
        REFIID riid,
        void **ppvObject)
{
    SheetsImpl *This = ENUMVAR_THIS(iface);
    return I_Sheets_QueryInterface(SHEETS_SHEETS(This), riid, ppvObject);
}

static ULONG WINAPI MSO_TO_OO_I_Sheets_EnumVAR_Release(
        IEnumVARIANT* iface)
{
    SheetsImpl *This = ENUMVAR_THIS(iface);
    return I_Sheets_Release(SHEETS_SHEETS(This));
}

/*** IEnumVARIANT methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Sheets_EnumVAR_Next(
        IEnumVARIANT* iface,
        ULONG celt,
        VARIANT *rgVar,
        ULONG *pCeltFetched)
{
    SheetsImpl *This = ENUMVAR_THIS(iface);
    HRESULT hres;
    ULONG l;
    long l1;
    int count;
    ULONG l2;
    IDispatch *dret;
    VARIANT varindex;

    if (This->enum_position<0)
        return S_FALSE;

    if (pCeltFetched != NULL)
       *pCeltFetched = 0;

    if (rgVar == NULL)
       return E_INVALIDARG;

    VariantInit(&varindex);
    /*Init Array*/
    for (l=0; l<celt; l++)
       VariantInit(&rgVar[l]);

    I_Sheets_get_Count(SHEETS_SHEETS(This), &count);
    V_VT(&varindex) = VT_I4;

    for (l1=This->enum_position, l2=0; l1<count && l2<celt; l1++, l2++) {
      V_I4(&varindex) = l1+1;
      hres = I_Sheets_get_Item(SHEETS_SHEETS(This), varindex, &dret);
      V_VT(&rgVar[l2]) = VT_DISPATCH;
      V_DISPATCH(&rgVar[l2]) = dret;
      if (FAILED(hres))
         goto error;
    }

    if (pCeltFetched != NULL)
       *pCeltFetched = l2;

   This->enum_position = l1;

   return  (l2 < celt) ? S_FALSE : S_OK;

error:
   for (l=0; l<celt; l++)
      VariantClear(&rgVar[l]);
   VariantClear(&varindex);
   return hres;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_EnumVAR_Skip(
        IEnumVARIANT* iface,
        ULONG celt)
{
    SheetsImpl *This = ENUMVAR_THIS(iface);
    int count;
    TRACE_IN;

    I_Sheets_get_Count(SHEETS_SHEETS(This), &count);
    This->enum_position += celt;

    if (This->enum_position>=(count)) {
        This->enum_position = count - 1;
        TRACE_OUT;
        return S_FALSE;
    }
    TRACE_OUT;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_EnumVAR_Reset(
        IEnumVARIANT* iface)
{
    SheetsImpl *This = ENUMVAR_THIS(iface);
    This->enum_position = 0;
    return S_OK;
}

static HRESULT WINAPI MSO_TO_OO_I_Sheets_EnumVAR_Clone(
        IEnumVARIANT* iface,
        IEnumVARIANT **ppEnum)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;
}

#undef ENUMVAR_THIS

const IEnumVARIANTVtbl MSO_TO_OO_I_Sheets_enumvarVtbl =
{
    MSO_TO_OO_I_Sheets_EnumVAR_QueryInterface,
    MSO_TO_OO_I_Sheets_EnumVAR_AddRef,
    MSO_TO_OO_I_Sheets_EnumVAR_Release,
    MSO_TO_OO_I_Sheets_EnumVAR_Next,
    MSO_TO_OO_I_Sheets_EnumVAR_Skip,
    MSO_TO_OO_I_Sheets_EnumVAR_Reset,
    MSO_TO_OO_I_Sheets_EnumVAR_Clone
};

extern HRESULT _I_SheetsConstructor(LPVOID *ppObj)
{
    SheetsImpl *sheets;
    TRACE_IN;
    TRACE("(%p)\n", ppObj);

    sheets = HeapAlloc(GetProcessHeap(), 0, sizeof(*sheets));
    if (!sheets)
    {
        return E_OUTOFMEMORY;
    }

    sheets->psheetsVtbl = &MSO_TO_OO_I_SheetsVtbl;
    sheets->penumeratorVtbl = &MSO_TO_OO_I_Sheets_enumvarVtbl;
    sheets->ref = 0;
    sheets->pwb = NULL;
    sheets->pOOSheets =NULL;
    sheets->enum_position = 0;

    *ppObj = SHEETS_SHEETS(sheets);
    
    CREATE_OBJECT;
    
    TRACE_OUT;
    return S_OK;
}

/*
Properties 
| Count | HPageBreaks | Visible | VPageBreaks 

Methods 
| Copy | Delete | FillAcrossSheets | Move | PrintOut  | PrintPreview | Select 
*/
