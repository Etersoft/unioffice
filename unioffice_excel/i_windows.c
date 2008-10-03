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

/*IWindows interface*/
    /*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Windows_AddRef(
        I_Windows* iface)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_QueryInterface(
        I_Windows* iface,
        REFIID riid,
        void **ppvObject)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static ULONG WINAPI MSO_TO_OO_I_Windows_Release(
        I_Windows* iface)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

   /*** I_Windows methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Application(
        I_Windows* iface,
        IDispatch **value)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Creator(
        I_Windows* iface,
        XlCreator *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Parent(
        I_Windows* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_Arrange(
        I_Windows* iface,
        XlArrangeStyle ArrangeStyle,
        VARIANT ActiveWorkbook,
        VARIANT SyncHorizontal,
        VARIANT SyncVertical,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Count(
        I_Windows* iface,
        long *retval)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get_Item(
        I_Windows* iface,
        VARIANT Index,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get__NewEnum(
        I_Windows* iface,
        IUnknown **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_get__Default(
        I_Windows* iface,
        VARIANT Index,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

    /*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Windows_GetTypeInfoCount(
        I_Windows* iface,
        UINT *pctinfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_GetTypeInfo(
        I_Windows* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_GetIDsOfNames(
        I_Windows* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Windows_Invoke(
        I_Windows* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

const I_WindowsVtbl MSO_TO_OO_I_WindowsVtbl =
{
    MSO_TO_OO_I_Windows_QueryInterface,
    MSO_TO_OO_I_Windows_AddRef,
    MSO_TO_OO_I_Windows_Release,
    MSO_TO_OO_I_Windows_GetTypeInfoCount,
    MSO_TO_OO_I_Windows_GetTypeInfo,
    MSO_TO_OO_I_Windows_GetIDsOfNames,
    MSO_TO_OO_I_Windows_Invoke,
    MSO_TO_OO_I_Windows_get_Application,
    MSO_TO_OO_I_Windows_get_Creator,
    MSO_TO_OO_I_Windows_get_Parent,
    MSO_TO_OO_I_Windows_Arrange,
    MSO_TO_OO_I_Windows_get_Count,
    MSO_TO_OO_I_Windows_get_Item,
    MSO_TO_OO_I_Windows_get__NewEnum,
    MSO_TO_OO_I_Windows_get__Default
};

extern HRESULT _I_WindowsConstructor(LPVOID *ppObj)
{
    WindowsImpl *windows;

    TRACE("(%p)\n", ppObj);

    windows = HeapAlloc(GetProcessHeap(), 0, sizeof(*windows));
    if (!windows)
    {
        return E_OUTOFMEMORY;
    }

    windows->_windowsVtbl = &MSO_TO_OO_I_WindowsVtbl;
    windows->ref = 0;
    windows->pApplication = NULL;

    *ppObj = &windows->_windowsVtbl;

    return S_OK;
}

/*
IWindow interface
*/

    /*** IUnknown methods ***/
static ULONG WINAPI MSO_TO_OO_I_Window_AddRef(
        I_Window* iface)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_QueryInterface(
        I_Window* iface,
        REFIID riid,
        void **ppvObject)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static ULONG WINAPI MSO_TO_OO_I_Window_Release(
        I_Window* iface)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

    /*** I_Window methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Window_get_Application(
        I_Window* iface,
        IDispatch **value)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Creator(
        I_Window* iface,
        XlCreator *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Parent(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Activate(
        I_Window* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActivateNext(
        I_Window* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActivatePrevious(
        I_Window* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActiveCell(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActiveChart(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActivePane(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ActiveSheet(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Caption(
        I_Window* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Caption(
        I_Window* iface,
        VARIANT RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Close(
        I_Window* iface,
        VARIANT SaveChanges,
        VARIANT Filename,
        VARIANT RouteWorkbook,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayFormulas(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayFormulas(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayGridlines(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayGridlines(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayHeadings(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayHeadings(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayHorizontalScrollBar(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayHorizontalScrollBar(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayOutline(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayOutline(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get__DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put__DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayVerticalScrollBar(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayVerticalScrollBar(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayWorkbookTabs(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayWorkbookTabs(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayZeros(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayZeros(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_EnableResize(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_EnableResize(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_FreezePanes(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_FreezePanes(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_GridlineColor(
        I_Window* iface,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_GridlineColor(
        I_Window* iface,
        long RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_GridlineColorIndex(
        I_Window* iface,
        XlColorIndex *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_GridlineColorIndex(
        I_Window* iface,
        XlColorIndex RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Height(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Height(
        I_Window* iface,
        double RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Index(
        I_Window* iface,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_LargeScroll(
        I_Window* iface,
        VARIANT Down,
        VARIANT Up,
        VARIANT toRight,
        VARIANT toLeft,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Left(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Left(
        I_Window* iface,
        double RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_NewWindow(
        I_Window* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_OnWindow(
        I_Window* iface,
        BSTR *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_OnWindow(
        I_Window* iface,
        BSTR RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Panes(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PrintOut(
        I_Window* iface,
        VARIANT From,
        VARIANT To_,
        VARIANT Copies,
        VARIANT Preview,
        VARIANT ActivePrinter,
        VARIANT PrintToFile,
        VARIANT Collate,
        VARIANT PrToFileName,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PrintPreview(
        I_Window* iface,
        VARIANT EnableChanges,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_RangeSelection(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ScrollColumn(
        I_Window* iface,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_ScrollColumn(
        I_Window* iface,
        long RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_ScrollRow(
        I_Window* iface,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_ScrollRow(
        I_Window* iface,
        long RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_ScrollWorkbookTabs(
        I_Window* iface,
        VARIANT Sheets,
        VARIANT Position,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SelectedSheets(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Selection(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_SmallScroll(
        I_Window* iface,
        VARIANT Down,
        VARIANT Up,
        VARIANT toRight,
        VARIANT toLeft,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Split(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Split(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitColumn(
        I_Window* iface,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitColumn(
        I_Window* iface,
        long RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitHorizontal(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitHorizontal(
        I_Window* iface,
        double RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitRow(
        I_Window* iface,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitRow(
        I_Window* iface,
        long RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_SplitVertical(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_SplitVertical(
        I_Window* iface,
        double RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_TabRatio(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_TabRatio(
        I_Window* iface,
        double RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Top(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Top(
        I_Window* iface,
        double RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Type(
        I_Window* iface,
        XlWindowType *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_UsableHeight(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_UsableWidth(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Visible(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Visible(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_VisibleRange(
        I_Window* iface,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Width(
        I_Window* iface,
        double *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Width(
        I_Window* iface,
        double RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_WindowNumber(
        I_Window* iface,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_WindowState(
        I_Window* iface,
        XlWindowState *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_WindowState(
        I_Window* iface,
        XlWindowState RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_Zoom(
        I_Window* iface,
        VARIANT *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_Zoom(
        I_Window* iface,
        VARIANT RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_View(
        I_Window* iface,
        XlWindowView *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_View(
        I_Window* iface,
        XlWindowView RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_get_DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_put_DisplayRightToLeft(
        I_Window* iface,
        VARIANT_BOOL RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PointsToScreenPixelsX(
        I_Window* iface,
        long Points,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_PointsToScreenPixelsY(
        I_Window* iface,
        long Points,
        long *RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_RangeFromPoint(
        I_Window* iface,
        long x,
        long y,
        IDispatch **RHS)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_ScrollIntoView(
        I_Window* iface,
        long left,
        long top,
        long width,
        long height,
        VARIANT Start)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

    /*** IDispatch methods ***/
static HRESULT WINAPI MSO_TO_OO_I_Window_GetTypeInfoCount(
        I_Window* iface,
        UINT *pctinfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_GetTypeInfo(
        I_Window* iface,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_GetIDsOfNames(
        I_Window* iface,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

static HRESULT WINAPI MSO_TO_OO_I_Window_Invoke(
        I_Window* iface,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr)
{
    TRACE(" \n");
    return E_NOTIMPL;
}

const I_WindowVtbl MSO_TO_OO_I_WindowVtbl =
{
    MSO_TO_OO_I_Window_QueryInterface,
    MSO_TO_OO_I_Window_AddRef,
    MSO_TO_OO_I_Window_Release,
    MSO_TO_OO_I_Window_GetTypeInfoCount,
    MSO_TO_OO_I_Window_GetTypeInfo,
    MSO_TO_OO_I_Window_GetIDsOfNames,
    MSO_TO_OO_I_Window_Invoke,
    MSO_TO_OO_I_Window_get_Application,
    MSO_TO_OO_I_Window_get_Creator,
    MSO_TO_OO_I_Window_get_Parent,
    MSO_TO_OO_I_Window_get_Activate,
    MSO_TO_OO_I_Window_get_ActivateNext,
    MSO_TO_OO_I_Window_get_ActivatePrevious,
    MSO_TO_OO_I_Window_get_ActiveCell,
    MSO_TO_OO_I_Window_get_ActiveChart,
    MSO_TO_OO_I_Window_get_ActivePane,
    MSO_TO_OO_I_Window_get_ActiveSheet,
    MSO_TO_OO_I_Window_get_Caption,
    MSO_TO_OO_I_Window_put_Caption,
    MSO_TO_OO_I_Window_get_Close,
    MSO_TO_OO_I_Window_get_DisplayFormulas,
    MSO_TO_OO_I_Window_put_DisplayFormulas,
    MSO_TO_OO_I_Window_get_DisplayGridlines,
    MSO_TO_OO_I_Window_put_DisplayGridlines,
    MSO_TO_OO_I_Window_get_DisplayHeadings,
    MSO_TO_OO_I_Window_put_DisplayHeadings,
    MSO_TO_OO_I_Window_get_DisplayHorizontalScrollBar,
    MSO_TO_OO_I_Window_put_DisplayHorizontalScrollBar,
    MSO_TO_OO_I_Window_get_DisplayOutline,
    MSO_TO_OO_I_Window_put_DisplayOutline,
    MSO_TO_OO_I_Window_get__DisplayRightToLeft,
    MSO_TO_OO_I_Window_put__DisplayRightToLeft,
    MSO_TO_OO_I_Window_get_DisplayVerticalScrollBar,
    MSO_TO_OO_I_Window_put_DisplayVerticalScrollBar,
    MSO_TO_OO_I_Window_get_DisplayWorkbookTabs,
    MSO_TO_OO_I_Window_put_DisplayWorkbookTabs,
    MSO_TO_OO_I_Window_get_DisplayZeros,
    MSO_TO_OO_I_Window_put_DisplayZeros,
    MSO_TO_OO_I_Window_get_EnableResize,
    MSO_TO_OO_I_Window_put_EnableResize,
    MSO_TO_OO_I_Window_get_FreezePanes,
    MSO_TO_OO_I_Window_put_FreezePanes,
    MSO_TO_OO_I_Window_get_GridlineColor,
    MSO_TO_OO_I_Window_put_GridlineColor,
    MSO_TO_OO_I_Window_get_GridlineColorIndex,
    MSO_TO_OO_I_Window_put_GridlineColorIndex,
    MSO_TO_OO_I_Window_get_Height,
    MSO_TO_OO_I_Window_put_Height,
    MSO_TO_OO_I_Window_get_Index,
    MSO_TO_OO_I_Window_LargeScroll,
    MSO_TO_OO_I_Window_get_Left,
    MSO_TO_OO_I_Window_put_Left,
    MSO_TO_OO_I_Window_NewWindow,
    MSO_TO_OO_I_Window_get_OnWindow,
    MSO_TO_OO_I_Window_put_OnWindow,
    MSO_TO_OO_I_Window_get_Panes,
    MSO_TO_OO_I_Window_PrintOut,
    MSO_TO_OO_I_Window_PrintPreview,
    MSO_TO_OO_I_Window_get_RangeSelection,
    MSO_TO_OO_I_Window_get_ScrollColumn,
    MSO_TO_OO_I_Window_put_ScrollColumn,
    MSO_TO_OO_I_Window_get_ScrollRow,
    MSO_TO_OO_I_Window_put_ScrollRow,
    MSO_TO_OO_I_Window_ScrollWorkbookTabs,
    MSO_TO_OO_I_Window_get_SelectedSheets,
    MSO_TO_OO_I_Window_get_Selection,
    MSO_TO_OO_I_Window_SmallScroll,
    MSO_TO_OO_I_Window_get_Split,
    MSO_TO_OO_I_Window_put_Split,
    MSO_TO_OO_I_Window_get_SplitColumn,
    MSO_TO_OO_I_Window_put_SplitColumn,
    MSO_TO_OO_I_Window_get_SplitHorizontal,
    MSO_TO_OO_I_Window_put_SplitHorizontal,
    MSO_TO_OO_I_Window_get_SplitRow,
    MSO_TO_OO_I_Window_put_SplitRow,
    MSO_TO_OO_I_Window_get_SplitVertical,
    MSO_TO_OO_I_Window_put_SplitVertical,
    MSO_TO_OO_I_Window_get_TabRatio,
    MSO_TO_OO_I_Window_put_TabRatio,
    MSO_TO_OO_I_Window_get_Top,
    MSO_TO_OO_I_Window_put_Top,
    MSO_TO_OO_I_Window_get_Type,
    MSO_TO_OO_I_Window_get_UsableHeight,
    MSO_TO_OO_I_Window_get_UsableWidth,
    MSO_TO_OO_I_Window_get_Visible,
    MSO_TO_OO_I_Window_put_Visible,
    MSO_TO_OO_I_Window_get_VisibleRange,
    MSO_TO_OO_I_Window_get_Width,
    MSO_TO_OO_I_Window_put_Width,
    MSO_TO_OO_I_Window_get_WindowNumber,
    MSO_TO_OO_I_Window_get_WindowState,
    MSO_TO_OO_I_Window_put_WindowState,
    MSO_TO_OO_I_Window_get_Zoom,
    MSO_TO_OO_I_Window_put_Zoom,
    MSO_TO_OO_I_Window_get_View,
    MSO_TO_OO_I_Window_put_View,
    MSO_TO_OO_I_Window_get_DisplayRightToLeft,
    MSO_TO_OO_I_Window_put_DisplayRightToLeft,
    MSO_TO_OO_I_Window_PointsToScreenPixelsX,
    MSO_TO_OO_I_Window_PointsToScreenPixelsY,
    MSO_TO_OO_I_Window_RangeFromPoint,
    MSO_TO_OO_I_Window_ScrollIntoView
};

extern HRESULT _I_WindowConstructor(LPVOID *ppObj)
{
    WindowImpl *window;

    TRACE("(%p)\n", ppObj);

    window = HeapAlloc(GetProcessHeap(), 0, sizeof(*window));
    if (!window)
    {
        return E_OUTOFMEMORY;
    }

    window->_windowVtbl = &MSO_TO_OO_I_WindowVtbl;
    window->ref = 0;
    window->pWindows = NULL;

    *ppObj = &window->_windowVtbl;

    return S_OK;
}

