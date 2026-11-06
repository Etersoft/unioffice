/*
 * Definition of special functions
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

HRESULT MSO_TO_OO_I_Workbooks_Initialize(
        I_Workbooks* iface,
        _Application *app);

HRESULT MSO_TO_OO_I_Font_Initialize(
        I_Font* iface,
        I_Range *range);

HRESULT MSO_TO_OO_I_Interior_Initialize(
        I_Interior* iface,
        I_Range *range);

HRESULT MSO_TO_OO_I_Borders_Initialize(
        I_Borders* iface,
        I_Range *range);

HRESULT MSO_TO_OO_I_Border_Initialize(
        I_Border* iface,
        I_Borders *borders,
        XlBordersIndex key);

HRESULT MSO_TO_OO_I_PageSetup_Initialize(
     I_PageSetup* pPageSetup,
     I_Worksheet* wsh);

HRESULT MSO_TO_OO_I_Workbook_Initialize(
        I_Workbook* iface,
        I_Workbooks *pwrks);

HRESULT MSO_TO_OO_I_Workbook_Initialize2(
        I_Workbook* iface,
        I_Workbooks *pwrks,
        BSTR Filename,
        VARIANT_BOOL astemplate);

HRESULT MSO_TO_OO_I_Sheets_Initialize(
        I_Sheets* iface,
        I_Workbook *wb);

HRESULT MSO_TO_OO_GetActiveSheet(
        I_Sheets* iface,
        I_Worksheet **wsh);

HRESULT MSO_TO_OO_I_Worksheet_Initialize(
        I_Worksheet* iface,
        I_Workbook *wb,
        IDispatch *oosheet);

HRESULT MSO_TO_OO_GetDispatchPropertyValue(
        _Application *app,
        IDispatch** pIDispatch);

HRESULT MSO_TO_OO_GetDispatchHelper(
        _Application *app,
        IDispatch** pIDispatch);

HRESULT MSO_TO_OO_ExecuteDispatchHelper_ActiveWorkBook(
        _Application *app,
        BSTR ooCommand,
        VARIANT ooParams);

HRESULT MSO_TO_OO_ExecuteDispatchHelper_WB(
        I_Workbook *wb,
        BSTR ooCommand,
        VARIANT ooParams);

HRESULT MSO_TO_OO_CloseWorkbook(
        I_Workbook *wb,
        BSTR filename);

HRESULT MSO_TO_OO_I_Range_Initialize(
        I_Range *iface,
        I_Range *pParentRange,
        struct CELL_COORD topLeft,
        struct CELL_COORD bottomRight);

HRESULT MSO_TO_OO_I_Range_Initialize2(
        I_Range* iface,
        IDispatch *oosheet);

HRESULT MSO_TO_OO_I_Range_Initialize3(
        I_Range* iface,
        IDispatch *oosheet,
        IDispatch *psheet,
        IDispatch *pApp);

HRESULT MSO_TO_OO_GetRangeAddress(
        I_Range* iface,
        LONG *lLeft,
        LONG *lTop,
        LONG *lRight,
        LONG *lBottom);

HRESULT MSO_TO_OO_GetActiveWorkbook(
        I_Workbooks* iface,
        I_Workbook **wb);

HRESULT MSO_TO_OO_GetActiveCells(
        I_Workbooks* iface,
        I_Range **ppRange);

HRESULT MSO_TO_OO_I_Range_Initialize_ByName(
        I_Range *iface,
        I_Range *pParentRange,
        VARIANT rangename);

HRESULT MSO_TO_OO_CorrectArg(
        VARIANT value,
        VARIANT *retval);

HRESULT MSO_TO_OO_MakeURLFromFilename(
        BSTR value,
        BSTR *retval);

LONG MSO_TO_OO_FindIndexWorksheetByName(
        I_Sheets *iface,
        BSTR name);

/*возвращает Workbook и индекс*/
LONG MSO_TO_OO_GlobalFindIndexWorksheetByName(
        _Application *app,
        BSTR name,
        IDispatch **retval);

HRESULT MSO_TO_OO_I_Shapes_Initialize(
        I_Shapes* iface,
        I_Worksheet *iwsh);

HRESULT MSO_TO_OO_I_Shape_Line_Initialize(
        I_Shape* iface,
        I_Shapes *ishapes,
        float x1, float y1, float x2, float y2);

HRESULT MSO_TO_OO_Names_Initialize(
        Names* iface,
        I_Workbook *wb);

HRESULT MSO_TO_OO_Workbook_SetVisible(
        I_Workbook *wb,
        VARIANT_BOOL vbvisible);

HRESULT MSO_TO_OO_I_Outline_Initialize(
        I_Outline* iface,
        I_Worksheet *iwsh);

HRESULT MSO_TO_OO_Name_Initialize_By_Name(
        Name* iface,
        Names *pnames,
        VARIANT varname);

HRESULT MSO_TO_OO_Name_Initialize_By_Index(
        Name* iface,
        Names *pnames,
        VARIANT varindex);

BOOL    Is_Variant_Null(
        VARIANT var);

HRESULT MSO_TO_OO_ApplyWorkbookUiState(
        _ApplicationImpl *app,
        I_Workbook *workbook);

HRESULT MSO_TO_OO_ApplyApplicationUiState(
        _ApplicationImpl *app);

BSTR MSO_TO_OO_SwapFormulaSeparators(
        const WCHAR *text,
        BOOL to_uno);

HRESULT MSO_TO_OO_ConvertFormula(
        IDispatch *document,
        IDispatch *position,
        const WCHAR *text,
        LONG from,
        LONG to,
        BSTR *out);

LONG MSO_TO_OO_SheetColumnCount(
        IDispatch *oosheet);

HRESULT MSO_TO_OO_ActivateWorkbook(
        I_Workbook *workbook);

HRESULT MSO_TO_OO_GetWorkbookWindowState(
        I_Workbook *workbook,
        XlWindowState *state);

HRESULT MSO_TO_OO_SetWorkbookWindowState(
        I_Workbook *workbook,
        XlWindowState state);

HRESULT MSO_TO_OO_PrintPreview(
        I_Workbook *workbook);

HRESULT MSO_TO_OO_Windows_Create(
        _Application *app,
        I_Workbook *workbook,
        IDispatch **result);

HRESULT MSO_TO_OO_Window_Create(
        I_Workbook *workbook,
        IDispatch **result);
