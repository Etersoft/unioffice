/*
 * Main header file
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

#define _WIN32_WINNT 0x0600

#ifndef __MSO_TO_OO_PRIVATE_H__
#define __MSO_TO_OO_PRIVATE_H__

#include <stdarg.h>

#define COBJMACROS

#include <windows.h>
#include <windef.h>
#include <winbase.h>
#include <winuser.h>
#include <winreg.h>
#include <ole2.h>
#include "mso_to_oo.h"
#include "main_constants.h"
#include "debug.h"


extern LONG dll_ref;

struct CELL_COORD
{
    long x;
    long y;
};

typedef struct
{
    const IClassFactoryVtbl *lpVtbl;
    LONG ref;
} ClassFactoryImpl;

typedef struct
{
    const I_BordersVtbl *_bordersVtbl;
    LONG ref;
    IDispatch *prange;           /*Указатель на Range*/
} BordersImpl;

typedef struct
{
    const I_BorderVtbl *_borderVtbl;
    LONG ref;
    IDispatch *pborders;           /*Указатель на Borders*/
    XlBordersIndex key;
} BorderImpl;

typedef struct
{
    const I_InteriorVtbl *_interiorVtbl;
    LONG ref;
    IDispatch *prange;           /*Указатель на Range*/
} InteriorImpl;

typedef struct
{
    const I_WorksheetVtbl *_worksheetVtbl;
    LONG ref;
    IDispatch *pOOSheet;      /*Указатель на Sheet из OpenOffice*/
    IDispatch *pwb;           /*Указатель на Parent Workbook*/
    IDispatch *pAllRange;
} WorksheetImpl;

typedef struct
{
    const I_SheetsVtbl *_sheetsVtbl;
    LONG ref;
    IDispatch *pwb;           /*Указатель на Workbook*/
    IDispatch *pOOSheets;     /*Указатель на Sheets openoffice*/
} SheetsImpl;

typedef struct
{
    const I_RangeVtbl *_rangeVtbl;
    LONG ref;
    IDispatch *pOORange;     /*Указатель на Range openoffice*/
    IDispatch *pwsheet;      /*Указатель на worksheet*/
    IDispatch *pApplication;  /*Указатель на Application*/
} RangeImpl;

typedef struct
{
    const I_ShapesVtbl *_shapesVtbl;
    LONG ref;
    IDispatch *pOOPage;     /*Указатель на Range openoffice*/
    IDispatch *pwsheet;      /*Указатель на worksheet*/
    IDispatch *pApplication;  /*Указатель на Application*/
} ShapesImpl;

typedef struct
{
    const I_PageSetupVtbl *_pagesetupVtbl;
    LONG ref;
    IDispatch *pwsheet;      /*Указатель на worksheet*/
    IDispatch *pApplication;  /*Указатель на Application*/
} PageSetupImpl;

typedef struct
{
    const I_WorkbookVtbl *_workbookVtbl;
    LONG ref;
    IDispatch *pApplication;  /*Указатель на Application*/
    IDispatch *pDoc;          /*Указатель на Document*/
    IDispatch *pSheets;       /*Указатель на Sheets*/
//    BSTR filename;            /*имя файла*/ 
} WorkbookImpl;

typedef struct
{
    const I_WorkbooksVtbl *_workbooksVtbl;
    LONG ref;
    IDispatch *pApplication;  /*Указатель на Application*/
    int count_workbooks;      /*кол-во workbook*/
    IDispatch **pworkbook;    /*массив workbook*/
    int current_workbook;     /*текущий workbook*/
} WorkbooksImpl;

typedef struct
{
    const I_ApplicationExcelVtbl *_applicationexcellVtbl;
    LONG ref;
    IDispatch *pdOOApp;
    IDispatch *pdOODesktop;
    IDispatch *pdWorkbooks;
} _ApplicationExcelImpl;

typedef struct
{
    const I_FontVtbl *_ifontVtbl;
    LONG ref;
    IDispatch *prange;        /*указатель на range*/
} _FontImpl;

/*
 * Vtable interfaces and static instances
 */
extern const I_ApplicationExcelVtbl MSO_TO_OO__I_ApplicationExcel_Vtbl;
extern const I_FontVtbl MSO_TO_OO__I_Font_Vtbl;
extern const I_WorkbooksVtbl MSO_TO_OO_I_WorkbooksVtbl;
extern const I_WorkbookVtbl MSO_TO_OO_I_WorkbookVtbl;
extern const I_RangeVtbl MSO_TO_OO_I_RangeVtbl;
extern const I_PageSetupVtbl MSO_TO_OO_I_PageSetupVtbl;
extern const I_SheetsVtbl MSO_TO_OO_I_SheetsVtbl;
extern const I_WorksheetVtbl MSO_TO_OO_I_WorksheetVtbl;
extern const I_InteriorVtbl MSO_TO_OO_I_InteriorVtbl;
extern const I_BordersVtbl MSO_TO_OO_I_BordersVtbl;
extern const I_BorderVtbl MSO_TO_OO_I_BorderVtbl;
extern const I_ShapesVtbl MSO_TO_OO_I_ShapesVtbl;

extern ClassFactoryImpl OOFFICE_ClassFactory;


extern _ApplicationExcelImpl MSO_TO_OO__ApplicationExcel;
extern _FontImpl MSO_TO_OO__Font;
extern WorkbooksImpl MSO_TO_OO_Workbooks;
extern WorkbookImpl MSO_TO_OO_Workbook;
extern RangeImpl MSO_TO_OO_Range;
extern PageSetupImpl MSO_TO_OO_PageSetup;
extern SheetsImpl MSO_TO_OO_Sheets;
extern WorksheetImpl MSO_TO_OO_Worksheet;
extern InteriorImpl MSO_TO_OO_Interior;
extern BordersImpl MSO_TO_OO_Borders;
extern BorderImpl MSO_TO_OO_Border;
extern ShapesImpl MSO_TO_OO_Shapes;
/*Constructors*/

extern HRESULT _ApplicationExcelConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_FontConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_WorkbooksConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_WorkbookConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_RangeConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_SheetsConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_WorksheetConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_InteriorConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_BordersConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_BorderConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
extern HRESULT _I_ShapesConstructor(IUnknown *pUnkOuter, LPVOID *ppObj);
#endif /* __OOFFICE_PRIVATE_H__ */

