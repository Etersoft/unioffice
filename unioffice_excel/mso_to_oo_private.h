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
#include <ocidl.h>
#include <oaidl.h>
#include <stddef.h>
#include "unioffice_excel.h"
#include "main_constants.h"
#include "debug.h"
#include "dispid_const.h"

// Dispinterfaces
#include "application_excel.h"


#define VER_2 1
#define VER_3 2

extern LONG dll_ref;
extern LONG OOVersion;

struct CELL_COORD
{
    long x;
    long y;
};

typedef struct
{
    const IClassFactoryVtbl *pclassfactoryVtbl;
    LONG ref;
} ClassFactoryImpl;

#define CLASSFACTORY_CLASSFACTORY(x) ((IClassFactory*)&(x)->pclassfactoryVtbl)

typedef struct
{
    const I_BordersVtbl *pbordersVtbl;
    const IEnumVARIANTVtbl *penumeratorVtbl;

    LONG ref;
    I_Range *pRange;               /*Pointer to IRange*/
    IDispatch *pOORange;           /*Pointer to OpenOffice range interface*/
    int enum_position;

} BordersImpl;

#define BORDERS_BORDERS(x) ((I_Borders*)&(x)->pbordersVtbl)
#define BORDERS_ENUM(x) ((IEnumVARIANT*)&(x)->penumeratorVtbl)

typedef struct
{
    const I_BorderVtbl *pborderVtbl;
    LONG ref;
    I_Borders *pBorders;           /*pointer to Borders*/
    IDispatch *pOORange;           /*Pointer to OpenOffice range interface*/
    XlBordersIndex key;
} BorderImpl;

#define BORDER_BORDER(x) ((I_Border*)&(x)->pborderVtbl)

typedef struct
{
    const I_FontVtbl *pfontVtbl;
    LONG ref;
    I_Range *pRange;        /*Pointer to IRange*/
    IDispatch *pOORange;     /*Pointer to OO Range*/
} FontImpl;

#define FONT_FONT(x) ((I_Font*)&(x)->pfontVtbl)

typedef struct
{
    const I_InteriorVtbl *pinteriorVtbl;
    LONG ref;
    I_Range *pRange;           /*Pointer to IRange*/
    IDispatch *pOORange;       /*Pointer to OORange*/
} InteriorImpl;

#define INTERIOR_INTERIOR(x) ((I_Interior*)&(x)->pinteriorVtbl)

typedef struct
{
    const I_OutlineVtbl *poutlineVtbl;
    LONG ref;
    I_Worksheet *pWorksheet;           /*Pointer to IWorksheet*/
    IDispatch *pOOSheet;               /*Pointer to OOSheet*/
} OutlineImpl;

#define OUTLINE_OUTLINE(x) ((I_Outline*)&(x)->poutlineVtbl)

typedef struct
{
    const I_WorksheetVtbl *pworksheetVtbl;
    LONG ref;
    IDispatch *pOOSheet;                   /*pointer to Sheet �� OpenOffice*/
    I_Workbook *pwb;                       /*pointer to Parent Workbook*/
    IDispatch *pAllRange;
} WorksheetImpl;

typedef struct
{
    const I_SheetsVtbl *psheetsVtbl;
    const IEnumVARIANTVtbl *penumeratorVtbl;

    LONG ref;
    IDispatch *pwb;                        /*pointer to Workbook*/
    IDispatch *pOOSheets;                  /*pointer to Sheets openoffice*/
    int enum_position;

} SheetsImpl;

#define SHEETS_SHEETS(x) ((I_Sheets*)&(x)->psheetsVtbl)
#define SHEETS_ENUM(x) ((IEnumVARIANT*)&(x)->penumeratorVtbl)

typedef struct
{
    const I_RangeVtbl *prangeVtbl;
    LONG ref;
    IDispatch *pOORange;                /*Pointer to OO Range*/
    IDispatch *pwsheet;                 /*pointer to worksheet*/
    int is_release;
} RangeImpl;

typedef struct
{
    const I_ShapesVtbl *_shapesVtbl;
    LONG ref;
    IDispatch *pOOPage;                 /*pointer to DrawPage openoffice*/
    IDispatch *pwsheet;                 /*pointer worksheet*/
} ShapesImpl;

typedef struct
{
    const I_ShapeVtbl *_shapeVtbl;
    LONG ref;
    IDispatch *pOOShape;               /*pointer to OOShape openoffice*/
    IDispatch *pShapes;                /*pointer to Shapes*/
} ShapeImpl;

typedef struct
{
    const I_PageSetupVtbl *ppagesetupVtbl;
    LONG ref;
    I_Worksheet *pWorksheet;      /*Pointer to IWorksheet*/
    IDispatch *pOOSheet;          /*Pointer to OOSheet*/
    IDispatch *pOODocument;       /*Pointer to OODocument*/
} PageSetupImpl;

#define PAGESETUP_PAGESETUP(x) ((I_PageSetup*)&(x)->ppagesetupVtbl)

typedef struct
{
    const NamesVtbl *pnamesVtbl;
    const IEnumVARIANTVtbl *penumeratorVtbl;

    LONG ref;
    IDispatch *pWorkbook;              /*pointer to Workbook*/
    IDispatch *pOONames;               /*pointer to OpenOffice Names*/
    int enum_position;

} NamesImpl;

#define NAMES_NAMES(x) ((Names*)&(x)->pnamesVtbl)
#define NAMES_ENUM(x) ((IEnumVARIANT*)&(x)->penumeratorVtbl)

typedef struct
{
    const NameVtbl *pnameVtbl;
    LONG ref;
    IDispatch *pNames;              /*pointer to Names*/
    IDispatch *pOOName;             /*pointer to OpenOffice Name*/
} NameImpl;

#define NAME_NAME(x) ((Name*)&(x)->pnameVtbl)

typedef struct
{
    const I_WindowsVtbl *_windowsVtbl;
    LONG ref;
    _Application *pApplication;        /*pointer to Application*/
} WindowsImpl;

typedef struct
{
    const I_WindowVtbl *_windowVtbl;
    
    LONG ref;
    I_Windows *pWindows;                     /*pointer to IWindows*/
} WindowImpl; 

typedef struct
{
    const I_WorkbookVtbl *pworkbookVtbl;
    LONG ref;
    IDispatch *pworkbooks;                   /*pointer to Application*/
    IDispatch *pDoc;                         /*pointer to Document*/
    IDispatch *pSheets;                      /*pointer to Sheets*/
//    BSTR filename;             
} WorkbookImpl;


typedef struct
{
    const I_WorkbooksVtbl *pworkbooksVtbl;
    const IEnumVARIANTVtbl *penumeratorVtbl;
    
    LONG ref;
    _Application *pApplication;  /*Pointer to Application*/
    
    I_Workbook **pworkbook;            /*pointer to aray of workbook*/
    int count_workbooks;               /*amount of workbook*/
    int capasity_workbooks;            /*capasity (to work with memory) */
    int current_workbook;              /*index of current workbook*/
    
    int enum_position;
    
} WorkbooksImpl;

#define WORKBOOKS_WORKBOOKS(x) ((I_Workbooks*)&(x)->pworkbooksVtbl)
#define WORKBOOKS_ENUM(x) ((IEnumVARIANT*)&(x)->penumeratorVtbl)

typedef struct
{
    const _ApplicationVtbl              *pApplicationVtbl;
    const Disp_ApplicationVtbl          *pDispApplicationVtbl;
    const IConnectionPointContainerVtbl *pConnectionPointContainerVtbl;
    const IConnectionPointVtbl          *pConnectionPointVtbl;
    

    LONG ref;
    IDispatch *pdOOApp;
    IDispatch *pdOODesktop;
    IDispatch *pdWorkbooks;

    VARIANT_BOOL screenupdating;
    VARIANT_BOOL displayalerts;
    VARIANT_BOOL visible;
    long sheetsinnewworkbook;

} _ApplicationImpl;

#define APPEXCEL(x) ((_Application*) &(x)->pApplicationVtbl)
#define DISPAPPEXCEL(x) ((Disp_Application*) &(x)->pDispApplicationVtbl)
#define CONPOINTCONT(x) ((IConnectionPointContainer*) &(x)->pConnectionPointContainerVtbl)
#define CONPOINT(x) ((IConnectionPoint*) &(x)->pConnectionPointVtbl)



#define DEFINE_THIS(class,ifild,iface) ((class*)((BYTE*)(iface)-offsetof(class,p ## ifild ## Vtbl)))

/*
 * Vtable interfaces and static instances
 */

extern ClassFactoryImpl OOFFICE_ClassFactory;

/*Constructors*/

extern HRESULT _ApplicationConstructor(LPVOID *ppObj);
extern HRESULT _I_FontConstructor(LPVOID *ppObj);
extern HRESULT _I_WorkbooksConstructor(LPVOID *ppObj);
extern HRESULT _I_WorkbookConstructor(LPVOID *ppObj);
extern HRESULT _I_RangeConstructor(LPVOID *ppObj);
extern HRESULT _I_SheetsConstructor(LPVOID *ppObj);
extern HRESULT _I_WorksheetConstructor(LPVOID *ppObj);
extern HRESULT _I_InteriorConstructor(LPVOID *ppObj);
extern HRESULT _I_BordersConstructor(LPVOID *ppObj);
extern HRESULT _I_BorderConstructor(LPVOID *ppObj);
extern HRESULT _I_ShapesConstructor(LPVOID *ppObj);
extern HRESULT _I_ShapeConstructor(LPVOID *ppObj);
extern HRESULT _NamesConstructor(LPVOID *ppObj);
extern HRESULT _NameConstructor(LPVOID *ppObj);
extern HRESULT _I_OutlineConstructor(LPVOID *ppObj);
extern HRESULT _I_WindowsConstructor(LPVOID *ppObj);
extern HRESULT _I_WindowConstructor(LPVOID *ppObj);
#endif /* __OOFFICE_PRIVATE_H__ */

