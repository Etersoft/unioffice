// application_excel.h

#ifndef  APPLICATION_EXCEL_H
#define  APPLICATION_EXCEL_H

#define COBJMACROS

#include <windows.h>
#include <windef.h>
#include <winbase.h>
#include <ole2.h>
#include <ocidl.h>
#include <oaidl.h>
#include <stddef.h>
#include "unioffice_excel.h"

/*****************************************************************************
 * _Application interface
 */
 
#ifndef __Disp_Application_FWD_DEFINED__
#define __Disp_Application_FWD_DEFINED__
typedef interface Disp_Application Disp_Application;
#endif //__Disp_Application_FWD_DEFINED__
 
 
//#if defined(__cplusplus) && !defined(CINTERFACE)

#define CINTERFACE

#ifdef CINTERFACE
typedef struct Disp_ApplicationVtbl {

    /*** IUnknown methods ***/
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(
        Disp_Application* This,
        REFIID riid,
        void **ppvObject);

    ULONG (STDMETHODCALLTYPE *AddRef)(
        Disp_Application* This);

    ULONG (STDMETHODCALLTYPE *Release)(
        Disp_Application* This);

    /*** IDispatch methods ***/
    HRESULT (STDMETHODCALLTYPE *GetTypeInfoCount)(
        Disp_Application* This,
        UINT *pctinfo);

    HRESULT (STDMETHODCALLTYPE *GetTypeInfo)(
        Disp_Application* This,
        UINT iTInfo,
        LCID lcid,
        ITypeInfo **ppTInfo);

    HRESULT (STDMETHODCALLTYPE *GetIDsOfNames)(
        Disp_Application* This,
        REFIID riid,
        LPOLESTR *rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgDispId);

    HRESULT (STDMETHODCALLTYPE *Invoke)(
        Disp_Application* This,
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pVarResult,
        EXCEPINFO *pExcepInfo,
        UINT *puArgErr);

    /*** _Application methods ***/
    IDispatch* (STDMETHODCALLTYPE *get_Application)(
        Disp_Application* This);

    XlCreator (STDMETHODCALLTYPE *get_Creator)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_Parent)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_ActiveCell)(
        Disp_Application* This,
        IDispatch **RHS);

    IDispatch* (STDMETHODCALLTYPE *get_ActiveChart)(
        Disp_Application* This);

     IDispatch* (STDMETHODCALLTYPE *get_ActiveDialog)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_ActiveMenuBar)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get_ActivePrinter)(
        Disp_Application* This,
        LCID lcid);

    BSTR (STDMETHODCALLTYPE *put_ActivePrinter)(
        Disp_Application* This,
        LCID lcid);

    IDispatch* (STDMETHODCALLTYPE *get_ActiveSheet)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_ActiveWindow)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_ActiveWorkbook)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_AddIns)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_Assistant)(
        Disp_Application* This,
        IDispatch **RHS);

    void (STDMETHODCALLTYPE *Calculate)(
        Disp_Application* This,
        LCID lcid);

    IDispatch* (STDMETHODCALLTYPE *get_Cells)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_Charts)(
        Disp_Application* This);

    IDispatch* (STDMETHODCALLTYPE *get_Columns)(
        Disp_Application* This,
        VARIANT param);

    IDispatch* (STDMETHODCALLTYPE *get_CommandBars)(
        Disp_Application* This);

    long (STDMETHODCALLTYPE *get_DDEAppReturnCode)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *DDEExecute)(
        Disp_Application* This,
        long Channel,
        BSTR String,
        LCID lcid);

    long (STDMETHODCALLTYPE *DDEInitiate)(
        Disp_Application* This,
        BSTR App,
        BSTR Topic,
        LCID lcid);

    void (STDMETHODCALLTYPE *DDEPoke)(
        Disp_Application* This,
        long Channel,
        VARIANT Item,
        VARIANT Data,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *DDERequest)(
        Disp_Application* This,
        long Channel,
        BSTR Item,
        LCID lcid);

    HRESULT (STDMETHODCALLTYPE *DDETerminate)(
        Disp_Application* This,
        long Channel,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_DialogSheets)(
        Disp_Application* This);

    VARIANT (STDMETHODCALLTYPE *Evaluate)(
        Disp_Application* This,
        VARIANT Name,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *_Evaluate)(
        Disp_Application* This,
        VARIANT Name,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *ExecuteExcel4Macro)(
        Disp_Application* This,
        BSTR String,
        LCID lcid);

    IDispatch* (STDMETHODCALLTYPE *Intersect)(
        Disp_Application* This,
        IDispatch *Arg1,
        IDispatch *Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_MenuBars)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Modules)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Names)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Range)(
        Disp_Application* This,
        VARIANT Cell1,
        VARIANT Cell2);

    IDispatch * (STDMETHODCALLTYPE *get_Rows)(
        Disp_Application* This,
        VARIANT param);

    VARIANT (STDMETHODCALLTYPE *Run)(
        Disp_Application* This,
        VARIANT Macro,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30);

    VARIANT (STDMETHODCALLTYPE *_Run2)(
        Disp_Application* This,
        VARIANT Macro,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_Selection)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *SendKeys)(
        Disp_Application* This,
        VARIANT Keys,
        VARIANT Wait,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_Sheets)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_ShortcutMenus)(
        Disp_Application* This,
        long Index);

    IDispatch * (STDMETHODCALLTYPE *get_ThisWorkbook)(
        Disp_Application* This,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_Toolbars)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *Union)(
        Disp_Application* This,
        IDispatch *Arg1,
        IDispatch *Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_Windows)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Workbooks)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_WorksheetFunction)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Worksheets)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Excel4IntlMacroSheets)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Excel4MacroSheets)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *ActivateMicrosoftApp)(
        Disp_Application* This,
        XlMSApplication Index,
        LCID lcid);

    void (STDMETHODCALLTYPE *AddChartAutoFormat)(
        Disp_Application* This,
        VARIANT Chart,
        BSTR Name,
        VARIANT Description,
        LCID lcid);

    void (STDMETHODCALLTYPE *AddCustomList)(
        Disp_Application* This,
        VARIANT ListArray,
        VARIANT ByRow,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_AlertBeforeOverwriting)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_AlertBeforeOverwriting)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    BSTR (STDMETHODCALLTYPE *get_AltStartupPath)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_AltStartupPath)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_AskToUpdateLinks)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_AskToUpdateLinks)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_EnableAnimations)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_EnableAnimations)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    IDispatch * (STDMETHODCALLTYPE *get_AutoCorrect)(
        Disp_Application* This);

    long (STDMETHODCALLTYPE *get_Build)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_CalculateBeforeSave)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_CalculateBeforeSave)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    XlCalculation (STDMETHODCALLTYPE *get_Calculation)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Calculation)(
        Disp_Application* This,
        LCID lcid,
        XlCalculation RHS);

    VARIANT (STDMETHODCALLTYPE *get_Caller)(
        Disp_Application* This,
        VARIANT Index,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_CanPlaySounds)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_CanRecordSounds)(
        Disp_Application* This,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *get_Caption)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_Caption)(
        Disp_Application* This,
        VARIANT vName);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_CellDragAndDrop)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_CellDragAndDrop)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    double (STDMETHODCALLTYPE *CentimetersToPoints)(
        Disp_Application* This,
        double Centimeters,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *CheckSpelling)(
        Disp_Application* This,
        BSTR Word,
        VARIANT CustomDictionary,
        VARIANT IgnoreUppercase,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *get_ClipboardFormats)(
        Disp_Application* This,
        VARIANT Index,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayClipboardWindow)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DisplayClipboardWindow)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ColorButtons)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ColorButtons)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    XlCommandUnderlines (STDMETHODCALLTYPE *get_CommandUnderlines)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_CommandUnderlines)(
        Disp_Application* This,
        LCID lcid,
        XlCommandUnderlines RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ConstrainNumeric)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_ConstrainNumeric)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT (STDMETHODCALLTYPE *ConvertFormula)(
        Disp_Application* This,
        VARIANT Formula,
        XlReferenceStyle FromReferenceStyle,
        VARIANT ToReferenceStyle,
        VARIANT ToAbsolute,
        VARIANT RelativeTo,
        long Lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_CopyObjectsWithCells)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_CopyObjectsWithCells)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    XlMousePointer (STDMETHODCALLTYPE *get_Cursor)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Cursor)(
        Disp_Application* This,
        LCID lcid,
        XlMousePointer RHS);

    long (STDMETHODCALLTYPE *get_CustomListCount)(
        Disp_Application* This,
        LCID lcid);

    XlCutCopyMode (STDMETHODCALLTYPE *get_CutCopyMode)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_CutCopyMode)(
        Disp_Application* This,
        LCID lcid,
        XlCutCopyMode RHS);

    long (STDMETHODCALLTYPE *get_DataEntryMode)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DataEntryMode)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    VARIANT (STDMETHODCALLTYPE *Dummy1)(
        Disp_Application* This,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4);

    VARIANT (STDMETHODCALLTYPE *Dummy2)(
        Disp_Application* This,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8);

    VARIANT (STDMETHODCALLTYPE *Dummy3)(
        Disp_Application* This);

    VARIANT (STDMETHODCALLTYPE *Dummy4)(
        Disp_Application* This,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15);

    VARIANT (STDMETHODCALLTYPE *Dummy5)(
        Disp_Application* This,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13);

    VARIANT (STDMETHODCALLTYPE *Dummy6)(
        Disp_Application* This);

    VARIANT (STDMETHODCALLTYPE *Dummy7)(
        Disp_Application* This);

    VARIANT (STDMETHODCALLTYPE *Dummy8)(
        Disp_Application* This,
        VARIANT Arg1);

    VARIANT (STDMETHODCALLTYPE *Dummy9)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *Dummy10)(
        Disp_Application* This,
        VARIANT arg);

    void (STDMETHODCALLTYPE *Dummy11)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get__Default)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get_DefaultFilePath)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DefaultFilePath)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    void (STDMETHODCALLTYPE *DeleteChartAutoFormat)(
        Disp_Application* This,
        BSTR Name,
        LCID lcid);

    void (STDMETHODCALLTYPE *DeleteCustomList)(
        Disp_Application* This,
        long ListNum,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_Dialogs)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayAlerts)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL vbDisplayAlerts);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayAlerts)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayFormulaBar)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DisplayFormulaBar)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayFullScreen)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DisplayFullScreen)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayNoteIndicator)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayNoteIndicator)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    XlCommentDisplayMode (STDMETHODCALLTYPE *get_DisplayCommentIndicator)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayCommentIndicator)(
        Disp_Application* This,
        XlCommentDisplayMode RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayExcel4Menus)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DisplayExcel4Menus)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayRecentFiles)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayRecentFiles)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayScrollBars)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DisplayScrollBars)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayStatusBar)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DisplayStatusBar)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    void (STDMETHODCALLTYPE *DoubleClick)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_EditDirectlyInCell)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_EditDirectlyInCell)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_EnableAutoComplete)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_EnableAutoComplete)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    XlEnableCancelKey (STDMETHODCALLTYPE *get_EnableCancelKey)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_EnableCancelKey)(
        Disp_Application* This,
        LCID lcid,
        XlEnableCancelKey RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_EnableSound)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_EnableSound)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_EnableTipWizard)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_EnableTipWizard)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT (STDMETHODCALLTYPE *get_FileConverters)(
        Disp_Application* This,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_FileSearch)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_FileFind)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *_FindFile)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_FixedDecimal)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_FixedDecimal)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    long (STDMETHODCALLTYPE *get_FixedDecimalPlaces)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_FixedDecimalPlaces)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    VARIANT (STDMETHODCALLTYPE *GetCustomListContents)(
        Disp_Application* This,
        long ListNum,
        LCID lcid);

    long (STDMETHODCALLTYPE *GetCustomListNum)(
        Disp_Application* This,
        VARIANT ListArray,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *GetOpenFilename)(
        Disp_Application* This,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        VARIANT MultiSelect,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *GetSaveAsFilename)(
        Disp_Application* This,
        VARIANT InitialFilename,
        VARIANT FileFilter,
        VARIANT FilterIndex,
        VARIANT Title,
        VARIANT ButtonText,
        LCID lcid);

    void (STDMETHODCALLTYPE *Goto)(
        Disp_Application* This,
        VARIANT Reference,
        VARIANT Scroll,
        LCID lcid);

    double (STDMETHODCALLTYPE *get_Height)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Height)(
        Disp_Application* This,
        LCID lcid,
        double RHS);

    void (STDMETHODCALLTYPE *Help)(
        Disp_Application* This,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_IgnoreRemoteRequests)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_IgnoreRemoteRequests)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    double (STDMETHODCALLTYPE *InchesToPoints)(
        Disp_Application* This,
        double Inches,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *InputBox)(
        Disp_Application* This,
        BSTR Prompt,
        VARIANT Title,
        VARIANT Default,
        VARIANT Left,
        VARIANT Top,
        VARIANT HelpFile,
        VARIANT HelpContextID,
        VARIANT Type,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_Interactive)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Interactive)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT (STDMETHODCALLTYPE *get_International)(
        Disp_Application* This,
        VARIANT Index,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_Iteration)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Iteration)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_LargeButtons)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_LargeButtons)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    double (STDMETHODCALLTYPE *get_Left)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Left)(
        Disp_Application* This,
        LCID lcid,
        double RHS);

    BSTR (STDMETHODCALLTYPE *get_LibraryPath)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *MacroOptions)(
        Disp_Application* This,
        VARIANT Macro,
        VARIANT Description,
        VARIANT HasMenu,
        VARIANT MenuText,
        VARIANT HasShortcutKey,
        VARIANT ShortcutKey,
        VARIANT Category,
        VARIANT StatusBar,
        VARIANT HelpContextID,
        VARIANT HelpFile,
        LCID lcid);

    void (STDMETHODCALLTYPE *MailLogoff)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *MailLogon)(
        Disp_Application* This,
        VARIANT Name,
        VARIANT Password,
        VARIANT DownloadNewMail,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *get_MailSession)(
        Disp_Application* This,
        LCID lcid);

    XlMailSystem (STDMETHODCALLTYPE *get_MailSystem)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_MathCoprocessorAvailable)(
        Disp_Application* This,
        LCID lcid);

    double (STDMETHODCALLTYPE *get_MaxChange)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_MaxChange)(
        Disp_Application* This,
        LCID lcid,
        double RHS);

    long (STDMETHODCALLTYPE *get_MaxIterations)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_MaxIterations)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    long (STDMETHODCALLTYPE *get_MemoryFree)(
        Disp_Application* This,
        LCID lcid);

    long (STDMETHODCALLTYPE *get_MemoryTotal)(
        Disp_Application* This,
        LCID lcid);

    long (STDMETHODCALLTYPE *get_MemoryUsed)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_MouseAvailable)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_MoveAfterReturn)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_MoveAfterReturn)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    XlDirection (STDMETHODCALLTYPE *get_MoveAfterReturnDirection)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_MoveAfterReturnDirection)(
        Disp_Application* This,
        LCID lcid,
        XlDirection RHS);

    IDispatch * (STDMETHODCALLTYPE *get_RecentFiles)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get_Name)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *NextLetter)(
        Disp_Application* This,
        LCID lcid);

    BSTR (STDMETHODCALLTYPE *get_NetworkTemplatesPath)(
        Disp_Application* This,
        LCID lcid);

    IDispatch * (STDMETHODCALLTYPE *get_ODBCErrors)(
        Disp_Application* This);

    long (STDMETHODCALLTYPE *get_ODBCTimeout)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ODBCTimeout)(
        Disp_Application* This,
        long RHS);

    BSTR (STDMETHODCALLTYPE *get_OnCalculate)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_OnCalculate)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    BSTR (STDMETHODCALLTYPE *get_OnData)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_OnData)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    BSTR (STDMETHODCALLTYPE *get_OnDoubleClick)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_OnDoubleClick)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    BSTR (STDMETHODCALLTYPE *get_OnEntry)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_OnEntry)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    void (STDMETHODCALLTYPE *OnKey)(
        Disp_Application* This,
        BSTR Key,
        VARIANT Procedure,
        LCID lcid);

    void (STDMETHODCALLTYPE *OnRepeat)(
        Disp_Application* This,
        BSTR Text,
        BSTR Procedure,
        LCID lcid);

    BSTR (STDMETHODCALLTYPE *get_OnSheetActivate)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_OnSheetActivate)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    BSTR (STDMETHODCALLTYPE *get_OnSheetDeactivate)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_OnSheetDeactivate)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    void (STDMETHODCALLTYPE *OnTime)(
        Disp_Application* This,
        VARIANT EarliestTime,
        BSTR Procedure,
        VARIANT LatestTime,
        VARIANT Schedule,
        LCID lcid);

    void (STDMETHODCALLTYPE *OnUndo)(
        Disp_Application* This,
        BSTR Text,
        BSTR Procedure,
        LCID lcid);

    BSTR (STDMETHODCALLTYPE *get_OnWindow)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_OnWindow)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    BSTR (STDMETHODCALLTYPE *get_OperatingSystem)(
        Disp_Application* This,
        LCID lcid);

    BSTR (STDMETHODCALLTYPE *get_OrganizationName)(
        Disp_Application* This,
        LCID lcid);

    BSTR (STDMETHODCALLTYPE *get_Path)(
        Disp_Application* This,
        LCID lcid);

    BSTR (STDMETHODCALLTYPE *get_PathSeparator)(
        Disp_Application* This,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *get_PreviousSelections)(
        Disp_Application* This,
        VARIANT Index,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_PivotTableSelection)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_PivotTableSelection)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_PromptForSummaryInfo)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_PromptForSummaryInfo)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    void (STDMETHODCALLTYPE *Quit)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *RecordMacro)(
        Disp_Application* This,
        VARIANT BasicCode,
        VARIANT XlmCode,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_RecordRelative)(
        Disp_Application* This,
        LCID lcid);

    XlReferenceStyle (STDMETHODCALLTYPE *get_ReferenceStyle)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_ReferenceStyle)(
        Disp_Application* This,
        LCID lcid,
        XlReferenceStyle RHS);

    VARIANT (STDMETHODCALLTYPE *get_RegisteredFunctions)(
        Disp_Application* This,
        VARIANT Index1,
        VARIANT Index2,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *RegisterXLL)(
        Disp_Application* This,
        BSTR Filename,
        LCID lcid);

    void (STDMETHODCALLTYPE *Repeat)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *ResetTipWizard)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_RollZoom)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_RollZoom)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    void (STDMETHODCALLTYPE *Save)(
        Disp_Application* This,
        VARIANT Filename,
        LCID lcid);

    void (STDMETHODCALLTYPE *SaveWorkspace)(
        Disp_Application* This,
        VARIANT Filename,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ScreenUpdating)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_ScreenUpdating)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    void (STDMETHODCALLTYPE *SetDefaultChart)(
        Disp_Application* This,
        VARIANT FormatName,
        VARIANT Gallery);

    long (STDMETHODCALLTYPE *get_SheetsInNewWorkbook)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_SheetsInNewWorkbook)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ShowChartTipNames)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ShowChartTipNames)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ShowChartTipValues)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ShowChartTipValues)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    BSTR (STDMETHODCALLTYPE *get_StandardFont)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_StandardFont)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    double (STDMETHODCALLTYPE *get_StandardFontSize)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_StandardFontSize)(
        Disp_Application* This,
        LCID lcid,
        double RHS);

    BSTR (STDMETHODCALLTYPE *get_StartupPath)(
        Disp_Application* This,
        LCID lcid);

    VARIANT (STDMETHODCALLTYPE *get_StatusBar)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_StatusBar)(
        Disp_Application* This,
        LCID lcid,
        VARIANT RHS);

    BSTR (STDMETHODCALLTYPE *get_TemplatesPath)(
        Disp_Application* This,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ShowToolTips)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ShowToolTips)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    double (STDMETHODCALLTYPE *get_Top)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Top)(
        Disp_Application* This,
        LCID lcid,
        double RHS);

    XlFileFormat (STDMETHODCALLTYPE *get_DefaultSaveFormat)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DefaultSaveFormat)(
        Disp_Application* This,
        XlFileFormat RHS);

    BSTR (STDMETHODCALLTYPE *get_TransitionMenuKey)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_TransitionMenuKey)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    long (STDMETHODCALLTYPE *get_TransitionMenuKeyAction)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_TransitionMenuKeyAction)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_TransitionNavigKeys)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_TransitionNavigKeys)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    void (STDMETHODCALLTYPE *Undo)(
        Disp_Application* This,
        LCID lcid);

    double (STDMETHODCALLTYPE *get_UsableHeight)(
        Disp_Application* This,
        LCID lcid);

    double (STDMETHODCALLTYPE *get_UsableWidth)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_UserControl)(
        Disp_Application* This,
        VARIANT_BOOL vbUserControl);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_UserControl)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get_UserName)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_UserName)(
        Disp_Application* This,
        LCID lcid,
        BSTR RHS);

    BSTR (STDMETHODCALLTYPE *get_Value)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_VBE)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get_Version)(
        Disp_Application* This,
        long Lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_Visible)(
        Disp_Application* This,
        LCID Lcid);

    void (STDMETHODCALLTYPE *put_Visible)(
        Disp_Application* This,
        LCID Lcid,
        VARIANT_BOOL vbVisible);

    void (STDMETHODCALLTYPE *Volatile)(
        Disp_Application* This,
        VARIANT Volatile,
        LCID lcid);

    void (STDMETHODCALLTYPE *_Wait)(
        Disp_Application* This,
        VARIANT Time,
        LCID lcid);

    double (STDMETHODCALLTYPE *get_Width)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_Width)(
        Disp_Application* This,
        LCID lcid,
        double RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_WindowsForPens)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_WindowState)(
        Disp_Application* This,
        LCID lcid,
        XlWindowState State);

    XlWindowState (STDMETHODCALLTYPE *get_WindowState)(
        Disp_Application* This,
        LCID lcid);

    long (STDMETHODCALLTYPE *get_UILanguage)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_UILanguage)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    long (STDMETHODCALLTYPE *get_DefaultSheetDirection)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_DefaultSheetDirection)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    long (STDMETHODCALLTYPE *get_CursorMovement)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_CursorMovement)(
        Disp_Application* This,
        LCID lcid,
        long RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ControlCharacters)(
        Disp_Application* This,
        LCID lcid);

    void (STDMETHODCALLTYPE *put_ControlCharacters)(
        Disp_Application* This,
        LCID lcid,
        VARIANT_BOOL RHS);

    VARIANT (STDMETHODCALLTYPE *_WSFunction)(
        Disp_Application* This,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_EnableEvents)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_EnableEvents)(
        Disp_Application* This,
        VARIANT_BOOL vbee);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayInfoWindow)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayInfoWindow)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *Wait)(
        Disp_Application* This,
        VARIANT Time,
        LCID lcid);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ExtendList)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ExtendList)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    IDispatch * (STDMETHODCALLTYPE *get_OLEDBErrors)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *GetPhonetic)(
        Disp_Application* This,
        VARIANT Text);

    IDispatch * (STDMETHODCALLTYPE *get_COMAddIns)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_DefaultWebOptions)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get_ProductCode)(
        Disp_Application* This);

    BSTR (STDMETHODCALLTYPE *get_UserLibraryPath)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_AutoPercentEntry)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_AutoPercentEntry)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    IDispatch * (STDMETHODCALLTYPE *get_LanguageSettings)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Dummy101)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *Dummy12)(
        Disp_Application* This,
        IDispatch *p1,
        IDispatch *p2);

    IDispatch * (STDMETHODCALLTYPE *get_AnswerWizard)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *CalculateFull)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *FindFile)(
        Disp_Application* This,
        LCID lcid);

    long (STDMETHODCALLTYPE *get_CalculationVersion)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ShowWindowsInTaskbar)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ShowWindowsInTaskbar)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    MsoFeatureInstall (STDMETHODCALLTYPE *get_FeatureInstall)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_FeatureInstall)(
        Disp_Application* This,
        MsoFeatureInstall RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_Ready)(
        Disp_Application* This);

    VARIANT (STDMETHODCALLTYPE *Dummy13)(
        Disp_Application* This,
        VARIANT Arg1,
        VARIANT Arg2,
        VARIANT Arg3,
        VARIANT Arg4,
        VARIANT Arg5,
        VARIANT Arg6,
        VARIANT Arg7,
        VARIANT Arg8,
        VARIANT Arg9,
        VARIANT Arg10,
        VARIANT Arg11,
        VARIANT Arg12,
        VARIANT Arg13,
        VARIANT Arg14,
        VARIANT Arg15,
        VARIANT Arg16,
        VARIANT Arg17,
        VARIANT Arg18,
        VARIANT Arg19,
        VARIANT Arg20,
        VARIANT Arg21,
        VARIANT Arg22,
        VARIANT Arg23,
        VARIANT Arg24,
        VARIANT Arg25,
        VARIANT Arg26,
        VARIANT Arg27,
        VARIANT Arg28,
        VARIANT Arg29,
        VARIANT Arg30);

    IDispatch * (STDMETHODCALLTYPE *get_FindFormat)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *putref_FindFormat)(
        Disp_Application* This,
        IDispatch *RHS);

    IDispatch * (STDMETHODCALLTYPE *get_ReplaceFormat)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *putref_ReplaceFormat)(
        Disp_Application* This,
        IDispatch *RHS);

    IDispatch * (STDMETHODCALLTYPE *get_UsedObjects)(
        Disp_Application* This);

    XlCalculationState (STDMETHODCALLTYPE *get_CalculationState)(
        Disp_Application* This);

    XlCalculationInterruptKey (STDMETHODCALLTYPE *get_CalculationInterruptKey)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_CalculationInterruptKey)(
        Disp_Application* This,
        XlCalculationInterruptKey RHS);

    IDispatch * (STDMETHODCALLTYPE *get_Watches)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayFunctionToolTips)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayFunctionToolTips)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    MsoAutomationSecurity (STDMETHODCALLTYPE *get_AutomationSecurity)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_AutomationSecurity)(
        Disp_Application* This,
        MsoAutomationSecurity RHS);

    IDispatch * (STDMETHODCALLTYPE *get_FileDialog)(
        Disp_Application* This,
        MsoFileDialogType fileDialogType);

    void (STDMETHODCALLTYPE *Dummy14)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *CalculateFullRebuild)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayPasteOptions)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayPasteOptions)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayInsertOptions)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayInsertOptions)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_GenerateGetPivotData)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_GenerateGetPivotData)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    IDispatch * (STDMETHODCALLTYPE *get_AutoRecover)(
        Disp_Application* This);

    long (STDMETHODCALLTYPE *get_Hwnd)(
        Disp_Application* This);

    long (STDMETHODCALLTYPE *get_Hinstance)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *CheckAbort)(
        Disp_Application* This,
        VARIANT KeepAbort);

    IDispatch * (STDMETHODCALLTYPE *get_ErrorCheckingOptions)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_AutoFormatAsYouTypeReplaceHyperlinks)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_AutoFormatAsYouTypeReplaceHyperlinks)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    IDispatch * (STDMETHODCALLTYPE *get_SmartTagRecognizers)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_NewWorkbook)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_SpellingOptions)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_Speech)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_MapPaperSize)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_MapPaperSize)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ShowStartupDialog)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ShowStartupDialog)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    BSTR (STDMETHODCALLTYPE *get_DecimalSeparator)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DecimalSeparator)(
        Disp_Application* This,
        BSTR RHS);

    BSTR (STDMETHODCALLTYPE *get_ThousandsSeparator)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_ThousandsSeparator)(
        Disp_Application* This,
        BSTR RHS);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_UseSystemSeparators)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_UseSystemSeparators)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    IDispatch * (STDMETHODCALLTYPE *get_ThisCell)(
        Disp_Application* This);

    IDispatch * (STDMETHODCALLTYPE *get_RTD)(
        Disp_Application* This);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_DisplayDocumentActionTaskPane)(
        Disp_Application* This);

    void (STDMETHODCALLTYPE *put_DisplayDocumentActionTaskPane)(
        Disp_Application* This,
        VARIANT_BOOL RHS);

    void (STDMETHODCALLTYPE *DisplayXMLSourcePane)(
        Disp_Application* This,
        VARIANT XmlMap);

    VARIANT_BOOL (STDMETHODCALLTYPE *get_ArbitraryXMLSupportAvailable)(
        Disp_Application* This);

    VARIANT (STDMETHODCALLTYPE *Support)(
        Disp_Application* This,
        IDispatch *Object,
        long ID,
        VARIANT arg);

} Disp_ApplicationVtbl;

interface Disp_Application {
    CONST_VTBL Disp_ApplicationVtbl* lpVtbl;
};

#endif // (CINTERFACE)

#ifdef COBJMACROS

/*** IUnknown methods ***/
#define Disp_Application_QueryInterface(This,riid,ppvObject) (This)->lpVtbl->QueryInterface(This,riid,ppvObject)
#define Disp_Application_AddRef(This) (This)->lpVtbl->AddRef(This)
#define Disp_Application_Release(This) (This)->lpVtbl->Release(This)
/*** IDispatch methods ***/
#define Disp_Application_GetTypeInfoCount(This,pctinfo) (This)->lpVtbl->GetTypeInfoCount(This,pctinfo)
#define Disp_Application_GetTypeInfo(This,iTInfo,lcid,ppTInfo) (This)->lpVtbl->GetTypeInfo(This,iTInfo,lcid,ppTInfo)
#define Disp_Application_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) (This)->lpVtbl->GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)
#define Disp_Application_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) (This)->lpVtbl->Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)
/*** _Application methods ***/
//#define _Application_get_Application(This,value) (This)->lpVtbl->get_Application(This,value)





#endif  // COBJMACROS


#endif  // APPLICATION_EXCEL_H
