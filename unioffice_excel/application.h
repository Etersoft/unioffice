#ifndef __UNIOFFICE_EXCEL_APPLICATION_H__
#define __UNIOFFICE_EXCEL_APPLICATION_H__

#include "unioffice_excel_private.h"
#include "workbooks.h"

#include "../OOWrappers/oo_servicemanager.h"
#include "../OOWrappers/oo_desktop.h"


class Application : public _Application
{
public:
           
       // IUnknown
       virtual HRESULT STDMETHODCALLTYPE QueryInterface(const IID& iid, void** ppv);
       virtual ULONG STDMETHODCALLTYPE AddRef();
       virtual ULONG STDMETHODCALLTYPE Release();
         
       // IDispatch    
       virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount( UINT * pctinfo );
       virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo);
       virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId);
       virtual HRESULT STDMETHODCALLTYPE Invoke(
               DISPID dispIdMember,
               REFIID riid,
               LCID lcid,
               WORD wFlags,
               DISPPARAMS * pDispParams,
               VARIANT * pVarResult,
               EXCEPINFO * pExcepInfo,
               UINT * puArgErr); 
                  
        // _Application
        virtual HRESULT STDMETHODCALLTYPE get_Application( 
            Application	**RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE get_Creator( 
             XlCreator *RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE get_Parent( 
             Application	**RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE get_ActiveCell( 
             Range **RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE get_ActiveChart( 
             Chart **RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE get_ActiveDialog( 
             DialogSheet **RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE get_ActiveMenuBar( 
             MenuBar **RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE get_ActivePrinter( 
             long lcid,
             BSTR *RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE put_ActivePrinter( 
             long lcid,
             BSTR RHS) ;
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ActiveSheet( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ActiveWindow( 
            /* [retval][out] */ Window **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ActiveWorkbook( 
            /* [retval][out] */ Workbook **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AddIns( 
            /* [retval][out] */ AddIns **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Assistant( 
            /* [retval][out] */ Assistant **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Calculate( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Cells( 
            /* [retval][out] */ Range **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Charts( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Columns( 
            /* [retval][out] */ Range **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CommandBars( 
            /* [retval][out] */ CommandBars **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DDEAppReturnCode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DDEExecute( 
            /* [in] */ long Channel,
            /* [in] */ BSTR String,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DDEInitiate( 
            /* [in] */ BSTR App,
            /* [in] */ BSTR Topic,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DDEPoke( 
            /* [in] */ long Channel,
            /* [in] */ VARIANT Item,
            /* [in] */ VARIANT Data,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DDERequest( 
            /* [in] */ long Channel,
            /* [in] */ BSTR Item,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DDETerminate( 
            /* [in] */ long Channel,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_DialogSheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE _Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ExecuteExcel4Macro( 
            /* [in] */ BSTR String,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Intersect( 
            /* [in] */ Range *Arg1,
            /* [in] */ Range *Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [optional][in] */ VARIANT Arg16,
            /* [optional][in] */ VARIANT Arg17,
            /* [optional][in] */ VARIANT Arg18,
            /* [optional][in] */ VARIANT Arg19,
            /* [optional][in] */ VARIANT Arg20,
            /* [optional][in] */ VARIANT Arg21,
            /* [optional][in] */ VARIANT Arg22,
            /* [optional][in] */ VARIANT Arg23,
            /* [optional][in] */ VARIANT Arg24,
            /* [optional][in] */ VARIANT Arg25,
            /* [optional][in] */ VARIANT Arg26,
            /* [optional][in] */ VARIANT Arg27,
            /* [optional][in] */ VARIANT Arg28,
            /* [optional][in] */ VARIANT Arg29,
            /* [optional][in] */ VARIANT Arg30,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_MenuBars( 
            /* [retval][out] */ MenuBars **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Modules( 
            /* [retval][out] */ Modules **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Names( 
            /* [retval][out] */ Names **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Range( 
            /* [in] */ VARIANT Cell1,
            /* [optional][in] */ VARIANT Cell2,
            /* [retval][out] */ Range **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Rows( 
            /* [retval][out] */ Range **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Run( 
            /* [optional][in] */ VARIANT Macro,
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [optional][in] */ VARIANT Arg16,
            /* [optional][in] */ VARIANT Arg17,
            /* [optional][in] */ VARIANT Arg18,
            /* [optional][in] */ VARIANT Arg19,
            /* [optional][in] */ VARIANT Arg20,
            /* [optional][in] */ VARIANT Arg21,
            /* [optional][in] */ VARIANT Arg22,
            /* [optional][in] */ VARIANT Arg23,
            /* [optional][in] */ VARIANT Arg24,
            /* [optional][in] */ VARIANT Arg25,
            /* [optional][in] */ VARIANT Arg26,
            /* [optional][in] */ VARIANT Arg27,
            /* [optional][in] */ VARIANT Arg28,
            /* [optional][in] */ VARIANT Arg29,
            /* [optional][in] */ VARIANT Arg30,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE _Run2( 
            /* [optional][in] */ VARIANT Macro,
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [optional][in] */ VARIANT Arg16,
            /* [optional][in] */ VARIANT Arg17,
            /* [optional][in] */ VARIANT Arg18,
            /* [optional][in] */ VARIANT Arg19,
            /* [optional][in] */ VARIANT Arg20,
            /* [optional][in] */ VARIANT Arg21,
            /* [optional][in] */ VARIANT Arg22,
            /* [optional][in] */ VARIANT Arg23,
            /* [optional][in] */ VARIANT Arg24,
            /* [optional][in] */ VARIANT Arg25,
            /* [optional][in] */ VARIANT Arg26,
            /* [optional][in] */ VARIANT Arg27,
            /* [optional][in] */ VARIANT Arg28,
            /* [optional][in] */ VARIANT Arg29,
            /* [optional][in] */ VARIANT Arg30,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Selection( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SendKeys( 
            /* [in] */ VARIANT Keys,
            /* [optional][in] */ VARIANT Wait,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Sheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShortcutMenus( 
            /* [in] */ long Index,
            /* [retval][out] */ Menu **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ThisWorkbook( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Workbook **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Toolbars( 
            /* [retval][out] */ Toolbars **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Union( 
            /* [in] */ Range *Arg1,
            /* [in] */ Range *Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [optional][in] */ VARIANT Arg16,
            /* [optional][in] */ VARIANT Arg17,
            /* [optional][in] */ VARIANT Arg18,
            /* [optional][in] */ VARIANT Arg19,
            /* [optional][in] */ VARIANT Arg20,
            /* [optional][in] */ VARIANT Arg21,
            /* [optional][in] */ VARIANT Arg22,
            /* [optional][in] */ VARIANT Arg23,
            /* [optional][in] */ VARIANT Arg24,
            /* [optional][in] */ VARIANT Arg25,
            /* [optional][in] */ VARIANT Arg26,
            /* [optional][in] */ VARIANT Arg27,
            /* [optional][in] */ VARIANT Arg28,
            /* [optional][in] */ VARIANT Arg29,
            /* [optional][in] */ VARIANT Arg30,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Windows( 
            /* [retval][out] */ Windows **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Workbooks( 
            /* [retval][out] */ Workbooks **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WorksheetFunction( 
            /* [retval][out] */ WorksheetFunction **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Worksheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Excel4IntlMacroSheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Excel4MacroSheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ActivateMicrosoftApp( 
            /* [in] */ XlMSApplication Index,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE AddChartAutoFormat( 
            /* [in] */ VARIANT Chart,
            /* [in] */ BSTR Name,
            /* [optional][in] */ VARIANT Description,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE AddCustomList( 
            /* [in] */ VARIANT ListArray,
            /* [optional][in] */ VARIANT ByRow,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AlertBeforeOverwriting( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AlertBeforeOverwriting( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AltStartupPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AltStartupPath( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AskToUpdateLinks( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AskToUpdateLinks( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableAnimations( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableAnimations( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoCorrect( 
            /* [retval][out] */ AutoCorrect **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Build( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CalculateBeforeSave( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CalculateBeforeSave( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Calculation( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlCalculation *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Calculation( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlCalculation RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Caller( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CanPlaySounds( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CanRecordSounds( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Caption( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Caption( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CellDragAndDrop( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CellDragAndDrop( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CentimetersToPoints( 
            /* [in] */ double Centimeters,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CheckSpelling( 
            /* [in] */ BSTR Word,
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ClipboardFormats( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayClipboardWindow( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayClipboardWindow( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_ColorButtons( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_ColorButtons( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CommandUnderlines( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlCommandUnderlines *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CommandUnderlines( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlCommandUnderlines RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ConstrainNumeric( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ConstrainNumeric( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ConvertFormula( 
            /* [in] */ VARIANT Formula,
            /* [in] */ XlReferenceStyle FromReferenceStyle,
            /* [optional][in] */ VARIANT ToReferenceStyle,
            /* [optional][in] */ VARIANT ToAbsolute,
            /* [optional][in] */ VARIANT RelativeTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CopyObjectsWithCells( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CopyObjectsWithCells( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Cursor( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlMousePointer *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Cursor( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlMousePointer RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CustomListCount( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CutCopyMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlCutCopyMode *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CutCopyMode( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlCutCopyMode RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DataEntryMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DataEntryMode( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy1( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy2( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy3( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy4( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy5( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy6( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy7( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy8( 
            /* [optional][in] */ VARIANT Arg1,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy9( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy10( 
            /* [optional][in] */ VARIANT arg,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy11( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get__Default( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DefaultFilePath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DefaultFilePath( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE DeleteChartAutoFormat( 
            /* [in] */ BSTR Name,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DeleteCustomList( 
            /* [in] */ long ListNum,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Dialogs( 
            /* [retval][out] */ Dialogs **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayAlerts( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayAlerts( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayFormulaBar( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayFormulaBar( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayFullScreen( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayFullScreen( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayNoteIndicator( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayNoteIndicator( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayCommentIndicator( 
            /* [retval][out] */ XlCommentDisplayMode *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayCommentIndicator( 
            /* [in] */ XlCommentDisplayMode RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayExcel4Menus( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayExcel4Menus( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayRecentFiles( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayRecentFiles( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayScrollBars( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayScrollBars( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayStatusBar( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayStatusBar( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DoubleClick( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EditDirectlyInCell( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EditDirectlyInCell( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableAutoComplete( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableAutoComplete( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableCancelKey( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlEnableCancelKey *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableCancelKey( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlEnableCancelKey RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableSound( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableSound( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableTipWizard( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableTipWizard( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FileConverters( 
            /* [optional][in] */ VARIANT Index1,
            /* [optional][in] */ VARIANT Index2,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_FileSearch( 
            /* [retval][out] */ FileSearch **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_FileFind( 
            /* [retval][out] */ IFind **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _FindFile( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FixedDecimal( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_FixedDecimal( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FixedDecimalPlaces( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_FixedDecimalPlaces( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE GetCustomListContents( 
            /* [in] */ long ListNum,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE GetCustomListNum( 
            /* [in] */ VARIANT ListArray,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE GetOpenFilename( 
            /* [optional][in] */ VARIANT FileFilter,
            /* [optional][in] */ VARIANT FilterIndex,
            /* [optional][in] */ VARIANT Title,
            /* [optional][in] */ VARIANT ButtonText,
            /* [optional][in] */ VARIANT MultiSelect,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE GetSaveAsFilename( 
            /* [optional][in] */ VARIANT InitialFilename,
            /* [optional][in] */ VARIANT FileFilter,
            /* [optional][in] */ VARIANT FilterIndex,
            /* [optional][in] */ VARIANT Title,
            /* [optional][in] */ VARIANT ButtonText,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Goto( 
            /* [optional][in] */ VARIANT Reference,
            /* [optional][in] */ VARIANT Scroll,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Height( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Height( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Help( 
            /* [optional][in] */ VARIANT HelpFile,
            /* [optional][in] */ VARIANT HelpContextID,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_IgnoreRemoteRequests( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_IgnoreRemoteRequests( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE InchesToPoints( 
            /* [in] */ double Inches,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE InputBox( 
            /* [in] */ BSTR Prompt,
            /* [optional][in] */ VARIANT Title,
            /* [optional][in] */ VARIANT Default,
            /* [optional][in] */ VARIANT Left,
            /* [optional][in] */ VARIANT Top,
            /* [optional][in] */ VARIANT HelpFile,
            /* [optional][in] */ VARIANT HelpContextID,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Interactive( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Interactive( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_International( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Iteration( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Iteration( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_LargeButtons( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_LargeButtons( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Left( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Left( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_LibraryPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE MacroOptions( 
            /* [optional][in] */ VARIANT Macro,
            /* [optional][in] */ VARIANT Description,
            /* [optional][in] */ VARIANT HasMenu,
            /* [optional][in] */ VARIANT MenuText,
            /* [optional][in] */ VARIANT HasShortcutKey,
            /* [optional][in] */ VARIANT ShortcutKey,
            /* [optional][in] */ VARIANT Category,
            /* [optional][in] */ VARIANT StatusBar,
            /* [optional][in] */ VARIANT HelpContextID,
            /* [optional][in] */ VARIANT HelpFile,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE MailLogoff( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE MailLogon( 
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT DownloadNewMail,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MailSession( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MailSystem( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlMailSystem *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MathCoprocessorAvailable( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MaxChange( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_MaxChange( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MaxIterations( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_MaxIterations( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_MemoryFree( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_MemoryTotal( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_MemoryUsed( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MouseAvailable( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MoveAfterReturn( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_MoveAfterReturn( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MoveAfterReturnDirection( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlDirection *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_MoveAfterReturnDirection( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlDirection RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_RecentFiles( 
            /* [retval][out] */ RecentFiles **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE NextLetter( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Workbook **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_NetworkTemplatesPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ODBCErrors( 
            /* [retval][out] */ ODBCErrors **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ODBCTimeout( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ODBCTimeout( 
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnCalculate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnCalculate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnData( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnData( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE OnKey( 
            /* [in] */ BSTR Key,
            /* [optional][in] */ VARIANT Procedure,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE OnRepeat( 
            /* [in] */ BSTR Text,
            /* [in] */ BSTR Procedure,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE OnTime( 
            /* [in] */ VARIANT EarliestTime,
            /* [in] */ BSTR Procedure,
            /* [optional][in] */ VARIANT LatestTime,
            /* [optional][in] */ VARIANT Schedule,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE OnUndo( 
            /* [in] */ BSTR Text,
            /* [in] */ BSTR Procedure,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnWindow( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnWindow( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_OperatingSystem( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_OrganizationName( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Path( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PathSeparator( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PreviousSelections( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PivotTableSelection( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_PivotTableSelection( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PromptForSummaryInfo( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_PromptForSummaryInfo( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Quit( void);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RecordMacro( 
            /* [optional][in] */ VARIANT BasicCode,
            /* [optional][in] */ VARIANT XlmCode,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_RecordRelative( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ReferenceStyle( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlReferenceStyle *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ReferenceStyle( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlReferenceStyle RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_RegisteredFunctions( 
            /* [optional][in] */ VARIANT Index1,
            /* [optional][in] */ VARIANT Index2,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RegisterXLL( 
            /* [in] */ BSTR Filename,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Repeat( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE ResetTipWizard( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_RollZoom( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_RollZoom( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Save( 
            /* [optional][in] */ VARIANT Filename,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SaveWorkspace( 
            /* [optional][in] */ VARIANT Filename,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ScreenUpdating( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ScreenUpdating( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE SetDefaultChart( 
            /* [optional][in] */ VARIANT FormatName,
            /* [optional][in] */ VARIANT Gallery);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SheetsInNewWorkbook( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_SheetsInNewWorkbook( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowChartTipNames( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowChartTipNames( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowChartTipValues( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowChartTipValues( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_StandardFont( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_StandardFont( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_StandardFontSize( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_StandardFontSize( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_StartupPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_StatusBar( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_StatusBar( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TemplatesPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowToolTips( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowToolTips( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Top( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Top( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DefaultSaveFormat( 
            /* [retval][out] */ XlFileFormat *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DefaultSaveFormat( 
            /* [in] */ XlFileFormat RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TransitionMenuKey( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_TransitionMenuKey( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TransitionMenuKeyAction( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_TransitionMenuKeyAction( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TransitionNavigKeys( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_TransitionNavigKeys( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Undo( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UsableHeight( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UsableWidth( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UserControl( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_UserControl( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UserName( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_UserName( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_VBE( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Version( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Volatile( 
            /* [optional][in] */ VARIANT Volatile,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _Wait( 
            /* [in] */ VARIANT Time,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Width( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Width( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WindowsForPens( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WindowState( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlWindowState *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_WindowState( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlWindowState RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_UILanguage( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_UILanguage( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DefaultSheetDirection( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DefaultSheetDirection( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CursorMovement( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CursorMovement( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ControlCharacters( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ControlCharacters( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _WSFunction( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [optional][in] */ VARIANT Arg16,
            /* [optional][in] */ VARIANT Arg17,
            /* [optional][in] */ VARIANT Arg18,
            /* [optional][in] */ VARIANT Arg19,
            /* [optional][in] */ VARIANT Arg20,
            /* [optional][in] */ VARIANT Arg21,
            /* [optional][in] */ VARIANT Arg22,
            /* [optional][in] */ VARIANT Arg23,
            /* [optional][in] */ VARIANT Arg24,
            /* [optional][in] */ VARIANT Arg25,
            /* [optional][in] */ VARIANT Arg26,
            /* [optional][in] */ VARIANT Arg27,
            /* [optional][in] */ VARIANT Arg28,
            /* [optional][in] */ VARIANT Arg29,
            /* [optional][in] */ VARIANT Arg30,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableEvents( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableEvents( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayInfoWindow( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayInfoWindow( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Wait( 
            /* [in] */ VARIANT Time,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ExtendList( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ExtendList( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_OLEDBErrors( 
            /* [retval][out] */ OLEDBErrors **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE GetPhonetic( 
            /* [optional][in] */ VARIANT Text,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_COMAddIns( 
            /* [retval][out] */ COMAddIns **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DefaultWebOptions( 
            /* [retval][out] */ DefaultWebOptions **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProductCode( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UserLibraryPath( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoPercentEntry( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AutoPercentEntry( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_LanguageSettings( 
            /* [retval][out] */ LanguageSettings **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Dummy101( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy12( 
            /* [in] */ PivotTable *p1,
            /* [in] */ PivotTable *p2);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AnswerWizard( 
            /* [retval][out] */ AnswerWizard **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CalculateFull( void);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FindFile( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CalculationVersion( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowWindowsInTaskbar( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowWindowsInTaskbar( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FeatureInstall( 
            /* [retval][out] */ MsoFeatureInstall *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_FeatureInstall( 
            /* [in] */ MsoFeatureInstall RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Ready( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy13( 
            /* [in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [optional][in] */ VARIANT Arg16,
            /* [optional][in] */ VARIANT Arg17,
            /* [optional][in] */ VARIANT Arg18,
            /* [optional][in] */ VARIANT Arg19,
            /* [optional][in] */ VARIANT Arg20,
            /* [optional][in] */ VARIANT Arg21,
            /* [optional][in] */ VARIANT Arg22,
            /* [optional][in] */ VARIANT Arg23,
            /* [optional][in] */ VARIANT Arg24,
            /* [optional][in] */ VARIANT Arg25,
            /* [optional][in] */ VARIANT Arg26,
            /* [optional][in] */ VARIANT Arg27,
            /* [optional][in] */ VARIANT Arg28,
            /* [optional][in] */ VARIANT Arg29,
            /* [optional][in] */ VARIANT Arg30,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FindFormat( 
            /* [retval][out] */ CellFormat **RHS);
        
        virtual /* [helpcontext][propputref][id] */ HRESULT STDMETHODCALLTYPE putref_FindFormat( 
            /* [in] */ CellFormat *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ReplaceFormat( 
            /* [retval][out] */ CellFormat **RHS);
        
        virtual /* [helpcontext][propputref][id] */ HRESULT STDMETHODCALLTYPE putref_ReplaceFormat( 
            /* [in] */ CellFormat *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UsedObjects( 
            /* [retval][out] */ UsedObjects **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CalculationState( 
            /* [retval][out] */ XlCalculationState *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CalculationInterruptKey( 
            /* [retval][out] */ XlCalculationInterruptKey *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CalculationInterruptKey( 
            /* [in] */ XlCalculationInterruptKey RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Watches( 
            /* [retval][out] */ Watches **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayFunctionToolTips( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayFunctionToolTips( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutomationSecurity( 
            /* [retval][out] */ MsoAutomationSecurity *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AutomationSecurity( 
            /* [in] */ MsoAutomationSecurity RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FileDialog( 
            /* [in] */ MsoFileDialogType fileDialogType,
            /* [retval][out] */ FileDialog **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy14( void);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CalculateFullRebuild( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayPasteOptions( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayPasteOptions( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayInsertOptions( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayInsertOptions( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_GenerateGetPivotData( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_GenerateGetPivotData( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoRecover( 
            /* [retval][out] */ AutoRecover **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Hwnd( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Hinstance( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CheckAbort( 
            /* [optional][in] */ VARIANT KeepAbort);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ErrorCheckingOptions( 
            /* [retval][out] */ ErrorCheckingOptions **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoFormatAsYouTypeReplaceHyperlinks( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AutoFormatAsYouTypeReplaceHyperlinks( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SmartTagRecognizers( 
            /* [retval][out] */ SmartTagRecognizers **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_NewWorkbook( 
            /* [retval][out] */ NewFile **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SpellingOptions( 
            /* [retval][out] */ SpellingOptions **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Speech( 
            /* [retval][out] */ Speech **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MapPaperSize( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_MapPaperSize( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowStartupDialog( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowStartupDialog( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DecimalSeparator( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DecimalSeparator( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ThousandsSeparator( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ThousandsSeparator( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UseSystemSeparators( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_UseSystemSeparators( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ThisCell( 
            /* [retval][out] */ Range **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_RTD( 
            /* [retval][out] */ RTD **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayDocumentActionTaskPane( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayDocumentActionTaskPane( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DisplayXMLSourcePane( 
            /* [optional][in] */ VARIANT XmlMap);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ArbitraryXMLSupportAvailable( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Support( 
            /* [in] */ IDispatch *Object,
            /* [in] */ long ID,
            /* [optional][in] */ VARIANT arg,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy20( 
            /* [in] */ long grfCompareFunctions,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MeasurementUnit( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_MeasurementUnit( 
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowSelectionFloaties( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowSelectionFloaties( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowMenuFloaties( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowMenuFloaties( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowDevTools( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowDevTools( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableLivePreview( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableLivePreview( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayDocumentInformationPanel( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayDocumentInformationPanel( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AlwaysUseClearType( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AlwaysUseClearType( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WarnOnFunctionNameConflict( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_WarnOnFunctionNameConflict( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FormulaBarHeight( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_FormulaBarHeight( 
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayFormulaAutoComplete( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayFormulaAutoComplete( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_GenerateTableRefs( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlGenerateTableRefs *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_GenerateTableRefs( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlGenerateTableRefs RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Assistance( 
            /* [retval][out] */ IAssistance **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CalculateUntilAsyncQueriesDone( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableLargeOperationAlert( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableLargeOperationAlert( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_LargeOperationCellThousandCount( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_LargeOperationCellThousandCount( 
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DeferAsyncQueries( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DeferAsyncQueries( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MultiThreadedCalculation( 
            /* [retval][out] */ MultiThreadedCalculation **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SharePointVersion( 
            /* [in] */ BSTR bstrUrl,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ActiveEncryptionSession( 
            /* [retval][out] */ long *RHS);
        
        virtual HRESULT STDMETHODCALLTYPE get_HighQualityModeForGraphics( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual HRESULT STDMETHODCALLTYPE put_HighQualityModeForGraphics( 
            /* [in] */ VARIANT_BOOL RHS);
            
                   
        Application()
        { 
            CREATE_OBJECT; 
            m_cRef = 1;
            m_pITypeInfo = NULL;
            
            HRESULT hr = Init();
            
            if ( FAILED(hr) )
            {
                 ERR( " \n " );
            }
            
            InterlockedIncrement(&g_cComponents);
        }
        virtual ~Application() 
        { 
            InterlockedDecrement(&g_cComponents);    
                
            DELETE_OBJECT; 
        } 
       
        HRESULT Init();
        
        OOServiceManager m_oo_service_manager; 
        OODesktop        m_oo_desktop;
             
private:
        
       long m_cRef; 
       
       ITypeInfo* m_pITypeInfo;
       
       VARIANT_BOOL m_b_screenupdating;
       VARIANT_BOOL m_b_displayalerts;
       VARIANT_BOOL m_b_visible;
       long         m_l_sheetsinnewworkbook; 
             
       CWorkbooks       m_workbooks;            
};

#endif //__UNIOFFICE_EXCEL_APPLICATION_H__
