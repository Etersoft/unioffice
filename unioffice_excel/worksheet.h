/*
 * header file - Worksheet
 *
 * Copyright (C) 2009 Sinitsin Ivan (Etersoft) <ivan@etersoft.ru>
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
 
 
#ifndef __UNIOFFICE_EXCEL_WORKSHEET_H__
#define __UNIOFFICE_EXCEL_WORKSHEET_H__

#include "unioffice_excel_private.h"
#include "../OOWrappers/oo_sheet.h"

class Worksheet: public _Worksheet
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
       
       
       // _Worksheet
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Activate( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Copy( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Delete( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CodeName( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get__CodeName( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put__CodeName( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Index( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Move( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Next( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
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
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PageSetup( 
            /* [retval][out] */ PageSetup	**RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Previous( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE __PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _Protect( 
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT DrawingObjects,
            /* [optional][in] */ VARIANT Contents,
            /* [optional][in] */ VARIANT Scenarios,
            /* [optional][in] */ VARIANT UserInterfaceOnly,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProtectContents( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProtectDrawingObjects( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProtectionMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProtectScenarios( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _SaveAs( 
            /* [in] */ BSTR Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Select( 
            /* [optional][in] */ VARIANT Replace,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Unprotect( 
            /* [optional][in] */ VARIANT Password,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlSheetVisibility *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlSheetVisibility RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Shapes( 
            /* [retval][out] */ Shapes **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TransitionExpEval( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_TransitionExpEval( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Arcs( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoFilterMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AutoFilterMode( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SetBackgroundPicture( 
            /* [in] */ BSTR Filename);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Buttons( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Calculate( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableCalculation( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableCalculation( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Cells( 
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ChartObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE CheckBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CheckSpelling( 
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [optional][in] */ VARIANT AlwaysSuggest,
            /* [optional][in] */ VARIANT SpellLang,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CircularReference( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ClearArrows( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Columns( 
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ConsolidationFunction( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlConsolidationFunction *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ConsolidationOptions( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ConsolidationSources( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayAutomaticPageBreaks( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayAutomaticPageBreaks( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Drawings( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE DrawingObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE DropDowns( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableAutoFilter( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableAutoFilter( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableSelection( 
            /* [retval][out] */ XlEnableSelection *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableSelection( 
            /* [in] */ XlEnableSelection RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableOutlining( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableOutlining( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnablePivotTable( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnablePivotTable( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE _Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FilterMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ResetAllPageBreaks( void);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE GroupBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE GroupObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Labels( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Lines( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE ListBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Names( 
            /* [retval][out] */ Names	**RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE OLEObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
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
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE OptionButtons( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Outline( 
            /* [retval][out] */ Outline	**RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Ovals( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Paste( 
            /* [optional][in] */ VARIANT Destination,
            /* [optional][in] */ VARIANT Link,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _PasteSpecial( 
            /* [optional][in] */ VARIANT Format,
            /* [optional][in] */ VARIANT Link,
            /* [optional][in] */ VARIANT DisplayAsIcon,
            /* [optional][in] */ VARIANT IconFileName,
            /* [optional][in] */ VARIANT IconIndex,
            /* [optional][in] */ VARIANT IconLabel,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Pictures( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE PivotTables( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE PivotTableWizard( 
            /* [optional][in] */ VARIANT SourceType,
            /* [optional][in] */ VARIANT SourceData,
            /* [optional][in] */ VARIANT TableDestination,
            /* [optional][in] */ VARIANT TableName,
            /* [optional][in] */ VARIANT RowGrand,
            /* [optional][in] */ VARIANT ColumnGrand,
            /* [optional][in] */ VARIANT SaveData,
            /* [optional][in] */ VARIANT HasAutoFormat,
            /* [optional][in] */ VARIANT AutoPage,
            /* [optional][in] */ VARIANT Reserved,
            /* [optional][in] */ VARIANT BackgroundQuery,
            /* [optional][in] */ VARIANT OptimizeCache,
            /* [optional][in] */ VARIANT PageFieldOrder,
            /* [optional][in] */ VARIANT PageFieldWrapCount,
            /* [optional][in] */ VARIANT ReadData,
            /* [optional][in] */ VARIANT Connection,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ PivotTable **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Range( 
            /* [in] */ VARIANT Cell1,
            /* [optional][in] */ VARIANT Cell2,
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Rectangles( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Rows( 
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Scenarios( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ScrollArea( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ScrollArea( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE ScrollBars( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ShowAllData( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ShowDataForm( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Spinners( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_StandardHeight( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_StandardWidth( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_StandardWidth( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE TextBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TransitionFormEntry( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_TransitionFormEntry( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Type( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlSheetType *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UsedRange( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_HPageBreaks( 
            /* [retval][out] */ HPageBreaks **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_VPageBreaks( 
            /* [retval][out] */ vPageBreaks **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_QueryTables( 
            /* [retval][out] */ QueryTables **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayPageBreaks( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayPageBreaks( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Comments( 
            /* [retval][out] */ Comments **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Hyperlinks( 
            /* [retval][out] */ HyperLinks **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ClearCircles( void);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CircleInvalid( void);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get__DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put__DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoFilter( 
            /* [retval][out] */ AutoFilter **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Scripts( 
            /* [retval][out] */ Scripts **RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _CheckSpelling( 
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [optional][in] */ VARIANT AlwaysSuggest,
            /* [optional][in] */ VARIANT SpellLang,
            /* [optional][in] */ VARIANT IgnoreFinalYaa,
            /* [optional][in] */ VARIANT SpellScript,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Tab( 
            /* [retval][out] */ Tab **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MailEnvelope( 
            /* [retval][out] */ MsoEnvelope **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SaveAs( 
            /* [in] */ BSTR Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [optional][in] */ VARIANT Local);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CustomProperties( 
            /* [retval][out] */ CustomProperties **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SmartTags( 
            /* [retval][out] */ SmartTags **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Protection( 
            /* [retval][out] */ Protection **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE PasteSpecial( 
            /* [optional][in] */ VARIANT Format,
            /* [optional][in] */ VARIANT Link,
            /* [optional][in] */ VARIANT DisplayAsIcon,
            /* [optional][in] */ VARIANT IconFileName,
            /* [optional][in] */ VARIANT IconIndex,
            /* [optional][in] */ VARIANT IconLabel,
            /* [optional][in] */ VARIANT NoHTMLFormatting,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Protect( 
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT DrawingObjects,
            /* [optional][in] */ VARIANT Contents,
            /* [optional][in] */ VARIANT Scenarios,
            /* [optional][in] */ VARIANT UserInterfaceOnly,
            /* [optional][in] */ VARIANT AllowFormattingCells,
            /* [optional][in] */ VARIANT AllowFormattingColumns,
            /* [optional][in] */ VARIANT AllowFormattingRows,
            /* [optional][in] */ VARIANT AllowInsertingColumns,
            /* [optional][in] */ VARIANT AllowInsertingRows,
            /* [optional][in] */ VARIANT AllowInsertingHyperlinks,
            /* [optional][in] */ VARIANT AllowDeletingColumns,
            /* [optional][in] */ VARIANT AllowDeletingRows,
            /* [optional][in] */ VARIANT AllowSorting,
            /* [optional][in] */ VARIANT AllowFiltering,
            /* [optional][in] */ VARIANT AllowUsingPivotTables);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ListObjects( 
            /* [retval][out] */ ListObjects **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE XmlDataQuery( 
            /* [in] */ BSTR XPath,
            /* [optional][in] */ VARIANT SelectionNamespaces,
            /* [optional][in] */ VARIANT Map,
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE XmlMapQuery( 
            /* [in] */ BSTR XPath,
            /* [optional][in] */ VARIANT SelectionNamespaces,
            /* [optional][in] */ VARIANT Map,
            /* [retval][out] */ Range	**RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [optional][in] */ VARIANT IgnorePrintAreas,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableFormatConditionsCalculation( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableFormatConditionsCalculation( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Sort( 
            /* [retval][out] */ Sort **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ExportAsFixedFormat( 
            /* [in] */ XlFixedFormatType Type,
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT Quality,
            /* [optional][in] */ VARIANT IncludeDocProperties,
            /* [optional][in] */ VARIANT IgnorePrintAreas,
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT OpenAfterPublish,
            /* [optional][in] */ VARIANT FixedFormatExtClassPtr);

        Worksheet()
        { 
            CREATE_OBJECT; 
            m_cRef = 1;
            m_pITypeInfo = NULL;
            
            m_p_application = NULL;
            m_p_parent = NULL;
            
            HRESULT hr = Init();
            
            if ( FAILED(hr) )
            {
                 ERR( " \n " );
            }
            
            InterlockedIncrement(&g_cComponents);
        }
        virtual ~Worksheet() 
        { 
            InterlockedDecrement(&g_cComponents);    
            
            m_p_application = NULL;
            m_p_parent = NULL;            
                
            DELETE_OBJECT; 
        } 

       HRESULT Init();
        
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* );
       
       HRESULT InitWrapper( OOSheet );
       
private:
        
       long m_cRef; 
       
       ITypeInfo* m_pITypeInfo;
      
       void*        m_p_application;
       void*        m_p_parent;  
       
       OOSheet      m_oo_sheet;
           
};

#endif //__UNIOFFICE_EXCEL_WORKSHEET_H__


