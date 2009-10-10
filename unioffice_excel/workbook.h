/*
 * header file - Workbook
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

#ifndef __UNIOFFICE_EXCEL_WORKBOOK_H__
#define __UNIOFFICE_EXCEL_WORKBOOK_H__

#include "unioffice_excel_private.h"
#include "../OOWrappers/oo_document.h"
#include "sheets.h"

class Workbook : public _Workbook
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
       
       
       virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_AcceptLabelsInFormulas( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_AcceptLabelsInFormulas( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Activate( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ActiveChart( 
            /* [retval][out] */ Chart **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ActiveSheet( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Author( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_Author( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoUpdateFrequency( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AutoUpdateFrequency( 
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_AutoUpdateSaveChanges( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_AutoUpdateSaveChanges( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ChangeHistoryDuration( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ChangeHistoryDuration( 
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_BuiltinDocumentProperties( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ChangeFileAccess( 
            /* [in] */ XlFileAccess Mode,
            /* [optional][in] */ VARIANT WritePassword,
            /* [optional][in] */ VARIANT Notify,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ChangeLink( 
            /* [in] */ BSTR Name,
            /* [in] */ BSTR NewName,
            /* [defaultvalue][optional][in] */ XlLinkType Type,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Charts( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Close( 
            /* [optional][in] */ VARIANT SaveChanges,
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT RouteWorkbook,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CodeName( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get__CodeName( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put__CodeName( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Colors( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Colors( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CommandBars( 
            /* [retval][out] */ CommandBars **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Comments( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_Comments( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ConflictResolution( 
            /* [retval][out] */ XlSaveConflictResolution *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ConflictResolution( 
            /* [in] */ XlSaveConflictResolution RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Container( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CreateBackup( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CustomDocumentProperties( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Date1904( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Date1904( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE DeleteNumberFormat( 
            /* [in] */ BSTR NumberFormat,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_DialogSheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayDrawingObjects( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlDisplayDrawingObjects *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayDrawingObjects( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlDisplayDrawingObjects RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ExclusiveAccess( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FileFormat( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlFileFormat *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ForwardMailer( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FullName( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_HasMailer( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_HasMailer( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_HasPassword( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_HasRoutingSlip( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_HasRoutingSlip( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_IsAddin( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_IsAddin( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Keywords( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_Keywords( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE LinkInfo( 
            /* [in] */ BSTR Name,
            /* [in] */ XlLinkInfo LinkInfo,
            /* [optional][in] */ VARIANT Type,
            /* [optional][in] */ VARIANT EditionRef,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE LinkSources( 
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Mailer( 
            /* [retval][out] */ Mailer **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE MergeWorkbook( 
            /* [in] */ VARIANT Filename);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Modules( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_MultiUserEditing( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Names( 
            /* [retval][out] */ Names **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE NewWindow( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Window **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_OnSave( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_OnSave( 
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
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE OpenLinks( 
            /* [in] */ BSTR Name,
            /* [optional][in] */ VARIANT ReadOnly,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Path( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PersonalViewListSettings( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_PersonalViewListSettings( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PersonalViewPrintSettings( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_PersonalViewPrintSettings( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Pivotcaches( 
            /* [retval][out] */ PivotCaches	**RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Post( 
            /* [optional][in] */ VARIANT DestName,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PrecisionAsDisplayed( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_PrecisionAsDisplayed( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
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
            /* [optional][in] */ VARIANT Structure,
            /* [optional][in] */ VARIANT Windows);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _ProtectSharing( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT SharingPassword);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProtectStructure( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ProtectWindows( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ReadOnly( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get__ReadOnlyRecommended( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RefreshAll( void);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Reply( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ReplyAll( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RemoveUser( 
            /* [in] */ long Index);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_RevisionNumber( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Route( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Routed( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_RoutingSlip( 
            /* [retval][out] */ RoutingSlip **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RunAutoMacros( 
            /* [in] */ XlRunAutoMacro Which,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Save( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE _SaveAs( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [defaultvalue][optional][in] */ XlSaveAsAccessMode AccessMode,
            /* [optional][in] */ VARIANT ConflictResolution,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SaveCopyAs( 
            /* [optional][in] */ VARIANT Filename,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Saved( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Saved( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SaveLinkValues( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_SaveLinkValues( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SendMail( 
            /* [in] */ VARIANT Recipients,
            /* [optional][in] */ VARIANT Subject,
            /* [optional][in] */ VARIANT ReturnReceipt,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SendMailer( 
            /* [optional][in] */ VARIANT FileFormat,
            /* [defaultvalue][optional][in] */ XlPriority Priority,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SetLinkOnData( 
            /* [in] */ BSTR Name,
            /* [optional][in] */ VARIANT Procedure,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Sheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowConflictHistory( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowConflictHistory( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Styles( 
            /* [retval][out] */ Styles **RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Subject( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_Subject( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_Title( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_Title( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Unprotect( 
            /* [optional][in] */ VARIANT Password,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE UnprotectSharing( 
            /* [optional][in] */ VARIANT SharingPassword);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE UpdateFromFile( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE UpdateLink( 
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UpdateRemoteReferences( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_UpdateRemoteReferences( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_UserControl( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE put_UserControl( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UserStatus( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CustomViews( 
            /* [retval][out] */ CustomViews **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Windows( 
            /* [retval][out] */ Windows **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Worksheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WriteReserved( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WriteReservedBy( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Excel4IntlMacroSheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Excel4MacroSheets( 
            /* [retval][out] */ Sheets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TemplateRemoveExtData( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_TemplateRemoveExtData( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE HighlightChangesOptions( 
            /* [optional][in] */ VARIANT When,
            /* [optional][in] */ VARIANT Who,
            /* [optional][in] */ VARIANT Where);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_HighlightChangesOnScreen( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_HighlightChangesOnScreen( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_KeepChangeHistory( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_KeepChangeHistory( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ListChangesOnNewSheet( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ListChangesOnNewSheet( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE PurgeChangeHistoryNow( 
            /* [in] */ long Days,
            /* [optional][in] */ VARIANT SharingPassword);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE AcceptAllChanges( 
            /* [optional][in] */ VARIANT When,
            /* [optional][in] */ VARIANT Who,
            /* [optional][in] */ VARIANT Where);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RejectAllChanges( 
            /* [optional][in] */ VARIANT When,
            /* [optional][in] */ VARIANT Who,
            /* [optional][in] */ VARIANT Where);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE PivotTableWizard( 
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
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ResetColors( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_VBProject( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FollowHyperlink( 
            /* [in] */ BSTR Address,
            /* [optional][in] */ VARIANT SubAddress,
            /* [optional][in] */ VARIANT NewWindow,
            /* [optional][in] */ VARIANT AddHistory,
            /* [optional][in] */ VARIANT ExtraInfo,
            /* [optional][in] */ VARIANT Method,
            /* [optional][in] */ VARIANT HeaderInfo);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE AddToFavorites( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_IsInplace( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
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
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE WebPagePreview( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PublishObjects( 
            /* [retval][out] */ PublishObjects **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WebOptions( 
            /* [retval][out] */ WebOptions **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ReloadAs( 
            /* [in] */ MsoEncoding Encoding);
        
        virtual /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE get_HTMLProject( 
            /* [retval][out] */ HTMLProject **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnvelopeVisible( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnvelopeVisible( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CalculationVersion( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy17( 
            /* [in] */ long calcid);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE sblt( 
            /* [in] */ BSTR s);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_VBASigned( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowPivotTableFieldList( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowPivotTableFieldList( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_UpdateLinks( 
            /* [retval][out] */ XlUpdateLinks *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_UpdateLinks( 
            /* [in] */ XlUpdateLinks RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE BreakLink( 
            /* [in] */ BSTR Name,
            /* [in] */ XlLinkType Type);
        
        virtual /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Dummy16( void);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SaveAs( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [defaultvalue][optional][in] */ XlSaveAsAccessMode AccessMode,
            /* [optional][in] */ VARIANT ConflictResolution,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [optional][in] */ VARIANT Local,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EnableAutoRecover( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EnableAutoRecover( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_RemovePersonalInformation( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_RemovePersonalInformation( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_FullNameURLEncoded( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CheckIn( 
            /* [optional][in] */ VARIANT SaveChanges,
            /* [optional][in] */ VARIANT Comments,
            /* [optional][in] */ VARIANT MakePublic);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CanCheckIn( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SendForReview( 
            /* [optional][in] */ VARIANT Recipients,
            /* [optional][in] */ VARIANT Subject,
            /* [optional][in] */ VARIANT ShowMessage,
            /* [optional][in] */ VARIANT IncludeAttachment);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ReplyWithChanges( 
            /* [optional][in] */ VARIANT ShowMessage);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE EndReview( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Password( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Password( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_WritePassword( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_WritePassword( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PasswordEncryptionProvider( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PasswordEncryptionAlgorithm( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PasswordEncryptionKeyLength( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SetPasswordEncryptionOptions( 
            /* [optional][in] */ VARIANT PasswordEncryptionProvider,
            /* [optional][in] */ VARIANT PasswordEncryptionAlgorithm,
            /* [optional][in] */ VARIANT PasswordEncryptionKeyLength,
            /* [optional][in] */ VARIANT PasswordEncryptionFileProperties);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_PasswordEncryptionFileProperties( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ReadOnlyRecommended( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ReadOnlyRecommended( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Protect( 
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT Structure,
            /* [optional][in] */ VARIANT Windows);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SmartTagOptions( 
            /* [retval][out] */ SmartTagOptions **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RecheckSmartTags( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Permission( 
            /* [retval][out] */ Permission **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SharedWorkspace( 
            /* [retval][out] */ SharedWorkspace **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Sync( 
            /* [retval][out] */ Sync **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SendFaxOverInternet( 
            /* [optional][in] */ VARIANT Recipients,
            /* [optional][in] */ VARIANT Subject,
            /* [optional][in] */ VARIANT ShowMessage);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_XmlNamespaces( 
            /* [retval][out] */ XmlNamespaces **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_XmlMaps( 
            /* [retval][out] */ XmlMaps **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE XmlImport( 
            /* [in] */ BSTR Url,
            /* [out] */ XmlMap **ImportMap,
            /* [optional][in] */ VARIANT Overwrite,
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ XlXmlImportResult *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_SmartDocument( 
            /* [retval][out] */ SmartDocument **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DocumentLibraryVersions( 
            /* [retval][out] */ DocumentLibraryVersions **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_InactiveListBorderVisible( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_InactiveListBorderVisible( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DisplayInkComments( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DisplayInkComments( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE XmlImportXml( 
            /* [in] */ BSTR Data,
            /* [out] */ XmlMap **ImportMap,
            /* [optional][in] */ VARIANT Overwrite,
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ XlXmlImportResult *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE SaveAsXMLData( 
            /* [in] */ BSTR Filename,
            /* [in] */ XmlMap *Map);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ToggleFormsDesign( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ContentTypeProperties( 
            /* [retval][out] */ MetaProperties **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Connections( 
            /* [retval][out] */ Connections **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE RemoveDocumentInformation( 
            /* [in] */ XlRemoveDocInfoType RemoveDocInfoType);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Signatures( 
            /* [retval][out] */ SignatureSet **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE CheckInWithVersion( 
            /* [optional][in] */ VARIANT SaveChanges,
            /* [optional][in] */ VARIANT Comments,
            /* [optional][in] */ VARIANT MakePublic,
            /* [optional][in] */ VARIANT VersionType);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ServerPolicy( 
            /* [retval][out] */ ServerPolicy **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE LockServerFile( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DocumentInspectors( 
            /* [retval][out] */ DocumentInspectors **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE GetWorkflowTasks( 
            /* [retval][out] */ WorkflowTasks **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE GetWorkflowTemplates( 
            /* [retval][out] */ WorkflowTemplates **RHS);
        
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
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ServerViewableItems( 
            /* [retval][out] */ ServerViewableItems **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_TableStyles( 
            /* [retval][out] */ TableStyles **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DefaultTableStyle( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DefaultTableStyle( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DefaultPivotTableStyle( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DefaultPivotTableStyle( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CheckCompatibility( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_CheckCompatibility( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_HasVBProject( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_CustomXMLParts( 
            /* [retval][out] */ CustomXMLParts **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Final( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Final( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Research( 
            /* [retval][out] */ Research **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Theme( 
            /* [retval][out] */ OfficeTheme **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ApplyTheme( 
            /* [in] */ BSTR Filename);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Excel8CompatibilityMode( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ConnectionsDisabled( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE EnableConnections( void);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ShowPivotChartActiveFields( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ShowPivotChartActiveFields( 
            /* [in] */ VARIANT_BOOL RHS);
        
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
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_IconSets( 
            /* [retval][out] */ IconSets **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_EncryptionProvider( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_EncryptionProvider( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_DoNotPromptForConvert( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_DoNotPromptForConvert( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_ForceFullCalculation( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_ForceFullCalculation( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE ProtectSharing( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT SharingPassword,
            /* [optional][in] */ VARIANT FileFormat); 
       
       Workbook()
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
       
       virtual ~Workbook()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }
       
       HRESULT Init();
       HRESULT Put_Visible( VARIANT_BOOL );
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* );
       
       HRESULT NewDocument( );
       HRESULT NewDocumentAsTemplate( BSTR );  
       
       OODocument&        GetOODocument( );
       
private:               
 
       long m_cRef; 
       
       ITypeInfo* m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;
              
       OODocument   m_oo_document;        
       
       CSheets      m_sheets;
};





#endif //__UNIOFFICE_EXCEL_WORKBOOK_H__
