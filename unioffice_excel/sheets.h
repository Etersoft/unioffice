/*
 * header file - Sheets
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

#ifndef __UNIOFFICE_EXCEL_SHEETS_H__
#define __UNIOFFICE_EXCEL_SHEETS_H__

#include "unioffice_excel_private.h"

class CSheets : public Sheets
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
               
        // Sheets
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [optional][in] */ VARIANT Count,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Copy( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Delete( 
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE FillAcrossSheets( 
            /* [in] */ Range	*Range,
            /* [defaultvalue][optional][in] */ XlFillWith Type,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Move( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS);
        
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
        
        virtual /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Select( 
            /* [optional][in] */ VARIANT Replace,
            /* [lcid][in] */ long lcid);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_HPageBreaks( 
            /* [retval][out] */ HPageBreaks **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_VPageBreaks( 
            /* [retval][out] */ vPageBreaks **RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE get__Default( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS);
        
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
               
       CSheets()
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
               
       virtual ~CSheets()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }               
               
       HRESULT Init();                      
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* );       
       
       

private:
 
       long m_cRef; 
       
       ITypeInfo*   m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;
       
       
       
               
};














#endif //__UNIOFFICE_EXCEL_SHEETS_H__
