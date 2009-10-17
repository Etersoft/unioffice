/*
 * header file - Outline
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

#ifndef __UNIOFFICE_EXCEL_OUTLINE_H__
#define __UNIOFFICE_EXCEL_OUTLINE_H__

#include "unioffice_excel_private.h"


class COutline : public IOutline, public Outline
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
			
			
		// Outline
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_AutomaticStyles( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_AutomaticStyles( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ShowLevels( 
            /* [optional][in] */ VARIANT RowLevels,
            /* [optional][in] */ VARIANT ColumnLevels,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_SummaryColumn( 
            /* [retval][out] */ XlSummaryColumn *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_SummaryColumn( 
            /* [in] */ XlSummaryColumn RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_SummaryRow( 
            /* [retval][out] */ XlSummaryRow *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_SummaryRow( 
            /* [in] */ XlSummaryRow RHS);		
		

       COutline()
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

       virtual ~COutline()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }

       HRESULT Init( ); 
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* ); 			   
	   
 	  
private:
		
       long         m_cRef; 
       
       ITypeInfo*   m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;			


		  
};

#endif //__UNIOFFICE_EXCEL_OUTLINE_H__
