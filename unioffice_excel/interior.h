/*
 * header file - Interior
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

#ifndef __UNIOFFICE_EXCEL_INTERIOR_H__
#define __UNIOFFICE_EXCEL_INTERIOR_H__

#include "unioffice_excel_private.h"


class CInterior : public IInterior, public Interior
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

			   // IInterior               
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Color( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Color( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ColorIndex( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_InvertIfNegative( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_InvertIfNegative( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Pattern( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Pattern( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PatternColor( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PatternColor( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PatternColorIndex( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PatternColorIndex( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ThemeColor( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_TintAndShade( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PatternThemeColor( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PatternThemeColor( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PatternTintAndShade( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PatternTintAndShade( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Gradient( 
            /* [retval][out] */ IDispatch **RHS);
       
       CInterior()
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

       virtual ~CInterior()
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

#endif // __UNIOFFICE_EXCEL_INTERIOR_H__
	   			 
