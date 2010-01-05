/*
 * header file - Borders
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

#ifndef __UNIOFFICE_EXCEL_BORDERS_H__
#define __UNIOFFICE_EXCEL_BORDERS_H__

#include "unioffice_excel_private.h"
#include "../OOWrappers/oo_borders.h"

class CBorders : public IBorders, public Borders
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

			   // IBorders              
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
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ XlBordersIndex Index,
            /* [retval][out] */ Border	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_LineStyle( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_LineStyle( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Value( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Weight( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Weight( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get__Default( 
            /* [in] */ XlBordersIndex Index,
            /* [retval][out] */ Border	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ThemeColor( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_TintAndShade( 
            /* [in] */ VARIANT RHS);


       CBorders()
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

       virtual ~CBorders()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }

       HRESULT Init( ); 
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* ); 
	   
	   HRESULT InitWrapper( OOBorders );
	   	  
private:
		
       long         m_cRef; 
       
       ITypeInfo*   m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;	
       
       OOBorders    m_oo_borders;
};

#endif // __UNIOFFICE_EXCEL_BORDERS_H__
	   			 
