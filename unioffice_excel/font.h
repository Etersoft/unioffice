/*
 * header file - Range
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

#ifndef __UNIOFFICE_EXCEL_FONT_H__
#define __UNIOFFICE_EXCEL_FONT_H__

#include "unioffice_excel_private.h"
#include "../OOWrappers/oo_font.h"

class CFont : public IFont, public Font
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

			   // _IFont               
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Background( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Background( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Bold( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Bold( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Color( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Color( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ColorIndex( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FontStyle( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FontStyle( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Italic( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Italic( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE get_OutlineFont( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE put_OutlineFont( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE get_Shadow( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE put_Shadow( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Size( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Size( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Strikethrough( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Strikethrough( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Subscript( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Subscript( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Superscript( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Superscript( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Underline( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Underline( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ThemeColor( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_TintAndShade( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ThemeFont( 
            /* [retval][out] */ XlThemeFont *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ThemeFont( 
            /* [in] */ XlThemeFont RHS);
       
       CFont()
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

       virtual ~CFont()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }

       HRESULT Init( ); 
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* ); 
	   
	   HRESULT InitWrapper( OOFont );	   
	   	   
private:
		
       long         m_cRef; 
       
       ITypeInfo*   m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;	
       
       OOFont      m_oo_font;
};

#endif // __UNIOFFICE_EXCEL_FONT_H__
	   			 
