/*
 * implementation of Font
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

#include "font.h"
#include "application.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CFont::QueryInterface(const IID& iid, void** ppv)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
ULONG STDMETHODCALLTYPE CFont::AddRef( )
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
ULONG STDMETHODCALLTYPE CFont::Release( )
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
       
       // IDispatch    
HRESULT STDMETHODCALLTYPE CFont::GetTypeInfoCount( UINT * pctinfo )
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
HRESULT STDMETHODCALLTYPE CFont::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;    		
}
        
HRESULT STDMETHODCALLTYPE CFont::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;   		
}
        
HRESULT STDMETHODCALLTYPE CFont::Invoke(
               DISPID dispIdMember,
               REFIID riid,
               LCID lcid,
               WORD wFlags,
               DISPPARAMS * pDispParams,
               VARIANT * pVarResult,
               EXCEPINFO * pExcepInfo,
               UINT * puArgErr)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
         
               
        // IRange     
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Background( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Background( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Bold( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Bold( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Color( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Color( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_ColorIndex( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_FontStyle( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_FontStyle( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Italic( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Italic( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Name( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Name( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_OutlineFont( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_OutlineFont( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Shadow( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Shadow( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Size( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Size( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Strikethrough( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Strikethrough( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Subscript( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Subscript( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Superscript( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Superscript( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Underline( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Underline( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_ThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_TintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}

        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_ThemeFont( 
            /* [retval][out] */ XlThemeFont *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_ThemeFont( 
            /* [in] */ XlThemeFont RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
            
HRESULT CFont::Init( )
{
     HRESULT hr = S_OK;   
      
     if (m_pITypeInfo == NULL)
     {
        ITypeLib* pITypeLib = NULL;
        hr = LoadRegTypeLib(LIBID_Office,
                                1, 0, // Номера версии
                                0x00,
                                &pITypeLib); 
        
       if (FAILED(hr))
       {
           ERR( " Typelib not register \n" );
           return hr;
       } 
        
       // Получить информацию типа для интерфейса объекта
       hr = pITypeLib->GetTypeInfoOfGuid(IID__IFont, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;		
}
         
HRESULT CFont::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;  		
}
        
HRESULT CFont::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;       
      
   TRACE_OUT;
   return S_OK; 		
}
        
HRESULT CFont::InitWrapper( OOFont _oo_font )
{
    m_oo_font = _oo_font;     
}            
            
