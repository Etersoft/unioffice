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
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<_IFont*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<_IFont*>(this));
    }     
    
    if ( iid == IID__IFont) {
        TRACE("IRange\n");
        *ppv = static_cast<_IFont*>(this);
    } 
    
    if ( iid == DIID_Font) {
        TRACE("Range \n");
        *ppv = static_cast<Font*>(this);
    }   
      
    if ( *ppv != NULL ) 
    {
        reinterpret_cast<IUnknown*>(*ppv)->AddRef();
         
        return S_OK;
    } else
    {    
        WCHAR str_clsid[39];
         
        StringFromGUID2( iid, str_clsid, 39);
        WTRACE(L"(%s) not supported \n", str_clsid);
        
        return E_NOINTERFACE;                          
    }   		
}
        
ULONG STDMETHODCALLTYPE CFont::AddRef( )
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);   		
}
        
ULONG STDMETHODCALLTYPE CFont::Release( )
{
      TRACE( " ref = %i \n", m_cRef );
      
      if (InterlockedDecrement(&m_cRef) == 0)
      {
              delete this;
              return 0;
      }
      
      return m_cRef;  		
}
        
       
       // IDispatch    
HRESULT STDMETHODCALLTYPE CFont::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;  		
}
        
HRESULT STDMETHODCALLTYPE CFont::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    *ppTInfo = NULL;
    
    if(iTInfo != 0)
    {
        return DISP_E_BADINDEX;
    }
    
    m_pITypeInfo->AddRef();
    *ppTInfo = m_pITypeInfo;
    
    return S_OK;    		
}
        
HRESULT STDMETHODCALLTYPE CFont::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    if (riid != IID_NULL )
    {
        return DISP_E_UNKNOWNINTERFACE;
    }
    
    HRESULT hr = m_pITypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
    
    if ( FAILED(hr) )
    {
     ERR( " name = %s \n", *rgszNames );     
    }
    
    return ( hr ); 		
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
    if ( riid != IID_NULL)
    {
        return DISP_E_UNKNOWNINTERFACE;
    }
    
    HRESULT hr = m_pITypeInfo->Invoke(
                 static_cast<IDispatch*>(static_cast<_IFont*>(this)), 
                 dispIdMember, 
                 wFlags, 
                 pDispParams, 
                 pVarResult, 
                 pExcepInfo, 
                 puArgErr);       
            
    if ( FAILED(hr) )
    { 
        ERR( " dispIdMember = %i   hr = %08x \n", dispIdMember, hr ); 
	    ERR( " wFlags = %i  \n", wFlags );   
	    ERR( " pDispParams->cArgs = %i \n", pDispParams->cArgs );
    }  
	             
    return ( hr ); 		
}
         
               
        // IRange     
HRESULT STDMETHODCALLTYPE CFont::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
   TRACE_IN;             
   
   if ( m_p_application == NULL )
   {
       ERR( " m_p_application == NULL \n " ); 
       TRACE_OUT;
       return ( S_FALSE );    
   }
            
   HRESULT hr = S_OK;
   
   hr = (static_cast<Application*>( m_p_application ))->get_Application( RHS );          
             
   TRACE_OUT;
   return hr;    		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;  		
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
        
HRESULT STDMETHODCALLTYPE CFont::get_Name( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    BSTR font_name = SysAllocString( L"" );
    
    hr = m_oo_font.getCharFontName( font_name );
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_font.getCharFontName \n" );
		TRACE_OUT;
		return ( hr );   	 
  	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_BSTR;
    V_BSTR( RHS ) = SysAllocString( font_name );
    
    SysFreeString( font_name );
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Name( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    hr = m_oo_font.setCharFontName( V_BSTR( &RHS ) );
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_font.setCharFontName \n" );
		TRACE_OUT;
		return ( hr );   	 
  	}
    
    TRACE_OUT;
    return ( hr );  		
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
            
