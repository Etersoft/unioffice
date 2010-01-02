/*
 * implementation of Interior
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

#include "interior.h"
#include "application.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CInterior::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<IInterior*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<IInterior*>(this));
    }     
    
    if ( iid == IID_IInterior) {
        TRACE("IRange\n");
        *ppv = static_cast<IInterior*>(this);
    } 
    
    if ( iid == DIID_Interior) {
        TRACE("Range \n");
        *ppv = static_cast<Interior*>(this);
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

ULONG STDMETHODCALLTYPE CInterior::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef); 	
} 

ULONG STDMETHODCALLTYPE CInterior::Release()
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
HRESULT STDMETHODCALLTYPE CInterior::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK; 		
} 

HRESULT STDMETHODCALLTYPE CInterior::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CInterior::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE CInterior::Invoke(
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
                 static_cast<IDispatch*>(static_cast<IInterior*>(this)), 
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

			   // IInterior               
HRESULT STDMETHODCALLTYPE CInterior::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE CInterior::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK; 		
} 
        
HRESULT STDMETHODCALLTYPE CInterior::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;             
    
   if ( m_p_parent == NULL )
   {
       ERR( " m_p_parent == NULL \n " ); 
       TRACE_OUT;
       return ( E_FAIL );    
   }    
            
   HRESULT hr = S_OK;
   
   hr = (static_cast<IDispatch*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr; 		
} 
        
HRESULT STDMETHODCALLTYPE CInterior::get_Color( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    		
	hr = m_oo_interior.getCellBackColor( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_interior.getCellBackColor \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = value;
    
    TRACE_OUT;
    return ( hr ); 			
} 
        
HRESULT STDMETHODCALLTYPE CInterior::put_Color( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    
    CorrectArg(RHS, &RHS);
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    if ( FAILED( hr ) )
    {
	    ERR( " VariantChangeTypeEx \n" );   
		TRACE_OUT;
		return ( E_FAIL );	 
    }
    
	value = V_I4( &RHS );
		
	hr = m_oo_interior.getCellBackColor( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_interior.getCellBackColor \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr ); 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_ColorIndex( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_InvertIfNegative( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_InvertIfNegative( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Pattern( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_Pattern( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternColorIndex( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_ThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_TintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternTintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternTintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Gradient( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
       
HRESULT CInterior::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_IInterior, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;		
} 
       
HRESULT CInterior::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK; 		
} 

HRESULT CInterior::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;       
      
   TRACE_OUT;
   return S_OK; 		
} 

HRESULT CInterior::InitWrapper( OOInterior _oo_interior )
{
    m_oo_interior = _oo_interior;     
} 
