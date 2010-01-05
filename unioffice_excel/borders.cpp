/*
 * implementation of Borders
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

#include "borders.h"
#include "application.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CBorders::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<IBorders*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<IBorders*>(this));
    }     
    
    if ( iid == IID_IBorders) {
        TRACE("IRange\n");
        *ppv = static_cast<IBorders*>(this);
    } 
    
    if ( iid == DIID_Borders) {
        TRACE("Range \n");
        *ppv = static_cast<Borders*>(this);
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

ULONG STDMETHODCALLTYPE CBorders::AddRef()       
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef); 		
} 

ULONG STDMETHODCALLTYPE CBorders::Release()
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
HRESULT STDMETHODCALLTYPE CBorders::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;		
} 

HRESULT STDMETHODCALLTYPE CBorders::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CBorders::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE CBorders::Invoke(
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
                 static_cast<IDispatch*>(static_cast<IBorders*>(this)), 
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

  		   // IBorders              
HRESULT STDMETHODCALLTYPE CBorders::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE CBorders::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK; 		
} 
        
HRESULT STDMETHODCALLTYPE CBorders::get_Parent( 
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
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Color( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_Color( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_ColorIndex( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
HRESULT STDMETHODCALLTYPE CBorders::get_Count( 
            /* [retval][out] */ long *RHS)
{
    TRACE_IN;
    
    *RHS = 12;
    
    TRACE_OUT;
    return ( S_OK );		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Item( 
            /* [in] */ XlBordersIndex Index,
            /* [retval][out] */ Border	**RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    hr = get__Default( Index, RHS );
    if ( FAILED( hr ) )
    {
	    ERR( " get_Items -> get__Default " );   	 
    }
    
    TRACE_OUT;
	return ( hr ); 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_LineStyle( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_LineStyle( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Value( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_Value( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Weight( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_Weight( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get__Default( 
            /* [in] */ XlBordersIndex Index,
            /* [retval][out] */ Border	**RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_ThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_TintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
            
HRESULT CBorders::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_IBorders, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;		
} 
       
HRESULT CBorders::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK; 		
} 

HRESULT CBorders::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;       
      
   TRACE_OUT;
   return S_OK;  		
}            
     
HRESULT CBorders::InitWrapper( OOBorders _oo_borders )
{
    m_oo_borders = _oo_borders;     
}

	        
