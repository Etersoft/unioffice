/*
 * implementation of Name
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

#include "name.h"

#include "application.h"
#include "names.h"


       // IUnknown
HRESULT STDMETHODCALLTYPE CName::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<IName*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<IName*>(this));
    }     
    
    if ( iid == IID_IName) {
        TRACE("IName\n");
        *ppv = static_cast<IName*>(this);
    } 
    
    if ( iid == DIID_Name ) {
        TRACE("Name \n");
        *ppv = static_cast<Name*>(this);
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

ULONG STDMETHODCALLTYPE CName::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);		
}

ULONG STDMETHODCALLTYPE CName::Release()
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
HRESULT STDMETHODCALLTYPE CName::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;		
}

HRESULT STDMETHODCALLTYPE CName::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CName::GetIDsOfNames(
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
    
    return hr; 		
}

HRESULT STDMETHODCALLTYPE CName::Invoke(
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
                 static_cast<IDispatch*>(static_cast<IName*>(this)), 
                 dispIdMember, 
                 wFlags, 
                 pDispParams, 
                 pVarResult, 
                 pExcepInfo, 
                 puArgErr);
      
    if ( FAILED(hr) )
    {
     ERR( " dispIdMember = %i \n", dispIdMember );     
    }  
                 
    return hr;		
}
		
		
        //Names
HRESULT STDMETHODCALLTYPE CName::get_Application( 
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
   
   _Application* p_application = NULL;
   
   hr = (static_cast<IUnknown*>( m_p_application ))->QueryInterface( IID__Application,(void**)(&p_application) ); 
   if ( FAILED( hr ) )
   {
       ERR( " IUnknown->QueryInterface \n" );
	   TRACE_OUT;
	   return ( hr );	  	
   }
   
   hr = p_application->get_Application( RHS );          
   
   if ( p_application != NULL )
   {
       p_application->Release();
	   p_application = NULL;	  	
   }
             
   TRACE_OUT;
   return hr;		
}
        
HRESULT STDMETHODCALLTYPE CName::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return ( S_OK );  		
}
        
HRESULT STDMETHODCALLTYPE CName::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;             
    
   if ( m_p_parent == NULL )
   {
       ERR( " m_p_parent == NULL \n " ); 
       TRACE_OUT;
       return ( S_FALSE );    
   }    
            
   HRESULT hr = S_OK;
   
   hr = (static_cast<IUnknown*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get__Default( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Index( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Category( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Category( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_CategoryLocal( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_CategoryLocal( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CName::Delete( void)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_MacroType( 
            /* [retval][out] */ XlXLMMacroType *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_MacroType( 
            /* [in] */ XlXLMMacroType RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CName::get_Name( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_IN;
   HRESULT hr;
   BSTR result;
   
   result = SysAllocString( m_oo_named_range.getName( ) );

   if ( lstrlenW( result ) == 0 )
   {
       ERR( " m_oo_named_range.getName \n" );  
       hr = E_FAIL;   
   } else
   {
      *RHS = SysAllocString( result );       
   }

   SysFreeString( result );

   TRACE_OUT;
   return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CName::put_Name( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ BSTR RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   hr = m_oo_named_range.setName( SysAllocString( RHS ) );
   if ( FAILED( hr ) )
   {
       ERR( " m_oo_named_range.setName \n" );     
   }
   
   TRACE_OUT;
   return ( hr );		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersTo( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersTo( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_ShortcutKey( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_ShortcutKey( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Value( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Value( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Visible( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Visible( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_NameLocal( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_NameLocal( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToLocal( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersToLocal( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToR1C1( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersToR1C1( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToR1C1Local( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersToR1C1Local( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToRange( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Comment( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Comment( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_WorkbookParameter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_WorkbookParameter( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_ValidWorkbookParameter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
      
HRESULT CName::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_IName, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr; 		
}
       
HRESULT CName::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;
}

HRESULT CName::Put_Parent( void* p_parent)
{
    TRACE_IN;
    
    m_p_parent = p_parent;
    
    TRACE_OUT;    
    return S_OK;
}

HRESULT CName::InitWrapper( OONamedRange oo_named_range )
{
    TRACE_IN;
    
    m_oo_named_range = oo_named_range;
	
	TRACE_OUT; 		
}

