/*
 * implementation of Outline
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

#include "outline.h"

#include "application.h"
#include "worksheet.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE COutline::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<IOutline*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<IOutline*>(this));
    }     
    
    if ( iid == IID_IOutline) {
        TRACE("IOutline\n");
        *ppv = static_cast<IOutline*>(this);
    } 
    
    if ( iid == DIID_Outline ) {
        TRACE("Outline \n");
        *ppv = static_cast<Outline*>(this);
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

ULONG STDMETHODCALLTYPE COutline::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);   		
}

ULONG STDMETHODCALLTYPE COutline::Release()
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
HRESULT STDMETHODCALLTYPE COutline::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;	
}

HRESULT STDMETHODCALLTYPE COutline::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE COutline::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE COutline::Invoke(
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
                 static_cast<IDispatch*>(static_cast<IOutline*>(this)), 
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
			
			
		// Outline
HRESULT STDMETHODCALLTYPE COutline::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE COutline::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK; 		
}
        
HRESULT STDMETHODCALLTYPE COutline::get_Parent( 
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
   
   hr = (static_cast<Worksheet*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;  		
}
        
HRESULT STDMETHODCALLTYPE COutline::get_AutomaticStyles( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_IN;
    *RHS = VARIANT_FALSE;
    TRACE_OUT;
    return S_OK;  		
}
        
         /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE COutline::put_AutomaticStyles( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
HRESULT STDMETHODCALLTYPE COutline::ShowLevels( 
            /* [optional][in] */ VARIANT RowLevels,
            /* [optional][in] */ VARIANT ColumnLevels,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    VARIANT param1, param2, res;
    long amount = 1;
	long type = toROWS;
	VARIANT param;
	
	VariantInit( &param );

    CorrectArg(RowLevels, &RowLevels);
    CorrectArg(ColumnLevels, &ColumnLevels);
    
    if ( !Is_Variant_Null( RowLevels ) ) {
        hr = VariantChangeTypeEx(&param, &RowLevels, 0, 0, VT_I4);
        if ( FAILED( hr ) ) {
            ERR(" VariantChangeTypeEx   %08x\n ", hr);
            VariantClear( &param );
            TRACE_OUT;
            return ( hr );
        }
        amount = V_I4( &param );
        type = toROWS;
    } else {
        hr = VariantChangeTypeEx(&param, &ColumnLevels, 0, 0, VT_I4);
        if ( FAILED( hr ) ) {
            ERR("VariantChangeTypeEx   %08x\n", hr);
            VariantClear( &param );
            TRACE_OUT;
            return ( hr );
        }
        amount = V_I4( &param );
		type = toCOLUMNS;
    }

    hr = m_oo_sheet.showLevel( amount, type );
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_sheet.showLevel \n" );   	 
    }
	
	VariantClear( &param );	
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE COutline::get_SummaryColumn( 
            /* [retval][out] */ XlSummaryColumn *RHS)
{
    TRACE_IN;
    *RHS = xlSummaryOnLeft;
    TRACE_OUT;
    return S_OK;  		
}
        
HRESULT STDMETHODCALLTYPE COutline::put_SummaryColumn( 
            /* [in] */ XlSummaryColumn RHS)
{
    TRACE_STUB;
    return S_OK;  		
}
        
HRESULT STDMETHODCALLTYPE COutline::get_SummaryRow( 
            /* [retval][out] */ XlSummaryRow *RHS)
{
    TRACE_IN;
    *RHS = xlSummaryAbove;
    TRACE_OUT;
    return S_OK;  		
}
        
HRESULT STDMETHODCALLTYPE COutline::put_SummaryRow( 
            /* [in] */ XlSummaryRow RHS)
{
    TRACE_STUB;
    return S_OK;  		
}
            
            
HRESULT COutline::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_IOutline, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;  		
}
       
HRESULT COutline::Put_Application( void* p_application) 
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK; 
}

HRESULT COutline::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;
   
   TRACE_OUT;
   return S_OK;		
}

HRESULT COutline::InitWrapper( OOSheet oo_sheet )
{
    TRACE_IN;
    
    m_oo_sheet= oo_sheet;
	
	TRACE_OUT; 		
}
