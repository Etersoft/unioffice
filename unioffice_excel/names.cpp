/*
 * implementation of Names
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

#include "names.h"

#include "application.h"
#include "worksheet.h"
#include "name.h"

using namespace com::sun::star::sheet;

       // IUnknown
HRESULT STDMETHODCALLTYPE CNames::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<INames*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<INames*>(this));
    }     
    
    if ( iid == IID_INames) {
        TRACE("INames\n");
        *ppv = static_cast<INames*>(this);
    } 
    
    if ( iid == DIID_Names ) {
        TRACE("Names \n");
        *ppv = static_cast<Names*>(this);
    }
	
    if ( iid == IID_IEnumVARIANT ) {
        TRACE("IEnumVARIANT \n");
        *ppv = static_cast<IEnumVARIANT*>(this);
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

ULONG STDMETHODCALLTYPE CNames::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);	
}

ULONG STDMETHODCALLTYPE CNames::Release()
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
HRESULT STDMETHODCALLTYPE CNames::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;		
}

HRESULT STDMETHODCALLTYPE CNames::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CNames::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE CNames::Invoke(
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
                 static_cast<IDispatch*>(static_cast<INames*>(this)), 
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
HRESULT STDMETHODCALLTYPE CNames::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE CNames::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK; 		
}
        
HRESULT STDMETHODCALLTYPE CNames::get_Parent( 
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
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CNames::Add( 
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT RefersTo,
            /* [optional][in] */ VARIANT Visible,
            /* [optional][in] */ VARIANT MacroType,
            /* [optional][in] */ VARIANT ShortcutKey,
            /* [optional][in] */ VARIANT Category,
            /* [optional][in] */ VARIANT NameLocal,
            /* [optional][in] */ VARIANT RefersToLocal,
            /* [optional][in] */ VARIANT CategoryLocal,
            /* [optional][in] */ VARIANT RefersToR1C1,
            /* [optional][in] */ VARIANT RefersToR1C1Local,
            /* [retval][out] */ Name	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CNames::Item( 
            /* [optional][in] */ VARIANT Index,
            /* [optional][in] */ VARIANT IndexLocal,
            /* [optional][in] */ VARIANT RefersTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Name	**RHS)
{
    TRACE_IN;
    
    HRESULT hr = _Default( Index, IndexLocal, RefersTo, lcid, RHS);
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CNames::_Default( 
            /* [optional][in] */ VARIANT Index,
            /* [optional][in] */ VARIANT IndexLocal,
            /* [optional][in] */ VARIANT RefersTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Name	**RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   CorrectArg(Index, &Index);
   CorrectArg(IndexLocal, &IndexLocal);
   CorrectArg(RefersTo, &RefersTo);
   
   CName* p_name = new CName;
   
   p_name->Put_Application( m_p_application );
   p_name->Put_Parent( this );
   
   OONamedRange    oo_name;
   
   if ( V_VT( &Index ) == VT_BSTR ) 
   {
   	   hr = S_OK;
   	   
  	   oo_name = m_oo_named_ranges.getByName( V_BSTR( &Index ) );
  	   
  	   if ( oo_name.IsNull() )
	   {
	  	  hr = E_FAIL;
          ERR( " m_oo_named_ranges.getByName \n" );
          
          if ( p_name != NULL )
              p_name->Release();
          
          TRACE_OUT;
          return ( hr );
       }
    } else 
	{
       if ( Is_Variant_Null( Index ) ) 
	   {
          ERR(" Empty param \n ");
           
          if ( p_name != NULL )
             p_name->Release();
          
          TRACE_OUT;           
          return E_FAIL;
       } else 
	   {
	   	   hr = S_OK;
			  	  
           oo_name = m_oo_named_ranges.getByIndex( V_I4( &Index ) );
  	   
           if ( oo_name.IsNull() )
	       {
		   	  hr = E_FAIL; 	
              ERR( " m_oo_named_ranges.getByIndex \n" );
          
              if ( p_name != NULL )
                  p_name->Release();
          
              TRACE_OUT;
              return ( hr );
           }
        }
    }   
   
   p_name->InitWrapper( oo_name );
             
   hr = p_name->QueryInterface( DIID_Name, (void**)RHS );
             
   if ( FAILED( hr ) )
   {
       ERR( " p_name.QueryInterface \n" );     
   }
             
   if ( p_name != NULL )
       p_name->Release();
   
   
   TRACE_OUT;
   return E_NOTIMPL; 		
}
       
HRESULT STDMETHODCALLTYPE CNames::get_Count( 
            /* [retval][out] */ long *RHS)
{
    TRACE_IN;
    HRESULT hr;
    long count = -1;
    
    count = m_oo_named_ranges.getCount( );
    
    if ( count < 0 )
    {
	    ERR( " m_oo_named_ranges.getCount \n" );  
		hr = E_FAIL; 	 
    } else
    {
        hr = S_OK;	  	  
    }
    
    *RHS = count;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CNames::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
   TRACE_IN;
   
   HRESULT hr = S_OK;
   
   hr = QueryInterface( IID_IEnumVARIANT, (void**)RHS );
   
   if ( FAILED( hr ) )
   {
        ERR( " FAILED get IID_IEnumVARIANT \n" );    
   }
   
   TRACE_OUT;
   return hr;		
}
			
HRESULT CNames::Init( )
{
     HRESULT hr = S_OK;   
      
     if (m_pITypeInfo == NULL)
     {
        ITypeLib* pITypeLib = NULL;
        hr = LoadRegTypeLib(LIBID_Office,
                                1, 0, // ������ ������
                                0x00,
                                &pITypeLib); 
        
       if (FAILED(hr))
       {
           ERR( " Typelib not register \n" );
           return hr;
       } 
        
       // �������� ���������� ���� ��� ���������� �������
       hr = pITypeLib->GetTypeInfoOfGuid(IID_INames, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;		
}
 
HRESULT CNames::Next ( ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched)
{
    TRACE_IN;    
        
    HRESULT hr;
    ULONG l;
    long l1;
    long count = 0;
    ULONG l2;
    Name *dret;    
    VARIANT varindex, vNull;

    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;

    if ( enum_position < 0 )
    {
        ERR( " enum_position < 0 \n" );
        return ( S_FALSE );
    }
    
    if ( pCeltFetched != NULL )
    {
       *pCeltFetched = 0;
    }
    
    if ( rgVar == NULL )
    {
        ERR( " rgVar == NULL \n" );
        return E_INVALIDARG;
    }

    VariantInit( &varindex );
    
    /*Init Array*/
    for ( l = 0; l < celt; l++)
       VariantInit( &rgVar[l] );

    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
        ERR( " get_Count \n" ); 
        return (E_FAIL);
    }
    
    V_VT( &varindex ) = VT_I4;

    for ( l1 = enum_position, l2 = 0; l1 < count && l2 < celt; l1++, l2++) {
      V_I4( &varindex ) = l1 ; 
      
      hr = Item( varindex, vNull, vNull, 0, &dret);
            
      V_VT( &rgVar[l2] )       = VT_DISPATCH;
      V_DISPATCH( &rgVar[l2] ) = static_cast<IDispatch*>( dret );
      
      if ( FAILED( hr ) )
      {
          ERR( " get_Item \n" );
          goto error;
      }
      
    }

    if (pCeltFetched != NULL)
    {
       *pCeltFetched = l2;
    }
    
    enum_position = l1;
    
    TRACE_OUT;     
    return  ((l2 < celt) ? S_FALSE : S_OK);

error:
      
    for ( l = 0; l < celt; l++)
    {
        VariantClear(&rgVar[l]);
    }
   
    VariantClear( &varindex );
   
    TRACE_OUT;
    return ( hr );		
}

HRESULT CNames::Skip ( ULONG celt)
{
    long count = 0;
    HRESULT hr;
    TRACE_IN;

    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
        ERR( " get_count \n" );     
    }   
    
    enum_position += celt;

    if ( enum_position >= count) 
    {
        enum_position = count - 1;
        TRACE_OUT;
        return S_FALSE;
    }
    
    TRACE_OUT;
    return S_OK; 		
}

HRESULT CNames::Reset( )
{
   TRACE_IN;
   
   enum_position = 0;
   
   TRACE_OUT;
   return S_OK; 		
}

HRESULT CNames::Clone(IEnumVARIANT** ppEnum)
{
   TRACE_IN;
   
   HRESULT hr = S_OK;
   
   hr = QueryInterface( IID_IEnumVARIANT, (void**)ppEnum );
   
   if ( FAILED( hr ) )
   {
        ERR( " FAILED get IID_IEnumVARIANT \n" );    
   }
   
   TRACE_OUT;
   return hr; 		
}
       
HRESULT CNames::Put_Application( void* p_application )
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;
}

HRESULT CNames::Put_Parent( void* p_parent )
{
   TRACE_IN;  
      
   m_p_parent = p_parent;
   
   TRACE_OUT;
   return S_OK;	
}

HRESULT CNames::InitWrapper( OONamedRanges oo_named_ranges )
{
    TRACE_IN;
    
    m_oo_named_ranges = oo_named_ranges;
	
	TRACE_OUT; 		
}		               
