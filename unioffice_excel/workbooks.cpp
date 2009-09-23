/*
 * implementation of Workbooks
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

#include "workbooks.h"
#include "application.h"

#include <algorithm>

using namespace std;

// IUnknown
HRESULT STDMETHODCALLTYPE CWorkbooks::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<Workbooks*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(this);
    }     
    
    if ( iid == IID_Workbooks ) {
        TRACE("Workbooks \n");
        *ppv = static_cast<Workbooks*>(this);
    } 
    
    if ( iid == IID_IEnumVARIANT ) {
        TRACE(" IEnumVARIANT \n");
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

ULONG STDMETHODCALLTYPE CWorkbooks::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);                     
}

ULONG STDMETHODCALLTYPE CWorkbooks::Release()
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
HRESULT STDMETHODCALLTYPE CWorkbooks::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;                    
}

HRESULT STDMETHODCALLTYPE CWorkbooks::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CWorkbooks::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE CWorkbooks::Invoke(
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
                 static_cast<IDispatch*>(this), 
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
       
//Workbooks
HRESULT STDMETHODCALLTYPE CWorkbooks::get_Application( 
             Application	**RHS)
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
        
HRESULT STDMETHODCALLTYPE CWorkbooks::get_Creator( 
             XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;                     
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::get_Parent( 
             IDispatch **RHS)
{
   TRACE_IN;             
    
   if ( m_p_parent == NULL )
   {
       ERR( " m_p_parent == NULL \n " ); 
       TRACE_OUT;
       return ( S_FALSE );    
   }    
            
   HRESULT hr = S_OK;
   
   hr = (static_cast<Application*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;                        
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::Add( 
             VARIANT Template,
             long lcid,
             Workbook **RHS)
{
    HRESULT hr;
    
    TRACE_IN;
             
    Workbook* pWorkbook = new Workbook;    
    
    pWorkbook->Put_Application( m_p_application );
    pWorkbook->Put_Parent( (void*)this );
           
    if (pWorkbook == NULL)
   {
       return E_OUTOFMEMORY;
   }         
      
   hr = pWorkbook->QueryInterface( CLSID_Workbook, ( void** ) RHS );  
   if ( FAILED( hr ) )
   {
       ERR( " QueryInterface ( CLSID_Workbook ) \n" );     
   }
             
   pWorkbook->AddRef( );             
   m_lst_of_workbook.push_back( pWorkbook ); 
   m_it_of_workbook = --m_lst_of_workbook.end();    
       
     
   // Init  
   if ((Is_Variant_Null( Template )) || (lstrlenW(V_BSTR(&Template)) == 0))
   {
       hr = pWorkbook->NewDocument(  );  
   } else
   {
       hr = pWorkbook->NewDocumentAsTemplate( V_BSTR(&Template) );        
   }  
   
   if ( FAILED( hr ) )
   {
       ERR( " create new document \n" );     
   }
     
   pWorkbook->Release();   
                     
   TRACE_OUT;
   return hr;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::Close( 
             long lcid)
{         
   TRACE_IN;
   
   HRESULT hr = S_OK;
   
   VARIANT SaveChanges;
   VARIANT Filename;
   VARIANT RouteWorkbook;
   
   VariantInit( &SaveChanges   );
   VariantInit( &Filename      );
   VariantInit( &RouteWorkbook );   
   
   V_VT( &SaveChanges   ) = VT_BOOL;
   V_BOOL( &SaveChanges   ) = VARIANT_FALSE;
   
   V_VT( &Filename      ) = VT_BSTR;
   V_BSTR( &Filename      ) = SysAllocString( L"" );
   
   V_VT( &RouteWorkbook ) = VT_BOOL;   
   V_BOOL( &RouteWorkbook ) = VARIANT_FALSE;   
   
   int iter = 0;
   
   while ( m_lst_of_workbook.size() > 0 )
   {
       (*(m_lst_of_workbook.begin()))->Close( SaveChanges, Filename, RouteWorkbook, lcid );
       
       iter++;
       
       if ( iter > 1000 )
           break;
   }
   
   if ( iter > 1000 )
   {
       ERR( " iter > 1000 \n" );
       hr = E_FAIL;     
   }
   
   VariantClear( &SaveChanges   );
   VariantClear( &Filename      );
   VariantClear( &RouteWorkbook );
   
   TRACE_OUT;
   return hr;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::get_Count( 
             long *RHS)
{
   *RHS = m_lst_of_workbook.size();          
             
   return S_OK;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::get_Item( 
             VARIANT Index,
             Workbook **RHS)
{
   TRACE_IN;          
             
   HRESULT hr = S_OK;
   
   hr = get__Default( Index, RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " get_Item \n" );     
   }          
             
   TRACE_OUT;
   return ( hr );                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::get__NewEnum( 
             IUnknown **RHS)
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
        
HRESULT STDMETHODCALLTYPE CWorkbooks::_Open( 
             BSTR Filename,
             VARIANT UpdateLinks,
             VARIANT ReadOnly,
             VARIANT Format,
             VARIANT Password,
             VARIANT WriteResPassword,
             VARIANT IgnoreReadOnlyRecommended,
             VARIANT Origin,
             VARIANT Delimiter,
             VARIANT Editable,
             VARIANT Notify,
             VARIANT Converter,
             VARIANT AddToMru,
             long lcid,
             Workbook **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::__OpenText( 
             BSTR Filename,
             VARIANT Origin,
             VARIANT StartRow,
             VARIANT DataType,
             XlTextQualifier TextQualifier,
             VARIANT ConsecutiveDelimiter,
             VARIANT Tab,
             VARIANT Semicolon,
             VARIANT Comma,
             VARIANT Space,
             VARIANT Other,
             VARIANT OtherChar,
             VARIANT FieldInfo,
             VARIANT TextVisualLayout,
             long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::get__Default( 
             VARIANT Index,
             Workbook **RHS)
{
   TRACE_IN;
   
   HRESULT hr = S_OK;
   VARIANT i4_index;
   
   CorrectArg( Index, &Index);
   
   VariantInit(&i4_index);
   
   if ( V_VT( &Index ) == VT_BSTR )
   {
       ERR(" BSTR parameters not supported \n");                
       hr = E_FAIL;
   } else {
        hr = VariantChangeTypeEx( &i4_index, &Index, 0, 0, VT_I4 );
        if ( FAILED( hr ) ) {
            ERR(" when VariantChangeTypeEx \n");
            
            TRACE_OUT;
            return E_FAIL;
        }       
           
        if ( (V_I4( &i4_index ) - 1 ) >= m_lst_of_workbook.size() ) {
            ERR(" Index i4_index = %i,  count = %i \n", V_I4(&i4_index) - 1, m_lst_of_workbook.size() );  
            
            TRACE_OUT;
            return E_FAIL;               
        } 
        
        std::list< Workbook* >::iterator it_workbook = m_lst_of_workbook.begin();
        long index = 0;
        while ( it_workbook != m_lst_of_workbook.end() )
        {
           if ( (V_I4(&i4_index) - 1) == index )
           {
                hr = (*it_workbook)->QueryInterface( CLSID_Workbook, (void**) RHS );
                
                if ( FAILED( hr ) )
                {
                    ERR( " QueryInterface \n" );
                }
                
                TRACE_OUT;
                return ( hr );
           }
           
           index++;   
           it_workbook++;   
        }
        
        ERR( " not find in array element i4_ndex = %i    index = %i \n", V_I4(&i4_index) - 1, index );
        
        TRACE_OUT;
        return ( E_FAIL );
    }
    
    ERR(" parameters \n");           
             
    TRACE_OUT;
    return ( hr );                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::_OpenText( 
             BSTR Filename,
             VARIANT Origin,
             VARIANT StartRow,
             VARIANT DataType,
             XlTextQualifier TextQualifier,
             VARIANT ConsecutiveDelimiter,
             VARIANT Tab,
             VARIANT Semicolon,
             VARIANT Comma,
             VARIANT Space,
             VARIANT Other,
             VARIANT OtherChar,
             VARIANT FieldInfo,
             VARIANT TextVisualLayout,
             VARIANT DecimalSeparator,
             VARIANT ThousandsSeparator,
             long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::Open( 
             BSTR Filename,
             VARIANT UpdateLinks,
             VARIANT ReadOnly,
             VARIANT Format,
             VARIANT Password,
             VARIANT WriteResPassword,
             VARIANT IgnoreReadOnlyRecommended,
             VARIANT Origin,
             VARIANT Delimiter,
             VARIANT Editable,
             VARIANT Notify,
             VARIANT Converter,
             VARIANT AddToMru,
             VARIANT Local,
             VARIANT CorruptLoad,
             long lcid,
             Workbook **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::OpenText( 
             BSTR Filename,
             VARIANT Origin,
             VARIANT StartRow,
             VARIANT DataType,
             XlTextQualifier TextQualifier,
             VARIANT ConsecutiveDelimiter,
             VARIANT Tab,
             VARIANT Semicolon,
             VARIANT Comma,
             VARIANT Space,
             VARIANT Other,
             VARIANT OtherChar,
             VARIANT FieldInfo,
             VARIANT TextVisualLayout,
             VARIANT DecimalSeparator,
             VARIANT ThousandsSeparator,
             VARIANT TrailingMinusNumbers,
             VARIANT Local,
             long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::OpenDatabase( 
             BSTR Filename,
             VARIANT CommandText,
             VARIANT CommandType,
             VARIANT BackgroundQuery,
             VARIANT ImportDataAs,
             Workbook **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::CheckOut( 
             BSTR Filename)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::CanCheckOut( 
             BSTR Filename,
             VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::_OpenXML( 
             BSTR Filename,
             VARIANT Stylesheets,
             Workbook **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::OpenXML( 
             BSTR Filename,
             VARIANT Stylesheets,
             VARIANT LoadOption,
             Workbook **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                      
}

HRESULT STDMETHODCALLTYPE CWorkbooks::Next( ULONG celt, VARIANT *rgVar , ULONG *pCeltFetched)
{
    TRACE_IN;    
        
    HRESULT hr;
    ULONG l;
    long l1;
    long count = 0;
    ULONG l2;
    Workbook *dret;
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
      V_I4( &varindex ) = l1 + 1;    //Because index of workbook start from 1
      
      hr = get_Item( varindex, &dret);
            
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

HRESULT STDMETHODCALLTYPE CWorkbooks::Skip( ULONG celt )
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

HRESULT STDMETHODCALLTYPE CWorkbooks::Reset( )
{
   TRACE_IN;
   
   enum_position = 0;
   
   TRACE_OUT;
   return S_OK;          
}
        
HRESULT STDMETHODCALLTYPE CWorkbooks::Clone( IEnumVARIANT** ppEnum)
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


HRESULT CWorkbooks::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_Workbooks, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }
     
     return hr;
}

HRESULT CWorkbooks::Put_Visible( VARIANT_BOOL RHS)
{
   HRESULT hr = S_OK;
   std::list< Workbook* >::iterator it_begin = m_lst_of_workbook.begin();
   
   TRACE_IN;
            
   while ( it_begin != m_lst_of_workbook.end() )
   {
       if ( FAILED( (*it_begin)->Put_Visible( RHS ) ) )
           hr = S_FALSE;      
       
       it_begin++;
   }     
     
   if ( FAILED( hr ) )
   {
       ERR( " \n " );     
   }     
        
   TRACE_OUT;
   return hr;        
}

HRESULT CWorkbooks::Put_Application( void* p_application )
{
    m_p_application = p_application;
        
    return S_OK;      
}

HRESULT CWorkbooks::Put_Parent( void* p_parent )
{
   m_p_parent = p_parent;
   
   return S_OK;     
}

HRESULT CWorkbooks::DeleteWorkbookFromVector( Workbook* _workbook_to_delete )
{
    HRESULT hr = S_OK;
    
    TRACE_IN;
    
    list< Workbook* >::iterator   it_workbook = m_lst_of_workbook.end();
    
    it_workbook = find( m_lst_of_workbook.begin(), m_lst_of_workbook.end(), _workbook_to_delete );
    
    if ( it_workbook != m_lst_of_workbook.end() )
    { // we find an element
    
        TRACE( " element find  \n " );
         
        (*it_workbook)->Release();
            
        m_lst_of_workbook.erase( it_workbook );
    
    } else
    { // element not found
        ERR( " element was not find \n" );  
    }
    
    TRACE_OUT;
    return ( hr );       
}
