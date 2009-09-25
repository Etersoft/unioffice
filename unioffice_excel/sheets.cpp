/*
 * implementation of Sheets
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

#include "sheets.h"

#include "application.h"
#include "workbook.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CSheets::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<Sheets*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(this);
    }     
    
    if ( iid == IID_Sheets ) {
        TRACE("Sheets \n");
        *ppv = static_cast<Sheets*>(this);
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

ULONG STDMETHODCALLTYPE CSheets::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);              
}

ULONG STDMETHODCALLTYPE CSheets::Release()
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
HRESULT STDMETHODCALLTYPE CSheets::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;          
}

HRESULT STDMETHODCALLTYPE CSheets::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CSheets::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE CSheets::Invoke(
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
               
        // Sheets
HRESULT STDMETHODCALLTYPE CSheets::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE CSheets::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Parent( 
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
   
   hr = (static_cast<Workbook*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;          
}
        
HRESULT STDMETHODCALLTYPE CSheets::Add( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [optional][in] */ VARIANT Count,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Copy( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Count( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Delete( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::FillAcrossSheets( 
            /* [in] */ Range	*Range,
            /* [defaultvalue][optional][in] */ XlFillWith Type,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Item( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;
   
   HRESULT hr = get__Default( Index, RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " call get__Default \n" );     
   }
   
   TRACE_OUT;
   return ( hr );            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Move( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::__PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Select( 
            /* [optional][in] */ VARIANT Replace,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_HPageBreaks( 
            /* [retval][out] */ HPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_VPageBreaks( 
            /* [retval][out] */ vPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get__Default( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::_PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [optional][in] */ VARIANT IgnorePrintAreas,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
            
HRESULT CSheets::Next ( ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched)
{
    TRACE_IN;    
        
    HRESULT hr;
    ULONG l;
    long l1;
    long count = 0;
    ULONG l2;
    IDispatch *dret;
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
      V_I4( &varindex ) = l1 + 1;    //Because index of sheets start from 1
      
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
        
HRESULT CSheets::Skip ( ULONG celt)
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

HRESULT CSheets::Reset( )
{
   TRACE_IN;
   
   enum_position = 0;
   
   TRACE_OUT;
   return S_OK;       
}

HRESULT CSheets::Clone(IEnumVARIANT** ppEnum)
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





HRESULT CSheets::Init()
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_Sheets, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;        
}

HRESULT CSheets::Put_Application( void* p_application )
{
    m_p_application = p_application;
        
    return S_OK;      
}

HRESULT CSheets::Put_Parent( void* p_parent )
{
   m_p_parent = p_parent;
   
   return S_OK;     
}








