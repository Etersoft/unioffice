/*
 * implementation of PageSetup
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

#include "page_setup.h"

#include "application.h"
#include "worksheet.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CPageSetup::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<IPageSetup*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<IPageSetup*>(this));
    }     
    
    if ( iid == IID_IPageSetup ) {
        TRACE("IPageSetup \n");
        *ppv = static_cast<IPageSetup*>(this);
    } 
    
    if ( iid == DIID_PageSetup ) {
        TRACE("PageSetup \n");
        *ppv = static_cast<PageSetup*>(this);
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

ULONG STDMETHODCALLTYPE CPageSetup::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef); 		
}

ULONG STDMETHODCALLTYPE CPageSetup::Release()
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
HRESULT STDMETHODCALLTYPE CPageSetup::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK; 	
}

HRESULT STDMETHODCALLTYPE CPageSetup::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CPageSetup::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE CPageSetup::Invoke(
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
                 static_cast<IDispatch*>(static_cast<IPageSetup*>(this)), 
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


HRESULT STDMETHODCALLTYPE CPageSetup::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;  		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_Parent( 
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
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_BlackAndWhite( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_BlackAndWhite( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_BottomMargin( 
            /* [retval][out] */ double *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.BottomMargin( );
    if ( (*RHS) < 0 )
    {
	    ERR( " BottomMargin < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_BottomMargin( 
            /* [in] */ double RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.BottomMargin( RHS );
    if ( FAILED( hr ) )
    {
	    ERR( " BottomMargin \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_CenterFooter( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_CenterFooter( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_CenterHeader( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_CenterHeader( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_CenterHorizontally( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.CenterHorizontally( );
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_CenterHorizontally( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.CenterHorizontally( RHS );
    
    if ( FAILED( hr ) )
    {
	    ERR( " CenterHorizontally \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_CenterVertically( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.CenterVertically( );
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_CenterVertically( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.CenterVertically( RHS );
    
    if ( FAILED( hr ) )
    {
	    ERR( " CenterVertically \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_ChartSize( 
            /* [retval][out] */ XlObjectSize *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_ChartSize( 
            /* [in] */ XlObjectSize RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_Draft( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_Draft( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_FirstPageNumber( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_FirstPageNumber( 
            /* [in] */ long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_FitToPagesTall( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    short result;
    
    result = m_oo_page_style.ScaleToPagesY( );
    if ( result < 0 )
    {
	    ERR( " ScaleToPagesY < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I2;
    V_I2( RHS ) = result;
    
    TRACE_OUT;
    return ( hr );			
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_FitToPagesTall( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    short value;
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I2);
    if ( FAILED( hr ) ) {
        ERR(" VariantChangeType \n");
        return ( hr );
    }
    
    value = V_I2( &RHS );
    
    hr = m_oo_page_style.ScaleToPagesY( value );
    if ( FAILED( hr ) )
    {
	    ERR( " ScaleToPagesY \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_FitToPagesWide( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    short result;
    
    result = m_oo_page_style.ScaleToPagesX( );
    if ( result < 0 )
    {
	    ERR( " ScaleToPagesX < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I2;
    V_I2( RHS ) = result;
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_FitToPagesWide( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    short value;
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I2);
    if ( FAILED( hr ) ) {
        ERR(" VariantChangeType \n");
        return ( hr );
    }
    
    value = V_I2( &RHS );
    
    hr = m_oo_page_style.ScaleToPagesX( value );
    if ( FAILED( hr ) )
    {
	    ERR( " ScaleToPagesX \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );	
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_FooterMargin( 
            /* [retval][out] */ double *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.FooterHeight( );
    if ( (*RHS) < 0 )
    {
	    ERR( " FooterHeight < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_FooterMargin( 
            /* [in] */ double RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.FooterHeight( RHS );
    if ( FAILED( hr ) )
    {
	    ERR( " FooterHeight \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_HeaderMargin( 
            /* [retval][out] */ double *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.HeaderHeight( );
    if ( (*RHS) < 0 )
    {
	    ERR( " HeaderHeight < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_HeaderMargin( 
            /* [in] */ double RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.HeaderHeight( RHS );
    if ( FAILED( hr ) )
    {
	    ERR( " HeaderHeight \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );			
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_LeftFooter( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_LeftFooter( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_LeftHeader( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_LeftHeader( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_LeftMargin( 
            /* [retval][out] */ double *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.LeftMargin( );
    if ( (*RHS) < 0 )
    {
	    ERR( " LeftMargin < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_LeftMargin( 
            /* [in] */ double RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.LeftMargin( RHS );
    if ( FAILED( hr ) )
    {
	    ERR( " LeftMargin \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_Order( 
            /* [retval][out] */ XlOrder *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_Order( 
            /* [in] */ XlOrder RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_Orientation( 
            /* [retval][out] */ XlPageOrientation *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    VARIANT_BOOL result = VARIANT_FALSE;
    
    result = m_oo_page_style.IsLandscape( );

    switch ( result )
    {
	    case VARIANT_TRUE:
			 *RHS = xlLandscape;
			 break;
		case VARIANT_FALSE: 
			 *RHS = xlPortrait;
			 break; 	   
    }
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_Orientation( 
            /* [in] */ XlPageOrientation RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    VARIANT_BOOL value;
    
    switch ( RHS )
    {
	    case xlLandscape:
			 value = VARIANT_TRUE;
			 break;
		case xlPortrait:
			value = VARIANT_FALSE; 
			break; 	   
    }
    
    hr = m_oo_page_style.IsLandscape( value );
    if ( FAILED( hr ) )
    {
	    ERR( " IsLandscape \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PaperSize( 
            /* [retval][out] */ XlPaperSize *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PaperSize( 
            /* [in] */ XlPaperSize RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintArea( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintArea( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintGridlines( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintGridlines( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintHeadings( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintHeadings( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintNotes( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintNotes( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintQuality( 
            /* [optional][in] */ VARIANT Index,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintQuality( 
            /* [optional][in] */ VARIANT Index,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintTitleColumns( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintTitleColumns( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintTitleRows( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintTitleRows( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_RightFooter( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_RightFooter( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_RightHeader( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_RightHeader( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_RightMargin( 
            /* [retval][out] */ double *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.RightMargin( );
    if ( (*RHS) < 0 )
    {
	    ERR( " RightMargin < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    TRACE_OUT;
    return ( hr ); 			
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_RightMargin( 
            /* [in] */ double RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.RightMargin( RHS );
    if ( FAILED( hr ) )
    {
	    ERR( " RightMargin \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_TopMargin( 
            /* [retval][out] */ double *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    *RHS = m_oo_page_style.TopMargin( );
    if ( (*RHS) < 0 )
    {
	    ERR( " TopMargin < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_TopMargin( 
            /* [in] */ double RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = m_oo_page_style.TopMargin( RHS );
    if ( FAILED( hr ) )
    {
	    ERR( " TopMargin \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::get_Zoom( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    short result;
    
    result = m_oo_page_style.PageScale( );
    if ( result < 0 )
    {
	    ERR( " PageScale < 0 \n" );   	 
	    hr = E_FAIL;
    }
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I2;
    V_I2( RHS ) = result;
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CPageSetup::put_Zoom( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr = S_OK;
    short value;
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I2);
    if ( FAILED( hr ) ) {
        ERR(" VariantChangeType \n");
        return ( hr );
    }
    
    value = V_I2( &RHS );
    
    hr = m_oo_page_style.PageScale( value );
    if ( FAILED( hr ) )
    {
	    ERR( " PageScale \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );	 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintComments( 
            /* [retval][out] */ XlPrintLocation *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintComments( 
            /* [in] */ XlPrintLocation RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_PrintErrors( 
            /* [retval][out] */ XlPrintErrors *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_PrintErrors( 
            /* [in] */ XlPrintErrors RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_CenterHeaderPicture( 
            /* [retval][out] */ Graphic **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_CenterFooterPicture( 
            /* [retval][out] */ Graphic **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_LeftHeaderPicture( 
            /* [retval][out] */ Graphic **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_LeftFooterPicture( 
            /* [retval][out] */ Graphic **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_RightHeaderPicture( 
            /* [retval][out] */ Graphic **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_RightFooterPicture( 
            /* [retval][out] */ Graphic **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_OddAndEvenPagesHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_OddAndEvenPagesHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_DifferentFirstPageHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_DifferentFirstPageHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_ScaleWithDocHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_ScaleWithDocHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_AlignMarginsHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CPageSetup::put_AlignMarginsHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_Pages( 
            /* [retval][out] */ Pages **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_EvenPage( 
            /* [retval][out] */ Page **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CPageSetup::get_FirstPage( 
            /* [retval][out] */ Page **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

HRESULT CPageSetup::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_IPageSetup, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;    		
}
       
HRESULT CPageSetup::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;    		
}

HRESULT CPageSetup::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;
   
   TRACE_OUT;
   return S_OK;  		
} 

HRESULT CPageSetup::InitWrapper( OOPageStyle& oo_page_style)
{
    m_oo_page_style = oo_page_style;
	
	return ( S_OK ); 		
}

