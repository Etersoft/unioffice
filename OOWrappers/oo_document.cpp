/*
 * implementation of OODocument
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

#include "../OOWrappers/oo_document.h"

OODocument::OODocument()
{
    TRACE_IN;
                                    
    m_pd_document = NULL;                                   
    
    TRACE_OUT;                        
}

OODocument::OODocument(const OODocument &obj)
{
   TRACE_IN;            
                               
   m_pd_document = obj.m_pd_document;
   if ( m_pd_document != NULL )
       m_pd_document->AddRef();  
       
   TRACE_OUT;                      
}

OODocument::~OODocument()
{
   TRACE_IN;
   
   if ( m_pd_document != NULL )
   {
       m_pd_document->Release();
       m_pd_document = NULL;        
   }                                  
   
   TRACE_OUT;                         
}

OODocument& OODocument::operator=( const OODocument &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_document != NULL )
   {
       m_pd_document->Release();
       m_pd_document = NULL;        
   } 
   
   m_pd_document = obj.m_pd_document;
   if ( m_pd_document != NULL )
       m_pd_document->AddRef();
   
   return ( *this );          
    
}

void OODocument::Init( IDispatch* p_oo_document )
{
   TRACE_IN; 
     
   if ( m_pd_document != NULL )
   {
       m_pd_document->Release();
       m_pd_document = NULL;        
   } 
   
   if ( p_oo_document == NULL )
   {
       ERR( " p_oo_document == NULL \n" );
       TRACE_OUT;
       return;     
   }
   
   m_pd_document = p_oo_document;
   m_pd_document->AddRef();
   
   TRACE_OUT;
   
   return;
}

bool OODocument::IsNull()
{
    return ( (m_pd_document == NULL) ? true : false );     
}

HRESULT OODocument::StoreAsURL( BSTR _filename, WrapPropertyArray& _property_array )
{
    HRESULT hr;
    VARIANT param0,param1;
    VARIANT result;
  
    TRACE_IN;

    if ( IsNull() )
    {
        ERR( " m_pd_document is NULL \n" );
		TRACE_OUT; 
        return ( E_FAIL );     
    }
  
    VariantInit(&param0);
    VariantInit(&param1);
    VariantInit(&result); 
     
    //type of document
    V_VT(&param0) = VT_BSTR;
    V_BSTR(&param0) = SysAllocString( _filename );
    
    V_VT(&param1) = VT_ARRAY | VT_DISPATCH;
    V_ARRAY(&param1) = _property_array.Get_SafeArray();
    
    TRACE( " Filename = " );
    int i=0;
    while (*(_filename+i)!=0) {
        WTRACE_HARD(L"%c",*(_filename+i));
        i++;
    }
    WTRACE_HARD(L"\n");
        
    hr = AutoWrap(DISPATCH_METHOD, &result, m_pd_document, L"StoreAsURL", 2, param1, param0);
    
    if ( FAILED( hr ) ) {
        ERR( " StoreAsURL \n" );
		TRACE_OUT; 
        return ( E_FAIL );
    }
    
    VariantClear(&param0);
    VariantClear(&param1);
    VariantClear(&result);
  
    TRACE_OUT;
     
    return ( hr );    
} 

HRESULT OODocument::Store( )
{
    TRACE_IN;
    HRESULT hr;
    VARIANT res;

    if ( IsNull() )
    {
        ERR( " m_pd_document is NULL \n" );
		TRACE_OUT; 
        return ( E_FAIL );     
    }
    
    VariantInit( &res );
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_document, L"Store", 0);
    if ( FAILED( hr ) )
    {
        ERR( " Store() \n" );     
    }
    
    VariantClear( &res );
    
    TRACE_OUT;
    return ( hr );    
}

HRESULT OODocument::close( VARIANT_BOOL _hard_close )
{
    HRESULT hr;
    VARIANT res;
    VARIANT hard_close;
    
    TRACE_IN;

    if ( IsNull() )
    {
        ERR( " m_pd_document is NULL \n" );
		TRACE_OUT; 
        return ( E_FAIL );     
    }
    
    VariantInit( &hard_close );
    VariantInit( &res );
    
    V_VT( &hard_close )   = VT_BOOL;
    V_BOOL( &hard_close ) = _hard_close;
    
    hr = AutoWrap( DISPATCH_METHOD, &res, m_pd_document, L"close", 1, hard_close );

    if ( FAILED( hr ) )
    { 
            ERR(" FAILED 1 CLOSE \n"); 
    }
            
    VariantClear( &res );
    VariantClear( &hard_close );       
               
    TRACE_OUT;        
    return ( hr );          
}

HRESULT OODocument::getSheets( OOSheets& oo_sheets )
{
    TRACE_IN;
    IDispatch* p_disp;
    VARIANT res;
    HRESULT hr;
     
    VariantInit(&res);
    
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        VariantClear( &res );
        TRACE_OUT;
        return ( E_FAIL );      
    } 
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_document, L"getSheets", 0);

    p_disp = V_DISPATCH( &res );
    
    if ( FAILED( hr ) ) {
        ERR(" getSheets \n ");
        TRACE_OUT;
        return ( hr );
    }
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_sheets.Init( p_disp );
    
    VariantClear( &res );
     
    TRACE_OUT;    
    return ( hr );       
}

HRESULT OODocument::protect( BSTR _password )
{
   TRACE_IN;
   HRESULT hr;
   VARIANT res;
   VARIANT param1;

    if ( IsNull() )
    {
        ERR( " m_pd_document is NULL \n" );
		TRACE_OUT; 
        return ( E_FAIL );     
    }
   
   VariantInit( &res );
   VariantInit( &param1 );
   
   V_VT( &param1 )   = VT_BSTR;
   V_BSTR( &param1 ) = SysAllocString( _password );
   
   hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_document, L"protect", 1, param1);
   if ( FAILED( hr ) )
   {
       ERR( " protect \n" );     
   }
   
   VariantClear( &res );
   VariantClear( &param1 );
   
   TRACE_OUT;
   return ( hr );         
}

HRESULT OODocument::unprotect( BSTR _password )
{
   TRACE_IN;
   HRESULT hr;
   VARIANT res;
   VARIANT param1;

   if ( IsNull() )
   {
       ERR( " m_pd_document is NULL \n" );
       TRACE_OUT; 
       return ( E_FAIL );     
   }
   
   VariantInit( &res );
   VariantInit( &param1 );
   
   V_VT( &param1 )   = VT_BSTR;
   V_BSTR( &param1 ) = SysAllocString( _password );
   
   hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_document, L"unprotect", 1, param1);
   if ( FAILED( hr ) )
   {
       ERR( " unprotect \n" );     
   }
   
   VariantClear( &res );
   VariantClear( &param1 );
   
   TRACE_OUT;
   return ( hr );        
}

HRESULT OODocument::getCurrentController( OOController& oo_controller )
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_disp;
    VARIANT res;
     
    VariantInit(&res);
    
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        TRACE_OUT;
        return ( E_FAIL );      
    } 
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_document, L"getCurrentController", 0);
    
    p_disp = V_DISPATCH( &res );
    
    if ( FAILED( hr ) ) {
        ERR(" getCurrentController \n ");
        TRACE_OUT;
        return ( hr );
    }
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_controller.Init( p_disp );
    
    VariantClear( &res ); 
    
    TRACE_OUT;
    return ( hr );
}

HRESULT OODocument::StyleFamilies( OOStyleFamilies& oo_style_families )
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_disp;
    VARIANT res;
     
    VariantInit(&res);
    
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        TRACE_OUT;
        return ( E_FAIL );      
    } 
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_document, L"StyleFamilies", 0);
    
    p_disp = V_DISPATCH( &res );
    	
    if ( FAILED( hr ) ) {
        ERR(" DISPATCH_PROPERTYGET - StyleFamilies \n ");
        TRACE_OUT;
        return ( hr );
    }
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_style_families.Init( p_disp );
    
    VariantClear( &res ); 
    
    TRACE_OUT;
    return ( hr );		 		
}

HRESULT OODocument::NamedRanges( OONamedRanges& oo_named_ranges)
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_disp;
    VARIANT res;
     
    VariantInit(&res);
    
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        TRACE_OUT;
        return ( E_FAIL );      
    } 
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_document, L"NamedRanges", 0);
    
    p_disp = V_DISPATCH( &res );
    	
    if ( FAILED( hr ) ) {
        ERR(" DISPATCH_PROPERTYGET - NamedRanges \n ");
        TRACE_OUT;
        return ( hr );
    }
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_named_ranges.Init( p_disp );
    
    VariantClear( &res ); 
    
    TRACE_OUT;
    return ( hr );			
}
