/*
 * implementation of OOPageStyle
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

#include "../OOWrappers/oo_page_style.h"


OOPageStyle::OOPageStyle()
{
    TRACE_IN;
                                    
    m_pd_page_style = NULL;                                   
    
    TRACE_OUT;                   
}

OOPageStyle::OOPageStyle(const OOPageStyle &obj)
{
   TRACE_IN;    
                               
   m_pd_page_style = obj.m_pd_page_style;
   if ( m_pd_page_style != NULL )
       m_pd_page_style->AddRef();  
       
   TRACE_OUT;                        
}
                       
OOPageStyle::~OOPageStyle()
{
   TRACE_IN;                    
                     
   if ( m_pd_page_style != NULL )
   {
       m_pd_page_style->Release();
       m_pd_page_style = NULL;        
   }
   
   TRACE_OUT;
}
   
OOPageStyle& OOPageStyle::operator=( const OOPageStyle &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_page_style != NULL )
   {
       m_pd_page_style->Release();
       m_pd_page_style = NULL;        
   } 
   
   m_pd_page_style = obj.m_pd_page_style;
   if ( m_pd_page_style != NULL )
       m_pd_page_style->AddRef();
   
   return ( *this );         
}
  
void OOPageStyle::Init( IDispatch* p_oo_page_style)
{
   TRACE_IN; 
     
   if ( m_pd_page_style != NULL )
   {
       m_pd_page_style->Release();
       m_pd_page_style = NULL;        
   } 
   
   if ( p_oo_page_style == NULL )
   {
       ERR( " p_oo_page_style == NULL \n" );
       return;     
   }
   
   m_pd_page_style = p_oo_page_style;
   m_pd_page_style->AddRef();
   
   TRACE_OUT;
   
   return;          
}
  
bool OOPageStyle::IsNull()
{
    return ( (m_pd_page_style == NULL) ? true : false );     
}

double OOPageStyle::LeftMargin( )
{
    TRACE_IN;
	double result;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( -1.0 ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_page_style, L"LeftMargin", 0);
    if ( FAILED( hr ) ) {
        ERR(" LeftMargin (GET) \n");
        TRACE_OUT;
        return ( -1.0 );
    } 
    
    hr = VariantChangeTypeEx(&res, &res, 0, 0, VT_R8);
    if ( FAILED( hr ) ) {
        ERR(" VariantChangeType \n");
        return ( -1.0 );
    }

    result = V_R8(&res) / 1000 * 28;
    
    VarR8Round( result, 0, &result );
    
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result ); 	   
}

HRESULT OOPageStyle::LeftMargin( double _value )
{
    TRACE_IN;
	double result;
	HRESULT hr;
	VARIANT res, param1;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( hr ) ;   	 
    }
	
	VariantInit( &res );
	VariantInit( &param1 );
	
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = static_cast<long>( _value / 28 * 1000 );
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_page_style, L"LeftMargin", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" LeftMargin (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );  		
}

double OOPageStyle::RightMargin( )
{
    TRACE_IN;
	double result;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( -1.0 ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_page_style, L"RightMargin", 0);
    if ( FAILED( hr ) ) {
        ERR(" LeftMargin (GET) \n");
        TRACE_OUT;
        return ( -1.0 );
    } 
    
    hr = VariantChangeTypeEx(&res, &res, 0, 0, VT_R8);
    if ( FAILED( hr ) ) {
        ERR(" VariantChangeType \n");
        return ( -1.0 );
    }

    result = V_R8(&res) / 1000 * 28;
    
    VarR8Round( result, 0, &result );
    
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result ); 	   
}

HRESULT OOPageStyle::RightMargin( double _value )
{
    TRACE_IN;
	double result;
	HRESULT hr;
	VARIANT res, param1;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( hr ) ;   	 
    }
	
	VariantInit( &res );
	VariantInit( &param1 );
	
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = static_cast<long>( _value / 28 * 1000 );
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_page_style, L"RightMargin", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" LeftMargin (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );  		
}

double OOPageStyle::TopMargin( )
{
    TRACE_IN;
	double result;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( -1.0 ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_page_style, L"TopMargin", 0);
    if ( FAILED( hr ) ) {
        ERR(" LeftMargin (GET) \n");
        TRACE_OUT;
        return ( -1.0 );
    } 
    
    hr = VariantChangeTypeEx(&res, &res, 0, 0, VT_R8);
    if ( FAILED( hr ) ) {
        ERR(" VariantChangeType \n");
        return ( -1.0 );
    }

    result = V_R8(&res) / 1000 * 28;
    
    VarR8Round( result, 0, &result );
    
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result ); 	   
}

HRESULT OOPageStyle::TopMargin( double _value )
{
    TRACE_IN;
	double result;
	HRESULT hr;
	VARIANT res, param1;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( hr ) ;   	 
    }
	
	VariantInit( &res );
	VariantInit( &param1 );
	
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = static_cast<long>( _value / 28 * 1000 );
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_page_style, L"TopMargin", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" LeftMargin (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );  		
}
