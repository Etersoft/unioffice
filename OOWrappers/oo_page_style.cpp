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

#include "./oo_page_style.h"

using namespace com::sun::star::uno;

OOPageStyle::OOPageStyle():XBase()
{                   
}
                       
OOPageStyle::~OOPageStyle()
{
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
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"LeftMargin", 0);
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
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"LeftMargin", 1, param1);
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
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"RightMargin", 0);
    if ( FAILED( hr ) ) {
        ERR(" RightMargin (GET) \n");
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
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"RightMargin", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" RightMargin (PUT) \n");
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
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"TopMargin", 0);
    if ( FAILED( hr ) ) {
        ERR(" TopMargin (GET) \n");
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
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"TopMargin", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" TopMargin (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );  		
}

double OOPageStyle::BottomMargin( )
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
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"BottomMargin", 0);
    if ( FAILED( hr ) ) {
        ERR(" BottomMargin (GET) \n");
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

HRESULT OOPageStyle::BottomMargin( double _value )
{
    TRACE_IN;
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
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"BottomMargin", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" BottomMargin (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );  		
}

VARIANT_BOOL OOPageStyle::IsLandscape( )
{
    TRACE_IN;
	VARIANT_BOOL result = VARIANT_FALSE;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( result ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"IsLandscape", 0);
    if ( FAILED( hr ) ) {
        ERR(" IsLandscape (GET) \n");
        TRACE_OUT;
        return ( result );
    } 
        
    result = V_BOOL( &res );    
        
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result ); 			 
}

HRESULT OOPageStyle::IsLandscape( VARIANT_BOOL _value )
{
    TRACE_IN;
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
	
    V_VT( &param1 ) = VT_BOOL;
    V_BOOL( &param1 ) = _value;
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"IsLandscape", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" IsLandscape (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );   			 
}

short  OOPageStyle::PageScale()
{
    TRACE_IN;
	short result;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( -1 ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"PageScale", 0);
    if ( FAILED( hr ) ) {
        ERR(" PageScale (GET) \n");
        TRACE_OUT;
        return ( -1 );
    } 
        
    result = V_I2( &res );    
        
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result );  		 
}

HRESULT  OOPageStyle::PageScale( short _value )
{
    TRACE_IN;
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
	
    V_VT( &param1 ) = VT_I2;
    V_I2( &param1 ) = _value;
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"PageScale", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" PageScale (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr ); 		 
}

short  OOPageStyle::ScaleToPagesY()
{
    TRACE_IN;
	short result;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( -1 ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"ScaleToPagesY", 0);
    if ( FAILED( hr ) ) {
        ERR(" ScaleToPagesY (GET) \n");
        TRACE_OUT;
        return ( -1 );
    } 
        
    result = V_I2( &res );    
        
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result );  		 
}

HRESULT  OOPageStyle::ScaleToPagesY( short _value )
{
    TRACE_IN;
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
	
    V_VT( &param1 ) = VT_I2;
    V_I2( &param1 ) = _value;
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"ScaleToPagesY", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" ScaleToPagesY (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr ); 		 
}

short  OOPageStyle::ScaleToPagesX()
{
    TRACE_IN;
	short result;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( -1 ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"ScaleToPagesX", 0);
    if ( FAILED( hr ) ) {
        ERR(" ScaleToPagesX (GET) \n");
        TRACE_OUT;
        return ( -1 );
    } 
        
    result = V_I2( &res );    
        
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result );  		 
}

HRESULT  OOPageStyle::ScaleToPagesX( short _value )
{
    TRACE_IN;
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
	
    V_VT( &param1 ) = VT_I2;
    V_I2( &param1 ) = _value;
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"ScaleToPagesX", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" ScaleToPagesX (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr ); 		 
}

double OOPageStyle::HeaderHeight( )
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
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"HeaderHeight", 0);
    if ( FAILED( hr ) ) {
        ERR(" HeaderHeight (GET) \n");
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

HRESULT OOPageStyle::HeaderHeight( double _value )
{
    TRACE_IN;
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
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"HeaderHeight", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" HeaderHeight (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );  		
}

double OOPageStyle::FooterHeight( )
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
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"FooterHeight", 0);
    if ( FAILED( hr ) ) {
        ERR(" FooterHeight (GET) \n");
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

HRESULT OOPageStyle::FooterHeight( double _value )
{
    TRACE_IN;
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
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"FooterHeight", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" FooterHeight (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );  		
}

VARIANT_BOOL OOPageStyle::CenterHorizontally( )
{
    TRACE_IN;
	VARIANT_BOOL result = VARIANT_FALSE;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( result ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CenterHorizontally", 0);
    if ( FAILED( hr ) ) {
        ERR(" CenterHorizontally (GET) \n");
        TRACE_OUT;
        return ( result );
    } 
        
    result = V_BOOL( &res );    
        
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result ); 			 
}

HRESULT OOPageStyle::CenterHorizontally( VARIANT_BOOL _value )
{
    TRACE_IN;
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
	
    V_VT( &param1 ) = VT_BOOL;
    V_BOOL( &param1 ) = _value;
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CenterHorizontally", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" CenterHorizontally (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );   			 
}

VARIANT_BOOL OOPageStyle::CenterVertically( )
{
    TRACE_IN;
	VARIANT_BOOL result = VARIANT_FALSE;
	HRESULT hr;
	VARIANT res;
	
	if ( IsNull() )
	{
	    ERR( " wrapper is null \n" );
	    TRACE_OUT;
		return ( result ) ;   	 
    }
	
	VariantInit( &res );
	
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CenterVertically", 0);
    if ( FAILED( hr ) ) {
        ERR(" CenterVertically (GET) \n");
        TRACE_OUT;
        return ( result );
    } 
        
    result = V_BOOL( &res );    
        
	VariantClear( &res );
	
	TRACE_OUT;
	return ( result ); 			 
}

HRESULT OOPageStyle::CenterVertically( VARIANT_BOOL _value )
{
    TRACE_IN;
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
	
    V_VT( &param1 ) = VT_BOOL;
    V_BOOL( &param1 ) = _value;
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CenterVertically", 1, param1);
    if ( FAILED( hr ) ) {
        ERR(" CenterVertically (PUT) \n");
        TRACE_OUT;
        return ( hr );
    } 
    
	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );   			 
}
