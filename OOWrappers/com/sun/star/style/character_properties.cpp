/*
 * source file - CharacterProperties
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

#include "character_properties.h"
   
com::sun::star::style::CharacterProperties::CharacterProperties( ):com::sun::star::uno::XBase()
{
 																                                                                        
}

com::sun::star::style::CharacterProperties::~CharacterProperties( )
{   
	           							
} 

HRESULT com::sun::star::style::CharacterProperties::setCharFontName( BSTR value)
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, param1;    
    
    VariantInit( &res );
    VariantInit( &param1 );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        TRACE_OUT;
        return ( E_FAIL );     
    }
    
    V_VT( &param1 ) = VT_BSTR;
    V_BSTR( &param1 ) = SysAllocString( value );
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CharFontName", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CharFontName \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr ); 		
}

HRESULT com::sun::star::style::CharacterProperties::getCharFontName( BSTR& value )
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;    
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        TRACE_OUT;
        return ( E_FAIL );     
    }
    
    hr = AutoWrap (DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CharFontName", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CharFontName \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	SysFreeString( value );
 	value = SysAllocString( V_BSTR( &res ) );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr ); 	 		
}

HRESULT com::sun::star::style::CharacterProperties::setCharShadowed( bool value )
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, param1;    
    
    VariantInit( &res );
    VariantInit( &param1 );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        TRACE_OUT;
        return ( E_FAIL );     
    }
    
    V_VT( &param1 ) = VT_BOOL;
    if ( value )
        V_BOOL( &param1 ) = VARIANT_TRUE;
    else
	    V_BOOL( &param1 ) = VARIANT_FALSE;    
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CharShadowed", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CharShadowed \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );  		
}

HRESULT com::sun::star::style::CharacterProperties::getCharShadowed( bool& value )
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;    
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        TRACE_OUT;
        return ( E_FAIL );     
    }
    
    hr = AutoWrap (DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CharShadowed", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CharShadowed \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	if ( V_BOOL( &res ) == VARIANT_TRUE )
 	    value = true;
    else
 	    value = false;
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr ); 		
}

