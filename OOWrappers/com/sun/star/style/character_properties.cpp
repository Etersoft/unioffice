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

HRESULT com::sun::star::style::CharacterProperties::setCharWeight( float value)
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
    
    V_VT( &param1 ) = VT_R4;
    V_R4( &param1 ) = value;    
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CharWeight", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CharWeight \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr ); 		
}

HRESULT com::sun::star::style::CharacterProperties::getCharWeight( float& value)
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
    
    hr = AutoWrap (DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CharWeight", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CharWeight \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_R4);
 	
 	value = V_R4( &res );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );  		
}

HRESULT com::sun::star::style::CharacterProperties::setCharPosture( com::sun::star::awt::FontSlant value)
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
    
    V_VT( &param1 ) = VT_I2;
    V_I2( &param1 ) = value;    
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CharPosture", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CharPosture \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );  		
}

HRESULT com::sun::star::style::CharacterProperties::getCharPosture( com::sun::star::awt::FontSlant& value)
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
    
    hr = AutoWrap (DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CharPosture", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CharPosture \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_I2);
 	
 	value = static_cast<com::sun::star::awt::FontSlant>( V_I2( &res ) );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr ); 		
}

HRESULT com::sun::star::style::CharacterProperties::setCharColor( long value)
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
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = value;    
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CharColor", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CharColor \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );   		
}

HRESULT com::sun::star::style::CharacterProperties::getCharColor( long& value)
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
    
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CharColor", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CharColor \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
 	
 	value = V_I4( &res );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );  		
}
         	  
HRESULT com::sun::star::style::CharacterProperties::setCharHeight( long value)
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
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = value;    
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CharHeight", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CharHeight \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );   		
}

HRESULT com::sun::star::style::CharacterProperties::getCharHeight( long& value)
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
    
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CharHeight", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CharHeight \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
 	
 	value = V_I4( &res );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );  		
}

HRESULT com::sun::star::style::CharacterProperties::setCharUnderline( long value )
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
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = value;    
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CharUnderline", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CharUnderline \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );   		
}

HRESULT com::sun::star::style::CharacterProperties::getCharUnderline( long& value)
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
    
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CharUnderline", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CharUnderline \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
 	
 	value = V_I4( &res );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );  		
}
