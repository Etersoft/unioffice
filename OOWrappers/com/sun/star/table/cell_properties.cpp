/*
 * source file - CellProperties
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

#include "cell_properties.h"
   
com::sun::star::table::CellProperties::CellProperties( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::table::CellProperties::~CellProperties( )
{              							
} 

HRESULT com::sun::star::table::CellProperties::setCellBackColor( long value)
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
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"CellBackColor", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call CellBackColor \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );   		
}

HRESULT com::sun::star::table::CellProperties::getCellBackColor( long& value)
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
    
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"CellBackColor", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call CellBackColor \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
 	
 	value = V_I4( &res );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );  		
}

HRESULT com::sun::star::table::CellProperties::setIsCellBackgroundTransparent( bool value)
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
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"IsCellBackgroundTransparent", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call IsCellBackgroundTransparent \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );   		
}

HRESULT com::sun::star::table::CellProperties::getIsCellBackgroundTransparent( bool& value)
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
    
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"IsCellBackgroundTransparent", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call IsCellBackgroundTransparent \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_BOOL);
 	
 	if ( V_BOOL(&res) == VARIANT_TRUE )
 	    value = true;
 	else 
	    value = false;    
 	 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );  		
}
