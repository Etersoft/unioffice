/*
 * source file - XCell
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

#include "x_cell.h"
   
com::sun::star::table::XCell::XCell( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::table::XCell::~XCell( )
{
 									  
									                 							
} 

HRESULT com::sun::star::table::XCell::getFormula( VARIANT *result )
{ 
    TRACE_IN;	 
	HRESULT hr;
		
	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( E_FAIL );   	 
    }	
	
	hr = AutoWrap(DISPATCH_METHOD, result, m_pd_wrapper, L"getFormula", 0);
	if ( FAILED( hr ) )
	{
	    ERR( " failed setFormula \n" );   	 
	    TRACE_OUT;
	    return ( E_FAIL );
    } 
	 
    TRACE_OUT;
	return ( hr );	 	                							
} 

HRESULT com::sun::star::table::XCell::setFormula( BSTR _value)
{
 	TRACE_IN;	 
    HRESULT hr;
    VARIANT res, param1;

	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( E_FAIL );   	 
    }

	VariantInit( &res );
	VariantInit( &param1 );

	V_VT( &param1 ) = VT_BSTR;
	V_BSTR( &param1 ) = SysAllocString( _value );

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"setFormula", 1, param1);
	if ( FAILED( hr ) )
	{
	    ERR( " failed setFormula \n" );   	 
    }

	VariantClear( &res );
	VariantClear( &param1 );

    TRACE_OUT;
	return ( hr );  		              							
} 
			  
HRESULT com::sun::star::table::XCell::getValue( VARIANT *result )
{       
    TRACE_IN;	 
	HRESULT hr;
		
	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( E_FAIL );   	 
    }	
	
	hr = AutoWrap(DISPATCH_METHOD, result, m_pd_wrapper, L"getValue", 0);
	if ( FAILED( hr ) )
	{
	    ERR( " failed setFormula \n" );   	 
	    TRACE_OUT;
	    return ( E_FAIL );
    } 
	 
    TRACE_OUT;
	return ( hr );		       							
} 

HRESULT com::sun::star::table::XCell::setValue( VARIANT value )
{
 	TRACE_IN;	 
    HRESULT hr;
    VARIANT res;

	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( E_FAIL );   	 
    }

	VariantInit( &res );

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"setValue", 1, value);
	if ( FAILED( hr ) )
	{
	    ERR( " failed setValue \n" );   	 
    }

	VariantClear( &res );

    TRACE_OUT;
	return ( hr );  				               							
} 

HRESULT com::sun::star::table::XCell::setString( BSTR _value)
{
 	TRACE_IN;	 
    HRESULT hr;
    VARIANT res, param1;

	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( E_FAIL );   	 
    }

	VariantInit( &res );
	VariantInit( &param1 );

	V_VT( &param1 ) = VT_BSTR;
	V_BSTR( &param1 ) = SysAllocString( _value );

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"setString", 1, param1);
	if ( FAILED( hr ) )
	{
	    ERR( " failed setString \n" );   	 
    }

	VariantClear( &res );
	VariantClear( &param1 );

    TRACE_OUT;
	return ( hr ); 		              							
} 
		  
com::sun::star::table::CellContentType com::sun::star::table::XCell::getType()
{
    TRACE_IN;
	VARIANT res;	 
	HRESULT hr;
	com::sun::star::table::CellContentType    ret_val = com::sun::star::table::EMPTY;
		
	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( ret_val );   	 
    }	
	
	VariantInit( &res );
	
	hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"getType", 0);
	if ( FAILED( hr ) )
	{
	    ERR( " failed setFormula \n" );   	 
	    TRACE_OUT;
	    return ( ret_val );
    } 
	
	ret_val = static_cast<com::sun::star::table::CellContentType>(V_I4( &res ));
	
	VariantClear( &res );
	 
    TRACE_OUT;
	return ( ret_val ); 									   
}			  
			  
long com::sun::star::table::XCell::getError()
{ 
  	              							
} 
 
