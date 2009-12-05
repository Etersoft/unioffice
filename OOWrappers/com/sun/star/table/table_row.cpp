/*
 * source file - TableRow
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

#include "table_row.h"
   
com::sun::star::table::TableRow::TableRow( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::table::TableRow::~TableRow( )
{              							
} 

HRESULT com::sun::star::table::TableRow::getHeight( long& value)
{
    TRACE_IN;
	VARIANT res;	 
	HRESULT hr;
		
	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( E_FAIL );   	 
    }	
	
	VariantInit( &res );
		
	hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"Height",0 );
	if ( FAILED( hr ) )
	{
	    ERR( " failed Height \n" );   	 
	    TRACE_OUT;
	    return ( E_FAIL );
    } 
	
	value = V_I4( &res );
	
	VariantClear( &res ); 
	 
    TRACE_OUT;
	return ( hr ); 		
}

HRESULT com::sun::star::table::TableRow::setHeight( long value)
{
    TRACE_IN;
	VARIANT res, param1;	 
	HRESULT hr;
		
	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( E_FAIL );   	 
    }	
	
	VariantInit( &res );
	VariantInit( &param1 );
	
	V_VT( &param1 ) = VT_I4;
	V_I4( &param1 ) = value;
	
	hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"Height", 1, param1 );
	if ( FAILED( hr ) )
	{
	    ERR( " failed Height \n" );   	 
	    TRACE_OUT;
	    return ( E_FAIL );
    } 
	
	
	VariantClear( &res );
	VariantClear( &param1 ); 
	 
    TRACE_OUT;
	return ( hr ); 		
}
	
HRESULT com::sun::star::table::TableRow::getOptimalHeight( bool& value)
{
 		
}

HRESULT com::sun::star::table::TableRow::setOptimalHeight( bool value)
{
 		
}

HRESULT com::sun::star::table::TableRow::getIsVisible( bool& value)
{
 		
}

HRESULT com::sun::star::table::TableRow::setIsVisible( bool value)
{
 		
}

HRESULT com::sun::star::table::TableRow::getIsStartOfNewPage( bool& value)
{
 		
}

HRESULT com::sun::star::table::TableRow::setIsStartOfNewPage( bool value)
{
 		
}
