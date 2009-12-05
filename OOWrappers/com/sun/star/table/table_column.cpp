/*
 * source file - TableColumn
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

#include "table_column.h"
   
com::sun::star::table::TableColumn::TableColumn( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::table::TableColumn::~TableColumn( )
{              							
} 

HRESULT com::sun::star::table::TableColumn::getWidth( long& value)
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
		
	hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"Width",0 );
	if ( FAILED( hr ) )
	{
	    ERR( " failed Width \n" );   	 
	    TRACE_OUT;
	    return ( E_FAIL );
    } 
	
	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
	
	value = V_I4( &res );
	
	VariantClear( &res ); 
	 
    TRACE_OUT;
	return ( hr ); 	 		
}

HRESULT com::sun::star::table::TableColumn::setWidth( long value )
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
	
	hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"Width", 1, param1 );
	if ( FAILED( hr ) )
	{
	    ERR( " failed Width \n" );   	 
	    TRACE_OUT;
	    return ( E_FAIL );
    } 
	
	
	VariantClear( &res );
	VariantClear( &param1 ); 
	 
    TRACE_OUT;
	return ( hr ); 		
}

HRESULT com::sun::star::table::TableColumn::getOptimalWidth( bool& value)
{
 		
}

HRESULT com::sun::star::table::TableColumn::setOptimalWidth( bool value)
{
 		
}

HRESULT com::sun::star::table::TableColumn::getIsVisible( bool& value)
{
 		
}

HRESULT com::sun::star::table::TableColumn::setIsVisible( bool value)
{
 		
}

HRESULT com::sun::star::table::TableColumn::getIsStartOfNewPage( bool& value)
{
 		
}

HRESULT com::sun::star::table::TableColumn::setIsStartOfNewPage( bool value)
{
 		
}

