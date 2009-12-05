/*
 * source file - CellRangeAddress
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

#include "cell_range_address.h"
   
com::sun::star::table::CellRangeAddress::CellRangeAddress( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::table::CellRangeAddress::~CellRangeAddress( )
{              							
} 

long com::sun::star::table::CellRangeAddress::Sheet()
{
 	 
}

long com::sun::star::table::CellRangeAddress::StartColumn()
{
    TRACE_IN;
	VARIANT res;	 
	HRESULT hr;
	long ret_val = 0;
		
	if ( IsNull() )
	{
	   	ERR( " m_pd_wrapper is NULL \n" ); 
	    return ( 0 );   	 
    }	
	
	VariantInit( &res );
		
	hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"StartColumn",0 );
	if ( FAILED( hr ) )
	{
	    ERR( " failed StartColumn \n" );   	 
	    TRACE_OUT;
	    return ( 0 );
    } 
	
	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
	
	ret_val = V_I4( &res );
	
	VariantClear( &res ); 
	 
    TRACE_OUT;
	return ( ret_val );  	 
}

long com::sun::star::table::CellRangeAddress::StartRow()
{
 	 
}

long com::sun::star::table::CellRangeAddress::EndColumn()
{
 	 
}

long com::sun::star::table::CellRangeAddress::EndRow()
{
 	 
}
