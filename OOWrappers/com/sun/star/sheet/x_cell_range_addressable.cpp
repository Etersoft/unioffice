/*
 * header file - XCellRangeAddressable
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
 
#include "x_cell_range_addressable.h"

com::sun::star::sheet::XCellRangeAddressable::XCellRangeAddressable( ):com::sun::star::uno::XBase()
{
 																	 
}

com::sun::star::sheet::XCellRangeAddressable::~XCellRangeAddressable( )
{
 																	  
}															 
			  
com::sun::star::table::CellRangeAddress  com::sun::star::sheet::XCellRangeAddressable::getRangeAddress()
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;
	com::sun::star::table::CellRangeAddress ret_val;    
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
        
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getRangeAddress", 0);
    if ( FAILED( hr ) )
    {
        ERR( " Call getRangeAddress \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( ret_val ); 										 
} 
