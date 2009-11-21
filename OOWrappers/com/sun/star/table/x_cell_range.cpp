/*
 * source file - XCellRange
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

#include "x_cell_range.h"
   
com::sun::star::table::XCellRange::XCellRange( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::table::XCellRange::~XCellRange( )
{              							
} 

com::sun::star::table::XCell com::sun::star::table::XCellRange::getCellByPosition( 
	  long _column, 
	  long _row)
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, param1, param2;
	com::sun::star::table::XCell ret_val;    
    
    VariantInit( &res );
    VariantInit( &param1 );
    VariantInit( &param2 );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = _column;
    V_VT( &param2 ) = VT_I4;
    V_I4( &param2 ) = _row;
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getCellByPosition", 2, param2, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call getCellByPosition \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }
    
    VariantClear( &res );
    VariantClear( &param1 );
    VariantClear( &param2 );
 
    TRACE_OUT;
    return ( ret_val );   	  
}

com::sun::star::table::XCellRange com::sun::star::table::XCellRange::getCellRangeByPosition( 
		   long _left, 
		   long _top, 
		   long _right, 
		   long _bottom)
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, param1, param2, param3, param4;
	com::sun::star::table::XCellRange ret_val;    
    
    VariantInit( &res );
    VariantInit( &param1 );
    VariantInit( &param2 );
    VariantInit( &param3 );
    VariantInit( &param4 );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = _left;
    V_VT( &param2 ) = VT_I4;
    V_I4( &param2 ) = _top;
    V_VT( &param3 ) = VT_I4;
    V_I4( &param3 ) = _right;
    V_VT( &param4 ) = VT_I4;
    V_I4( &param4 ) = _bottom;
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getCellRangeByPosition", 4, param4, param3, param2, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call getCellRangeByPosition \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }
    
    VariantClear( &res );
    VariantClear( &param1 );
    VariantClear( &param2 );
    VariantClear( &param3 );
    VariantClear( &param4 );
 
    TRACE_OUT;
    return ( ret_val );		   
}

com::sun::star::table::XCellRange com::sun::star::table::XCellRange::getCellRangeByName( BSTR _name )
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, param1;
	com::sun::star::table::XCellRange ret_val;    
    
    VariantInit( &res );
    VariantInit( &param1 );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    V_VT( &param1 ) = VT_BSTR;
    V_BSTR( &param1 ) = SysAllocString( _name );
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getCellRangeByName", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call getCellRangeByName \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }
    
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( ret_val );								  
}
