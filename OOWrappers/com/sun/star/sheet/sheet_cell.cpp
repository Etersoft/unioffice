/*
 * source file - SheetCell
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

#include "sheet_cell.h"
   
com::sun::star::sheet::SheetCell::SheetCell( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::sheet::SheetCell::~SheetCell( )
{              							
} 

com::sun::star::awt::Size com::sun::star::sheet::SheetCell::getSize()
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;
	com::sun::star::awt::Size ret_val;  						  
						   
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
        
    hr = AutoWrap (DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"Size", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call Size \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }						   
						   
    VariantClear( &res );
 
    TRACE_OUT;
    return ( ret_val );							    						  
}

com::sun::star::awt::Point com::sun::star::sheet::SheetCell::getPosition()
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;
	com::sun::star::awt::Point ret_val;  						  
						   
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
        
    hr = AutoWrap (DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"Position", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call Position \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }						   
						   
    VariantClear( &res );
 
    TRACE_OUT;
    return ( ret_val );							    						  
}
