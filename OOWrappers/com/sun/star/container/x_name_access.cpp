/*
 * source file - XNameAccess
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

#include "x_name_access.h"
   
com::sun::star::container::XNameAccess::XNameAccess( ):XElementAccess()
{                                                                       
}

com::sun::star::container::XNameAccess::~XNameAccess( )
{              							
} 

com::sun::star::uno::XBase com::sun::star::container::XNameAccess::getByName( BSTR name )
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, var_name;
	com::sun::star::uno::XBase ret_val;    
    
    VariantInit( &res );
    VariantInit( &var_name );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    V_VT( &var_name ) = VT_BSTR;
    V_BSTR( &var_name ) = SysAllocString( name );
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getByName", 1, var_name);
    if ( FAILED( hr ) )
    {
        ERR( " Call getByName \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }
 
    VariantClear( &res );
    VariantClear( &var_name );
 
    TRACE_OUT;
    return ( ret_val ); 						   
}


  
