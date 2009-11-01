/*
 * source file - XIndexAccess
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

#include "x_index_access.h"
   
com::sun::star::container::XIndexAccess::XIndexAccess( ):XElementAccess()
{                                                                       
}

com::sun::star::container::XIndexAccess::~XIndexAccess( )
{              							
} 

long com::sun::star::container::XIndexAccess::getCount()
{
   TRACE_IN;
   VARIANT res;
   HRESULT hr = E_FAIL;
   long count = -1;
   
   VariantInit( &res );
   
   if ( IsNull() )
   {
       ERR( " m_pd_wrapper is NULL \n" );
       VariantClear( &res );
       return ( -1 );     
   }
   
   hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"getCount", 0);
   
   if ( FAILED( hr ) ) 
   {
       ERR(" Call getCount \n");
       count = -1;
   } else
   {
       count = V_I4( &res );        
   }
   
   VariantClear( &res );
   
   TRACE_OUT;
   return ( count ); 	 
}

com::sun::star::uno::XBase com::sun::star::container::XIndexAccess::getByIndex( long index )
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, var_index;
	com::sun::star::uno::XBase ret_val;    
    
    VariantInit( &res );
    VariantInit( &var_index );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    V_VT( &var_index ) = VT_I4;
    V_I4( &var_index ) = index;
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getByIndex", 1, var_index);
    if ( FAILED( hr ) )
    {
        ERR( " Call getByIndex \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }
    
 
    VariantClear( &res );
    VariantClear( &var_index );
 
    TRACE_OUT;
    return ( ret_val ); 						   
}


  
