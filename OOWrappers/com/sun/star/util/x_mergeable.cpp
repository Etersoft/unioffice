/*
 * source file - XMergeable
 *
 * Copyright (C) 2010 Sinitsin Ivan (Etersoft) <ivan@etersoft.ru>
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

#include "x_mergeable.h"
   
com::sun::star::util::XMergeable::XMergeable( ):com::sun::star::uno::XBase()
{
 																                                                                        
}

com::sun::star::util::XMergeable::~XMergeable( )
{   
	           							
} 

HRESULT com::sun::star::util::XMergeable::merge( bool value )
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
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"merge", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call Width \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );		
}

HRESULT com::sun::star::util::XMergeable::getIsMerged( bool& value)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL;  		
}

