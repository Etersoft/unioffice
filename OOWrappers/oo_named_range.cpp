/*
 * implementation of OONamedRange
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

#include "./oo_named_range.h"

using namespace com::sun::star::uno;

OONamedRange::OONamedRange():XBase()
{                 
}
                       
OONamedRange::~OONamedRange()
{
}

BSTR OONamedRange::getName( )
{
    TRACE_IN;
    HRESULT hr;
    VARIANT res;
    BSTR result;

	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" );   
		TRACE_OUT;
		return ( SysAllocString( L"" ) );	 
    }
    
    VariantInit( &res );
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"getName", 0);
    if ( FAILED( hr ) )
    {
        ERR( " getName \n" );     
        result = SysAllocString( L"" );
    } else
    {
        result = SysAllocString( V_BSTR( &res ) );      
    }
    
    VariantClear( &res );
    
    TRACE_OUT;     
    return ( result );
}

HRESULT OONamedRange::setName( BSTR bstr_name )
{
    TRACE_IN;
    
    HRESULT hr;
    VARIANT param1, res;

	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" ); 
		TRACE_OUT;
		return ( E_FAIL );  	 
    }
    
    VariantInit( &param1 );
    VariantInit( &res );  
        
    V_VT(&param1)   = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(bstr_name);

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"setName", 1, param1);
    
    if ( FAILED( hr ) )
    {
        ERR( " setName \n" );     
    }    
    
    VariantClear( &res );
    VariantClear( &param1 );
    
    TRACE_OUT;
    return ( hr );      
}
