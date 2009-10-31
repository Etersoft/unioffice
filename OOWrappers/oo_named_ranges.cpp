/*
 * implementation of OONamedRanges
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

#include "../OOWrappers/oo_named_ranges.h"

using namespace com::sun::star::uno;

OONamedRanges::OONamedRanges():XBase()
{                  
}
                       
OONamedRanges::~OONamedRanges()
{
}

long OONamedRanges::getCount( )
{
    TRACE_IN;
    long count = -1;
    HRESULT hr;
    VARIANT res;
    
    if ( IsNull() )
    {
	    ERR(" m_pd_wrapper is null \n");   	 
	    
	    TRACE_OUT;
	   	return ( -1 );
    }
    
    VariantInit( &res );
    
	hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"getCount", 0);
	if ( FAILED( hr ) )
	{
        ERR( " getCount \n" );
        
        VariantClear( &res );
        TRACE_OUT;
        return ( -1 );
    }
	
	if ( V_VT( &res ) != VT_I4 )
	{
        hr = VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
        if ( FAILED( hr ) ) {
            ERR( "Error when VariantChangeTypeEx \n" );
            VariantClear( &res );
            TRACE_OUT;        
            return ( -1 );
        }
    }	   	  
    
	count = V_I4( &res );	
	
	VariantClear( &res );
	
	TRACE_OUT;
	return ( count );	 
}

HRESULT OONamedRanges::getNameByName( VARIANT index, OONamedRange& oo_name )
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_disp;
    VARIANT res;
     
    VariantInit(&res);
	    
    if ( IsNull() )
    {
	    ERR(" m_pd_wrapper is null \n");   	 
	    
	    TRACE_OUT;
	   	return ( E_FAIL );
    }    
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"getByName", 1, index );
    
    if (FAILED(hr)) {
        TRACE(" when getByIndex \n");
        
        TRACE_OUT;
        return ( hr );
    }
    
    p_disp = V_DISPATCH( &res );
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_name.Init( p_disp );
    
    VariantClear( &res ); 
    
    TRACE_OUT;
    return ( hr );		
}

HRESULT OONamedRanges::getNameByIndex( VARIANT index, OONamedRange& oo_name)
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_disp;
    VARIANT res;
     
    VariantInit(&res);
	    
    if ( IsNull() )
    {
	    ERR(" m_pd_wrapper is null \n");   	 
	    
	    TRACE_OUT;
	   	return ( E_FAIL );
    }    
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"getByIndex", 1, index );
    
    if (FAILED(hr)) {
        TRACE(" when getByIndex \n");
        
        TRACE_OUT;
        return ( hr );
    }
    
    p_disp = V_DISPATCH( &res );
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_name.Init( p_disp );
    
    VariantClear( &res ); 
    
    TRACE_OUT;
    return ( hr );	 		
}
