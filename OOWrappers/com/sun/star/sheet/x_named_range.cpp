/*
 * source file - XNamedRange
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

#include "x_named_range.h"
   
com::sun::star::sheet::XNamedRange::XNamedRange( ):com::sun::star::container::XNamed()
{                                                                       
}

com::sun::star::sheet::XNamedRange::~XNamedRange( )
{              							
}

BSTR com::sun::star::sheet::XNamedRange::getContent()
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res; 
	BSTR ret_val = SysAllocString(L"");  
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getContent",0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call getContent \n" );
    } else
    {
	  	SysFreeString( ret_val );  
	  	ret_val = SysAllocString( V_BSTR( &res ) );
    }
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( ret_val );	 
}

HRESULT com::sun::star::sheet::XNamedRange::setContent( BSTR _content)
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, param1;    
    
    VariantInit( &res );
    VariantInit( &param1 );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( E_FAIL );     
    }
    
    V_VT( &param1 ) = VT_BSTR;
    V_BSTR( &param1 ) = SysAllocString( _content );
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"setContent", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call setContent \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );	 		
}
			  
com::sun::star::table::SCellAddress com::sun::star::sheet::XNamedRange::getReferencePosition()
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;
	com::sun::star::table::SCellAddress ret_val;    
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getReferencePosition", 0);
    if ( FAILED( hr ) )
    {
        ERR( " Call getReferencePosition \n" );
    } else
    {
        ret_val.Init( V_DISPATCH( &res ) );      
    }
    
    VariantClear( &res );
 
    TRACE_OUT;
    return ( ret_val ); 									
}

HRESULT com::sun::star::sheet::XNamedRange::setReferencePosition( com::sun::star::table::SCellAddress _cell_address)
{
 	 TRACE_NOTIMPL;	 		
}
			  
long com::sun::star::sheet::XNamedRange::getType()
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res; 
	long ret_val = -1;  
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( ret_val );     
    }
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"getType",0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call getContent \n" );
    } else
    {
	  	ret_val = V_I4( &res );
    }
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( ret_val );		 	 
}

HRESULT com::sun::star::sheet::XNamedRange::setType( long _type )
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res, param1;    
    
    VariantInit( &res );
    VariantInit( &param1 );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        return ( E_FAIL );     
    }
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = _type;
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"setType", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call setType \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );	 		
}
