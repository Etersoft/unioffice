/*
 * source file - XModel
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

#include "x_model.h"
   
com::sun::star::frame::XModel::XModel( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::frame::XModel::~XModel( )
{              							
} 

HRESULT com::sun::star::frame::XModel::lockControllers()
{
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;    
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        TRACE_OUT;
        return ( E_FAIL );     
    }
        
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"lockControllers", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call lockControllers \n" );
    }
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );				               							
} 

HRESULT com::sun::star::frame::XModel::unlockControllers()
{ 
    TRACE_IN; 
    HRESULT hr;
    VARIANT res;    
    
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
        TRACE_OUT;
        return ( E_FAIL );     
    }
        
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"unlockControllers", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call unlockControllers \n" );
    }
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );		               							
}
