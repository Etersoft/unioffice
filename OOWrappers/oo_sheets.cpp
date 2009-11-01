/*
 * implementation of OOSheets
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

#include "../OOWrappers/oo_sheets.h"

using namespace com::sun::star::uno;

OOSheets::OOSheets():XBase()
{                   
}

OOSheets::~OOSheets()
{                       
}    

HRESULT OOSheets::insertNewByName( BSTR _name, long _index )
{
    TRACE_IN;
    HRESULT hr;
    VARIANT param1, param2;
    VARIANT res;
    
    VariantInit( &param1 );
    VariantInit( &param2 );
    VariantInit( &res );
    
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        return ( E_FAIL );     
    }
    
    V_VT( &param1 ) = VT_BSTR;
    V_BSTR( &param1 ) = SysAllocString( _name );
    
    V_VT( &param2 ) = VT_I4;
    V_I4( &param2 ) = _index;

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"insertNewByName", 2, param2, param1);
    if ( FAILED( hr ) ) 
    {
        ERR(" insertNewByName\n");
    }

    VariantClear( &param1 );
    VariantClear( &param2 );
    VariantClear( &res );
    
    TRACE_OUT;
    return ( hr );          
}

HRESULT OOSheets::removeByName( BSTR _name )
{
    TRACE_IN;
    HRESULT hr;
    
    VARIANT param1;
    VARIANT res;
    
    VariantInit( &param1 );
    VariantInit( &res );

    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        return ( E_FAIL );     
    }
    
    V_VT( &param1 ) = VT_BSTR;
    V_BSTR( &param1 ) = SysAllocString( _name );

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"removeByName", 1, param1);
    if ( FAILED( hr ) ) 
    {
        ERR(" removeByName \n");
    }

    VariantClear( &param1 );
    VariantClear( &res );
       
    TRACE_OUT;
    return ( hr );
}
