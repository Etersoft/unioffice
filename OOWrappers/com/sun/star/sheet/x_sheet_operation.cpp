/*
 * source file - XSheetOperation
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

#include "x_sheet_operation.h"
   
com::sun::star::sheet::XSheetOperation::XSheetOperation( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::sheet::XSheetOperation::~XSheetOperation( )
{              							
} 

HRESULT com::sun::star::sheet::XSheetOperation::clearContents( long value )
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
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = value;    
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_wrapper, L"clearContents", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call clearContents \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );  		
}
