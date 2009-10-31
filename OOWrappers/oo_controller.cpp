/*
 * implementation of OOController
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

#include "./oo_controller.h"

using namespace com::sun::star::uno;

OOController::OOController():XBase()
{                  
}

OOController::~OOController()
{                        
}    

HRESULT OOController::getFrame( OOFrame& oo_frame)
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_disp;
    VARIANT res;
     
    VariantInit(&res);
    
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        TRACE_OUT;
        return ( E_FAIL );      
    } 
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"getFrame", 0);
    
    p_disp = V_DISPATCH( &res );
    
    if ( FAILED( hr ) ) {
        ERR(" getFrame \n ");
        TRACE_OUT;
        return ( hr );
    }
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_frame.Init( p_disp );
    
    VariantClear( &res ); 
    
    TRACE_OUT;
    return ( hr ); 		
}
