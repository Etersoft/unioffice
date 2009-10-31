/*
 * implementation of OODispatchHelper
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

#include "./oo_dispatch_helper.h"

using namespace com::sun::star::uno;

OODispatchHelper::OODispatchHelper():XBase()
{                    
}

OODispatchHelper::~OODispatchHelper()
{                      
}    

HRESULT OODispatchHelper::executeDispatch( 
		OODispatchProvider dispatch_provider, 
		BSTR url, 
		BSTR target_frame_name, 
		long search_flags, 
		WrapPropertyArray& arguments)
{
    TRACE_IN;
	HRESULT hr;
	VARIANT param1, param2, param3, param4, param5;
	VARIANT res;
	
	VariantInit( &param1 );
	VariantInit( &param2 );
	VariantInit( &param3 );
	VariantInit( &param4 );
	VariantInit( &param5 );
	VariantInit( &res );
	
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        TRACE_OUT;
        return ( E_FAIL );      
    } 
	
	V_VT(&param1) = VT_DISPATCH;
    V_DISPATCH(&param1) = dispatch_provider.GetIDispatch();
    V_VT(&param2) = VT_BSTR;
    V_BSTR(&param2) = SysAllocString( url );
    V_VT(&param3) = VT_BSTR;
    V_BSTR(&param3) = SysAllocString( target_frame_name );
    V_VT(&param4) = VT_I4;
    V_I4(&param4) = search_flags;
    V_VT(&param5) = VT_ARRAY | VT_DISPATCH;
    V_ARRAY(&param5) = arguments.Get_SafeArray();	
	
	hr = AutoWrap (
		 DISPATCH_METHOD, 
		 &res, 
		 m_pd_wrapper, 
		 L"executeDispatch", 
		 5, 
		 param5, 
		 param4, 
		 param3, 
		 param2, 
		 param1);
    
	if ( FAILED( hr ) )
	{
	    ERR( " executeDispatch \n" );   	 
    }
	
	VariantClear( &param1 );
	VariantClear( &param2 );
	VariantClear( &param3 );
	VariantClear( &param4 );
	VariantClear( &param5 );
	VariantClear( &res );
		
	TRACE_OUT;
	return ( hr ); 		
}
