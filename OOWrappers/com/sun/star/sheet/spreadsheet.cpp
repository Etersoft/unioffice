/*
 * source file - Spreadsheet
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

#include "spreadsheet.h"
   
com::sun::star::sheet::Spreadsheet::Spreadsheet( ):com::sun::star::uno::XBase()
{                                                                       
}

com::sun::star::sheet::Spreadsheet::~Spreadsheet( )
{              							
} 

HRESULT com::sun::star::sheet::Spreadsheet::isVisible( bool _value )
{
 	TRACE_IN;
	HRESULT hr;
	VARIANT res, param1;

	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" );   	 
    }
	
	VariantInit( &res );
	VariantInit( &param1 );
		
	V_VT( &param1 ) = VT_BOOL;
	if ( _value )
	   V_BOOL( &param1 ) = VARIANT_TRUE;	
    else
   	   V_BOOL( &param1 ) = VARIANT_FALSE;
		
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"IsVisible", 1, param1); 		
    
    if ( FAILED( hr ) )
    {
	    ERR( " DISPATCH_PROPERTYPUT - IsVisible \n" );   	 
    }
    
    VariantClear( &res );
    VariantClear( &param1 );
    
    TRACE_OUT;
    return ( hr );
}

bool com::sun::star::sheet::Spreadsheet::isVisible()
{
 	TRACE_IN;
	HRESULT hr;
	VARIANT res;
	bool result = true;

	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" );   	 
    }
	
	VariantInit( &res );
		
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"IsVisible", 0); 		
    
    if ( FAILED( hr ) )
    {
	    ERR( " DISPATCH_PROPERTYGET - IsVisible \n" );   	 
    }
       
    hr = VariantChangeTypeEx(&res, &res, 0, 0, VT_BOOL);
    if ( FAILED( hr ) ) {
        ERR( " VariantChangeTypeEx \n" );
        TRACE_OUT;
        return ( result );    
	}
    
    if ( V_BOOL( &res ) == VARIANT_TRUE )
        result = true;
    else
    	result = false;
    
    VariantClear( &res );
    
    TRACE_OUT;
    return ( result );
}

BSTR com::sun::star::sheet::Spreadsheet::PageStyle()
{
    TRACE_IN;
	BSTR result = SysAllocString(L"");
	HRESULT hr;
	VARIANT res;

	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" );   	 
    }
	
	hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"PageStyle",0);
	if ( FAILED( hr ) )
	{
	    ERR( " PageStyle \n" );   	 
	    TRACE_OUT;
	    return ( result );
    }
	
	SysFreeString( result );
	result = SysAllocString( V_BSTR( &res ) );
	
	TRACE_OUT; 		
	return ( result );
}
