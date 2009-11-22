/*
 * implementation of OOSheet
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

#include "./oo_sheet.h"

using namespace com::sun::star::uno;

OOSheet::OOSheet():XBase()
{                 
}
                       
OOSheet::~OOSheet()
{
}

OOSheet& OOSheet::operator=( const XBase &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_wrapper != NULL )
   {
       m_pd_wrapper->Release();
       m_pd_wrapper = NULL;        
   } 

   m_pd_wrapper = obj.m_pd_wrapper;
   if ( m_pd_wrapper != NULL )
       m_pd_wrapper->AddRef();
   
   return ( *this );  		 
}

OOSheet& OOSheet::operator=( const OOSheet &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_wrapper != NULL )
   {
       m_pd_wrapper->Release();
       m_pd_wrapper = NULL;        
   } 

   m_pd_wrapper = obj.m_pd_wrapper;
   if ( m_pd_wrapper != NULL )
       m_pd_wrapper->AddRef();
   
   return ( *this );  		 
}

BSTR OOSheet::getName( )
{
    TRACE_IN;
    HRESULT hr;
    VARIANT res;
    BSTR result;

	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" );   	 
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

HRESULT OOSheet::setName( BSTR bstr_name )
{
    TRACE_IN;
    
    HRESULT hr;
    VARIANT param1, res;

	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" );   	 
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
    VariantClear(&param1 );
    
    TRACE_OUT;
    return ( hr );      
}

HRESULT OOSheet::unprotect( BSTR _password )
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
	
	V_VT( &param1 )   = VT_BSTR;
	V_BSTR( &param1 ) = SysAllocString( _password );
	
	hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"unprotect", 1, param1);
    if ( FAILED( hr ) )
    {
	    ERR( " unprotect \n" );   	 
    }

	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );   		  
} 

HRESULT OOSheet::protect( BSTR _password )
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
	
	V_VT( &param1 )   = VT_BSTR;
	V_BSTR( &param1 ) = SysAllocString( _password );
	
	hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"protect", 1, param1);
    if ( FAILED( hr ) )
    {
	    ERR( " protect \n" );   	 
    }

	VariantClear( &res );
	VariantClear( &param1 );
	
	TRACE_OUT;
	return ( hr );   		  
} 

HRESULT OOSheet::isVisible( VARIANT_BOOL _value )
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
	V_BOOL( &param1 ) = _value;	
		
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

VARIANT_BOOL OOSheet::isVisible()
{
 	TRACE_IN;
	HRESULT hr;
	VARIANT res;
	VARIANT_BOOL result = VARIANT_TRUE;

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
    
    result = V_BOOL( &res );
    
    VariantClear( &res );
    
    TRACE_OUT;
    return ( result );
}

BSTR    OOSheet::PageStyle()
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

HRESULT OOSheet::showLevel( long amount, long type )
{
    TRACE_IN;
    HRESULT hr;
    VARIANT param1, param2, res;
    
	if ( IsNull() )
	{
	    ERR( " m_pd_wrapper is NULL \n" );   	 
    }
	    
    VariantInit( &param1 );
    VariantInit( &param2 );
    VariantInit( &res );    
    
    V_VT( &param1 ) = VT_I4;
    V_I4( &param1 ) = amount;
    
    V_VT( &param2 ) = VT_I4;
    V_I4( &param2 ) = type;    
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_wrapper, L"showLevel", 2, param2, param1);
    if ( FAILED( hr ) ) {
        ERR("showLevel \n");
    }    
    
    VariantClear( &param1 );
    VariantClear( &param2 );
    VariantClear( &res );
 		
 	TRACE_OUT;	
    return ( hr );
}
