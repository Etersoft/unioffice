/*
 * source file - Point
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

#include "point.h"
   
com::sun::star::awt::Point::Point( ):com::sun::star::uno::XBase()
{
 																                                                                        
}

com::sun::star::awt::Point::~Point( )
{   
	           							
} 

com::sun::star::awt::Point& com::sun::star::awt::Point::operator=( const com::sun::star::awt::Point &obj)
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

HRESULT com::sun::star::awt::Point::setX( long value)
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
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"X", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call X \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr ); 		
}

HRESULT com::sun::star::awt::Point::getX( long& value)
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
    
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"X", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call X \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
 	
 	value = V_I4( &res );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr ); 		
}
				
HRESULT com::sun::star::awt::Point::setY( long value)
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
    
    hr = AutoWrap (DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"Y", 1, param1);
    if ( FAILED( hr ) )
    {
        ERR( " Call Y \n" );
    }
 
    VariantClear( &res );
    VariantClear( &param1 );
 
    TRACE_OUT;
    return ( hr );  		
}

HRESULT com::sun::star::awt::Point::getY( long& value)
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
    
    hr = AutoWrap(DISPATCH_PROPERTYGET, &res, m_pd_wrapper, L"Y", 0 );
    if ( FAILED( hr ) )
    {
        ERR( " Call Y \n" );
        TRACE_OUT;
        return ( E_FAIL );
    }
 	
 	VariantChangeTypeEx(&res, &res, 0, 0, VT_I4);
 	
 	value = V_I4( &res );
 
    VariantClear( &res );
 
    TRACE_OUT;
    return ( hr );  		
}
