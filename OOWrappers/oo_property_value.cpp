/*
 * implementation of OOPropertyValue
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

#include "./oo_property_value.h"

using namespace com::sun::star::uno;

OOPropertyValue::OOPropertyValue():XBase()
{                        
}

OOPropertyValue::~OOPropertyValue()
{                       
}

IDispatch* OOPropertyValue::GetOOProperty()
{
    if ( m_pd_wrapper != NULL )
    {
       m_pd_wrapper->AddRef();     
    } else
    {
        ERR( " m_pd_wrapper is NULL \n" );      
    }           
          
    return ( m_pd_wrapper );           
}

HRESULT OOPropertyValue::Set_PropertyName( BSTR _name )
{
    HRESULT hr = S_OK;
    VARIANT param;
    VARIANT res;
    
    TRACE_IN;

    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
		TRACE_OUT; 
        return ( E_FAIL );     
    }
    
    VariantInit(&param);
    VariantInit(&res);
    
    V_VT(&param) = VT_BSTR;
    V_BSTR(&param) = SysAllocString(_name);
    
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"Name", 1, param);
    
    if ( FAILED( hr ) )
    {
        ERR( " \n " );     
    }
    
    VariantClear(&param);
    VariantClear(&res);
    
    TRACE_OUT;
    
    return ( hr );        
}

HRESULT OOPropertyValue::Set_PropertyValue( BSTR _value )
{
    HRESULT hr = S_OK;
    VARIANT param;
    VARIANT res;
    
    TRACE_IN;

    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
		TRACE_OUT; 
        return ( E_FAIL );     
    }
    
    VariantInit(&param);
    VariantInit(&res);
    
    V_VT(&param) = VT_BSTR;
    V_BSTR(&param) = SysAllocString(_value);
    
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"Value", 1, param);
    
    if ( FAILED( hr ) )
    {
        ERR( " \n " );     
    }
    
    VariantClear(&param);
    VariantClear(&res);
    
    TRACE_OUT;
    
    return ( hr );         
}
 
HRESULT OOPropertyValue::Set_PropertyValue( VARIANT_BOOL _value)
{
    HRESULT hr = S_OK;
    VARIANT param;
    VARIANT res;
    
    TRACE_IN;
    
    if ( IsNull() )
    {
        ERR( " m_pd_wrapper is NULL \n" );
		TRACE_OUT; 
        return ( E_FAIL );     
    }
    
    VariantInit(&param);
    VariantInit(&res);
    
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = _value;
    
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_wrapper, L"Value", 1, param);
    
    if ( FAILED( hr ) )
    {
        ERR( " \n " );     
    }
    
    VariantClear(&param);
    VariantClear(&res);
    
    TRACE_OUT;
    
    return ( hr );   
      
} 
        
HRESULT OOPropertyValue::Set_Property( BSTR _name, BSTR _value )
{
    TRACE_IN;
    HRESULT hr = S_OK;
    
    hr = Set_PropertyName( _name );
    
    if ( FAILED( hr ) )
    {
        ERR( " Set_PropertyName() \n " );     
    }
    
    hr = Set_PropertyValue( _value );
    
    if ( FAILED( hr ) )
    {
        ERR( " Set_PropertyValue() \n " );     
    }
    
    TRACE_OUT;
    return ( hr );         
}
