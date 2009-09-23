/*
 * implementation of OOServiceManager
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

#include "../OOWrappers/oo_servicemanager.h"


OOServiceManager::OOServiceManager()
{
    CLSID clsid;                                
    HRESULT hr;
    
    TRACE_IN;
                                    
    m_pd_servicemanager = NULL;                                   
    
    hr = CLSIDFromProgID(L"com.sun.star.ServiceManager", &clsid);
    if (FAILED(hr)) {
        ERR(" CLSIDFromProgID  com.sun.star.ServiceManager \n");
        return;
    }

    /* Start server and get IDispatch...*/
    hr = CoCreateInstance( clsid, NULL, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER, IID_IDispatch, (void**) &m_pd_servicemanager);
    if (FAILED(hr)) {
        ERR(" CoCreateInstance \n");
        return;
    }
    
    TRACE_OUT;
}

OOServiceManager::~OOServiceManager()
{
   TRACE_IN;
   
   if ( m_pd_servicemanager != NULL )
   {
       m_pd_servicemanager->Release();
       m_pd_servicemanager = NULL;        
   }                                  
   
   TRACE_OUT;
}
  
OODesktop OOServiceManager::Get_Desktop( )
{
    OODesktop   ret_val;
    IDispatch*  p_disp = NULL;
    
    TRACE_IN;
    
    p_disp = CreateInstance( SysAllocString( L"com.sun.star.frame.Desktop" ) );
    
    if ( p_disp == NULL )
    {
        ERR( " p_disp == NULL \n" );     
    }
    
    ret_val.Init( p_disp );
    
    p_disp->Release();
    
    TRACE_OUT;
    
    return ( ret_val );   
}

IDispatch* OOServiceManager::CreateInstance( BSTR str_value )
{
    VARIANT     param1, result;
    HRESULT     hr;
    IDispatch*  p_disp = NULL;
    
    TRACE_IN;
    
    if ( m_pd_servicemanager == NULL )
    {
        ERR( " m_pd_servicemanager is NULL \n" ); 
        return ( NULL );     
    }
    
    VariantInit( &param1 );
    VariantInit( &result );
    
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = str_value;
    
    /* Get Desktop and its assoc. IDispatch...*/
    hr = AutoWrap(DISPATCH_METHOD, &result, m_pd_servicemanager, L"CreateInstance", 1, param1);
    
    if (FAILED(hr)) {
        ERR(" CreateInstance \n");
        return ( NULL );
    }
    
    p_disp = result.pdispVal;
    if ( p_disp == NULL )
    {
        ERR( "p_disp == NULL \n" ); 
        return ( NULL );      
    }
    p_disp->AddRef();
    
    VariantClear( &param1 );
    VariantClear( &result );
    
    TRACE_OUT;
    
    return ( p_disp );   
}

IDispatch* OOServiceManager::Bridge_GetStruct( BSTR str_value )
{
    VARIANT     param1, result;
    HRESULT     hr;
    IDispatch*  p_disp = NULL;
    
    TRACE_IN;
    
    if ( m_pd_servicemanager == NULL )
    {
        ERR( " m_pd_servicemanager is NULL \n" ); 
        return ( NULL );     
    }
    
    VariantInit( &param1 );
    VariantInit( &result );
    
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = str_value;
    
    /* Get Desktop and its assoc. IDispatch...*/
    hr = AutoWrap(DISPATCH_METHOD, &result, m_pd_servicemanager, L"Bridge_GetStruct", 1, param1);
    
    if (FAILED(hr)) {
        ERR(" CreateInstance \n");
        return ( NULL );
    }
    
    p_disp = result.pdispVal;
    if ( p_disp == NULL )
    {
        ERR( "p_disp == NULL \n" );   
        return ( NULL );  
    }
    p_disp->AddRef();
    
    VariantClear( &param1 );
    VariantClear( &result );
    
    TRACE_OUT;
    
    return ( p_disp );   
}

OOPropertyValue OOServiceManager::Get_PropertyValue( )
{
    OOPropertyValue   ret_val;
    IDispatch*        p_disp = NULL;
    
    TRACE_IN;
    
    p_disp = Bridge_GetStruct( SysAllocString( L"com.sun.star.beans.PropertyValue" ) );
    
    if ( p_disp == NULL )
    {
        ERR( " p_disp == NULL \n" );     
    }
    
    ret_val.Init( p_disp );
    
    p_disp->Release();
    
    TRACE_OUT;
    
    return ( ret_val );   
}
