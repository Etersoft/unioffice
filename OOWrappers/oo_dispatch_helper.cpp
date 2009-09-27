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

#include "../OOWrappers/oo_dispatch_helper.h"

OODispatchHelper::OODispatchHelper()
{
    TRACE_IN;
                                    
    m_pd_dispatch_helper = NULL;                                   
    
    TRACE_OUT;                    
}

OODispatchHelper::OODispatchHelper(const OODispatchHelper &obj)
{
   TRACE_IN;      
                               
   m_pd_dispatch_helper = obj.m_pd_dispatch_helper;
   if ( m_pd_dispatch_helper != NULL )
       m_pd_dispatch_helper->AddRef();  
       
   TRACE_OUT;                         
}

OODispatchHelper::~OODispatchHelper()
{
   TRACE_IN;                    
                     
   if ( m_pd_dispatch_helper != NULL )
   {
       m_pd_dispatch_helper->Release();
       m_pd_dispatch_helper = NULL;        
   }
   
   TRACE_OUT;                        
}    
   
OODispatchHelper& OODispatchHelper::operator=( const OODispatchHelper &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_dispatch_helper != NULL )
   {
       m_pd_dispatch_helper->Release();
       m_pd_dispatch_helper = NULL;        
   } 
   
   m_pd_dispatch_helper = obj.m_pd_dispatch_helper;
   if ( m_pd_dispatch_helper != NULL )
       m_pd_dispatch_helper->AddRef();
   
   return ( *this );           
}
  
void OODispatchHelper::Init( IDispatch* p_oo_dispatch_helper)
{
   TRACE_IN; 
     
   if ( m_pd_dispatch_helper != NULL )
   {
       m_pd_dispatch_helper->Release();
       m_pd_dispatch_helper = NULL;        
   } 
   
   if ( p_oo_dispatch_helper == NULL )
   {
       ERR( " p_oo_dispatch_helper == NULL \n" );
       return;     
   }
   
   m_pd_dispatch_helper = p_oo_dispatch_helper;
   m_pd_dispatch_helper->AddRef();
   
   TRACE_OUT;
   
   return;     
}
  
bool OODispatchHelper::IsNull()
{
    return ( (m_pd_dispatch_helper == NULL) ? true : false );     
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
		 m_pd_dispatch_helper, 
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
