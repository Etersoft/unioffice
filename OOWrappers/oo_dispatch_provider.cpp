/*
 * implementation of OODispatchProvider
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

#include "../OOWrappers/oo_dispatch_provider.h"

OODispatchProvider::OODispatchProvider()
{
    TRACE_IN;
                                    
    m_pd_dispatch_provider = NULL;                                   
    
    TRACE_OUT;                    
}

OODispatchProvider::OODispatchProvider(const OODispatchProvider &obj)
{
   TRACE_IN;      
                               
   m_pd_dispatch_provider = obj.m_pd_dispatch_provider;
   if ( m_pd_dispatch_provider != NULL )
       m_pd_dispatch_provider->AddRef();  
       
   TRACE_OUT;                         
}

OODispatchProvider::~OODispatchProvider()
{
   TRACE_IN;                    
                     
   if ( m_pd_dispatch_provider != NULL )
   {
       m_pd_dispatch_provider->Release();
       m_pd_dispatch_provider = NULL;        
   }
   
   TRACE_OUT;                        
}    
   
OODispatchProvider& OODispatchProvider::operator=( const OODispatchProvider &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_dispatch_provider != NULL )
   {
       m_pd_dispatch_provider->Release();
       m_pd_dispatch_provider = NULL;        
   } 
   
   m_pd_dispatch_provider = obj.m_pd_dispatch_provider;
   if ( m_pd_dispatch_provider != NULL )
       m_pd_dispatch_provider->AddRef();
   
   return ( *this );           
}
  
void OODispatchProvider::Init( IDispatch* p_oo_dispatch_provider)
{
   TRACE_IN; 
     
   if ( m_pd_dispatch_provider != NULL )
   {
       m_pd_dispatch_provider->Release();
       m_pd_dispatch_provider = NULL;        
   } 
   
   if ( p_oo_dispatch_provider == NULL )
   {
       ERR( " p_oo_dispatch_provider == NULL \n" );
       return;     
   }
   
   m_pd_dispatch_provider = p_oo_dispatch_provider;
   m_pd_dispatch_provider->AddRef();
   
   TRACE_OUT;
   
   return;     
}
  
bool OODispatchProvider::IsNull()
{
    return ( (m_pd_dispatch_provider == NULL) ? true : false );     
}

IDispatch* OODispatchProvider::GetIDispatch()
{
    TRACE_IN;
	
	if ( !IsNull() )
	{
	    m_pd_dispatch_provider->AddRef();       	 
    }
	
	TRACE_OUT;
	return ( m_pd_dispatch_provider ); 		   
}
