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
