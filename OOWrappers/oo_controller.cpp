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

#include "../OOWrappers/oo_controller.h"

OOController::OOController()
{
    TRACE_IN;
                                    
    m_pd_controller = NULL;                                   
    
    TRACE_OUT;                    
}

OOController::OOController(const OOController &obj)
{
   TRACE_IN;      
                               
   m_pd_controller = obj.m_pd_controller;
   if ( m_pd_controller != NULL )
       m_pd_controller->AddRef();  
       
   TRACE_OUT;                         
}

OOController::~OOController()
{
   TRACE_IN;                    
                     
   if ( m_pd_controller != NULL )
   {
       m_pd_controller->Release();
       m_pd_controller = NULL;        
   }
   
   TRACE_OUT;                        
}    
   
OOController& OOController::operator=( const OOController &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_controller != NULL )
   {
       m_pd_controller->Release();
       m_pd_controller = NULL;        
   } 
   
   m_pd_controller = obj.m_pd_controller;
   if ( m_pd_controller != NULL )
       m_pd_controller->AddRef();
   
   return ( *this );           
}
  
void OOController::Init( IDispatch* p_oo_controller)
{
   TRACE_IN; 
     
   if ( m_pd_controller != NULL )
   {
       m_pd_controller->Release();
       m_pd_controller = NULL;        
   } 
   
   if ( p_oo_controller == NULL )
   {
       ERR( " p_oo_controller == NULL \n" );
       return;     
   }
   
   m_pd_controller = p_oo_controller;
   m_pd_controller->AddRef();
   
   TRACE_OUT;
   
   return;     
}
  
bool OOController::IsNull()
{
    return ( (m_pd_controller == NULL) ? true : false );     
}
