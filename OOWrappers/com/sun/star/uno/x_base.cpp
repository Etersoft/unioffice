/*
 * source file - XBase
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

#include "x_base.h"
   
com::sun::star::uno::XBase::XBase( )
{                                   
    m_pd_wrapper = NULL;                                    
}

com::sun::star::uno::XBase::XBase(const XBase &obj)
{                     
   m_pd_wrapper = obj.m_pd_wrapper;
   if ( m_pd_wrapper != NULL )
       m_pd_wrapper->AddRef();    										
}

com::sun::star::uno::XBase::~XBase( )
{              
   if ( m_pd_wrapper != NULL )
   {
       m_pd_wrapper->Release();
       m_pd_wrapper = NULL;        
   }							
}   
   
com::sun::star::uno::XBase& com::sun::star::uno::XBase::operator=( const XBase &obj)
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
  
void com::sun::star::uno::XBase::Init( IDispatch* p_wrapper)
{
   if ( m_pd_wrapper != NULL )
   {
       m_pd_wrapper->Release();
       m_pd_wrapper = NULL;        
   } 
   
   if ( p_wrapper == NULL )
   {
       return;     
   }
   
   m_pd_wrapper = p_wrapper;
   m_pd_wrapper->AddRef(); 	 
}
  
bool com::sun::star::uno::XBase::IsNull()
{
    return ( (m_pd_wrapper == NULL) ? true : false ); 	 
}
