/*
 * implementation of OONamedRange
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

#include "../OOWrappers/oo_named_range.h"


OONamedRange::OONamedRange()
{
    TRACE_IN;
                                    
    m_pd_named_range = NULL;                                   
    
    TRACE_OUT;                   
}

OONamedRange::OONamedRange(const OONamedRange &obj)
{
   TRACE_IN;    
                               
   m_pd_named_range = obj.m_pd_named_range;
   if ( m_pd_named_range != NULL )
       m_pd_named_range->AddRef();  
       
   TRACE_OUT;                        
}
                       
OONamedRange::~OONamedRange()
{
   TRACE_IN;                    
                     
   if ( m_pd_named_range != NULL )
   {
       m_pd_named_range->Release();
       m_pd_named_range = NULL;        
   }
   
   TRACE_OUT;
}
   
OONamedRange& OONamedRange::operator=( const OONamedRange &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_named_range != NULL )
   {
       m_pd_named_range->Release();
       m_pd_named_range = NULL;        
   } 
   
   m_pd_named_range = obj.m_pd_named_range;
   if ( m_pd_named_range != NULL )
       m_pd_named_range->AddRef();
   
   return ( *this );         
}
  
void OONamedRange::Init( IDispatch* p_oo_named_range)
{
   TRACE_IN; 
     
   if ( m_pd_named_range != NULL )
   {
       m_pd_named_range->Release();
       m_pd_named_range = NULL;        
   } 
   
   if ( p_oo_named_range == NULL )
   {
       ERR( " p_oo_named_range == NULL \n" );
       return;     
   }
   
   m_pd_named_range = p_oo_named_range;
   m_pd_named_range->AddRef();
   
   TRACE_OUT;
   
   return;          
}
  
bool OONamedRange::IsNull()
{
    return ( (m_pd_named_range == NULL) ? true : false );     
}
