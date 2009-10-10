/*
 * implementation of OOPageStyle
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

#include "../OOWrappers/oo_page_style.h"


OOPageStyle::OOPageStyle()
{
    TRACE_IN;
                                    
    m_pd_page_style = NULL;                                   
    
    TRACE_OUT;                   
}

OOPageStyle::OOPageStyle(const OOPageStyle &obj)
{
   TRACE_IN;    
                               
   m_pd_page_style = obj.m_pd_page_style;
   if ( m_pd_page_style != NULL )
       m_pd_page_style->AddRef();  
       
   TRACE_OUT;                        
}
                       
OOPageStyle::~OOPageStyle()
{
   TRACE_IN;                    
                     
   if ( m_pd_page_style != NULL )
   {
       m_pd_page_style->Release();
       m_pd_page_style = NULL;        
   }
   
   TRACE_OUT;
}
   
OOPageStyle& OOPageStyle::operator=( const OOPageStyle &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_page_style != NULL )
   {
       m_pd_page_style->Release();
       m_pd_page_style = NULL;        
   } 
   
   m_pd_page_style = obj.m_pd_page_style;
   if ( m_pd_page_style != NULL )
       m_pd_page_style->AddRef();
   
   return ( *this );         
}
  
void OOPageStyle::Init( IDispatch* p_oo_page_style)
{
   TRACE_IN; 
     
   if ( m_pd_page_style != NULL )
   {
       m_pd_page_style->Release();
       m_pd_page_style = NULL;        
   } 
   
   if ( p_oo_page_style == NULL )
   {
       ERR( " p_oo_page_style == NULL \n" );
       return;     
   }
   
   m_pd_page_style = p_oo_page_style;
   m_pd_page_style->AddRef();
   
   TRACE_OUT;
   
   return;          
}
  
bool OOPageStyle::IsNull()
{
    return ( (m_pd_page_style == NULL) ? true : false );     
}
