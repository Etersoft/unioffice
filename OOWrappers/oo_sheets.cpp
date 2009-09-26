/*
 * implementation of OOSheets
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

#include "../OOWrappers/oo_sheets.h"

OOSheets::OOSheets()
{
    TRACE_IN;
                                    
    m_pd_sheets = NULL;                                   
    
    TRACE_OUT;                    
}

OOSheets::OOSheets(const OOSheets &obj)
{
   TRACE_IN;
         
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   }        
                               
   m_pd_sheets = obj.m_pd_sheets;
   if ( m_pd_sheets != NULL )
       m_pd_sheets->AddRef();  
       
   TRACE_OUT;                         
}

OOSheets::~OOSheets()
{
   TRACE_IN;                    
                     
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   }
   
   TRACE_OUT;                        
}    
   
OOSheets& OOSheets::operator=( const OOSheets &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   } 
   
   m_pd_sheets = obj.m_pd_sheets;
   if ( m_pd_sheets != NULL )
       m_pd_sheets->AddRef();
   
   return ( *this );           
}
  
void OOSheets::Init( IDispatch* p_oo_sheets)
{
   TRACE_IN; 
     
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   } 
   
   if ( p_oo_sheets == NULL )
   {
       ERR( " p_oo_sheets == NULL \n" );
       return;     
   }
   
   m_pd_sheets = p_oo_sheets;
   m_pd_sheets->AddRef();
   
   TRACE_OUT;
   
   return;     
}
  
bool OOSheets::IsNull()
{
    return ( (m_pd_sheets == NULL) ? true : false );     
}
  
long OOSheets::getCount( )
{
   TRACE_NOTIMPL;
   return ( -1 );     
}
