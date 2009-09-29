/*
 * implementation of OOFrame
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

#include "../OOWrappers/oo_frame.h"

OOFrame::OOFrame()
{
    TRACE_IN;
                                    
    m_pd_frame = NULL;                                   
    
    TRACE_OUT;                    
}

OOFrame::OOFrame(const OOFrame &obj)
{
   TRACE_IN;      
                               
   m_pd_frame = obj.m_pd_frame;
   if ( m_pd_frame != NULL )
       m_pd_frame->AddRef();  
       
   TRACE_OUT;                         
}

OOFrame::~OOFrame()
{
   TRACE_IN;                    
                     
   if ( m_pd_frame != NULL )
   {
       m_pd_frame->Release();
       m_pd_frame = NULL;        
   }
   
   TRACE_OUT;                        
}    
   
OOFrame& OOFrame::operator=( const OOFrame &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_frame != NULL )
   {
       m_pd_frame->Release();
       m_pd_frame = NULL;        
   } 
   
   m_pd_frame = obj.m_pd_frame;
   if ( m_pd_frame != NULL )
       m_pd_frame->AddRef();
   
   return ( *this );           
}
  
void OOFrame::Init( IDispatch* p_oo_frame)
{
   TRACE_IN; 
     
   if ( m_pd_frame != NULL )
   {
       m_pd_frame->Release();
       m_pd_frame = NULL;        
   } 
   
   if ( p_oo_frame == NULL )
   {
       ERR( " p_oo_frame == NULL \n" );
       return;     
   }
   
   m_pd_frame = p_oo_frame;
   m_pd_frame->AddRef();
   
   TRACE_OUT;
   
   return;     
}
  
bool OOFrame::IsNull()
{
    return ( (m_pd_frame == NULL) ? true : false );     
}

OODispatchProvider OOFrame::GetDispatchProvider( )
{
 	TRACE_IN;
	OODispatchProvider ret_val;
	
	if ( IsNull() )
	{
	    ERR( " m_pd_frame is Null \n" );   	 
    }
	
	ret_val.Init( m_pd_frame );
	 			   
 	TRACE_OUT;			   
    return ( ret_val ); 				   
}
