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

#include "./oo_frame.h"

using namespace com::sun::star::uno;

OOFrame::OOFrame():XBase()
{                    
}

OOFrame::~OOFrame()
{                       
}    

OODispatchProvider OOFrame::GetDispatchProvider( )
{
 	TRACE_IN;
	OODispatchProvider ret_val;
	
	if ( IsNull() )
	{
	    ERR( " m_pd_frame is Null \n" );   	 
    }
	
	ret_val.Init( m_pd_wrapper );
	 			   
 	TRACE_OUT;			   
    return ( ret_val ); 				   
}
