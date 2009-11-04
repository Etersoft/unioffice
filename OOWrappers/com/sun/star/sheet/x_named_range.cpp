/*
 * source file - XNamedRange
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

#include "x_named_range.h"
   
com::sun::star::sheet::XNamedRange::XNamedRange( ):com::sun::star::container::XNamed()
{                                                                       
}

com::sun::star::sheet::XNamedRange::~XNamedRange( )
{              							
}

BSTR com::sun::star::sheet::XNamedRange::getContent()
{
 	 TRACE_NOTIMPL;	 
}

HRESULT com::sun::star::sheet::XNamedRange::setContent( BSTR _content)
{
 	 TRACE_NOTIMPL;	 		
}
			  
com::sun::star::table::SCellAddress com::sun::star::sheet::XNamedRange::getReferencePosition()
{
 	 TRACE_NOTIMPL;	 									
}

HRESULT com::sun::star::sheet::XNamedRange::setReferencePosition( com::sun::star::table::SCellAddress _cell_address)
{
 	 TRACE_NOTIMPL;	 		
}
			  
long com::sun::star::sheet::XNamedRange::getType()
{
 	 TRACE_NOTIMPL;	 	 
}

HRESULT com::sun::star::sheet::XNamedRange::setType( long _type )
{
 	 TRACE_NOTIMPL;	 		
}
