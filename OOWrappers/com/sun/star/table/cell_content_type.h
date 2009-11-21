/*
 * header file - XCell
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

#ifndef __UNIOFFICE_CELL_CONTENT_TYPE_H__
#define __UNIOFFICE_CELL_CONTENT_TYPE_H__

namespace com
{
    namespace sun
    {
        namespace star
        {
		    namespace table 
			{ 
 		  
			enum CellContentType
			{
			 	 EMPTY,
				 VALUE,
				 TEXT,
				 FORMULA	  
			};

            } // namespace table
        } // namespace star
    } // namespace sun
} // namespace com

#endif // __UNIOFFICE_CELL_CONTENT_TYPE_H__
