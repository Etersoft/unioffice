/*
 * header file - XCellRangeReferrer
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

#ifndef __UNIOFFICE_X_CELL_RANGE_REFERRER_H__
#define __UNIOFFICE_X_CELL_RANGE_REFERRER_H__

#include "../uno/x_base.h"
#include "../table/x_cell_range.h"

namespace com
{
    namespace sun
    {
        namespace star
        {
		    namespace sheet 
			{ 
 		  
			class XCellRangeReferrer: public virtual com::sun::star::uno::XBase 
			{
			public:
       
			  XCellRangeReferrer( );
			  virtual ~XCellRangeReferrer( );   
         
		      com::sun::star::table::XCellRange getReferredCells();
         
			protected:            
								  
			};

            } // namespace sheet
        } // namespace star
    } // namespace sun
} // namespace com

#endif // __UNIOFFICE_X_CELL_RANGE_REFERRER_H__
