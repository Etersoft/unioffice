/*
 * header file - SheetCellRange
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

#ifndef __UNIOFFICE_SHEET_CELL_RANGE_H__
#define __UNIOFFICE_SHEET_CELL_RANGE_H__

#include "../uno/x_base.h"
#include "../table/cell_range.h"
#include "../table/x_column_row_range.h"
#include "../sheet/x_cell_range_addressable.h"
#include "../style/character_properties.h"
#include "../sheet/x_sheet_operation.h"
#include "../util/x_mergeable.h"
#include "../awt/size.h"


namespace com
{
    namespace sun
    {
        namespace star
        {
		    namespace sheet 
			{ 
 		  
			class SheetCellRange: 
				  public virtual com::sun::star::table::CellRange,
				  public virtual com::sun::star::table::XColumnRowRange,
				  public virtual com::sun::star::sheet::XCellRangeAddressable,
				  public virtual com::sun::star::style::CharacterProperties,
				  public virtual com::sun::star::sheet::XSheetOperation,
				  public virtual com::sun::star::util::XMergeable
			{
			public:
       
			  SheetCellRange( );
			  virtual ~SheetCellRange( );     

              com::sun::star::awt::Size getSize();
              

			protected:            
								  
			};

            } // namespace sheet
        } // namespace star
    } // namespace sun
} // namespace com

#endif // __UNIOFFICE_SHEET_CELL_RANGE_H__
