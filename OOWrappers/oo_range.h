/*
 * header file - OORange
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

#ifndef __UNIOFFICE_OO_WRAP_RANGE_H__
#define __UNIOFFICE_OO_WRAP_RANGE_H__

#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../Common/tools.h"

#include "./com/sun/star/uno/x_interface.h"
#include "./com/sun/star/sheet/sheet_cell.h"
#include "./com/sun/star/sheet/sheet_cell_range.h"
#include "./com/sun/star/table/x_cell_range.h"
#include "./com/sun/star/table/table_column.h"
#include "./com/sun/star/table/table_row.h"
#include "./com/sun/star/table/x_table_columns.h"
#include "./com/sun/star/table/x_table_rows.h"

using namespace com::sun::star::uno;
using namespace com::sun::star::sheet;
using namespace com::sun::star::table;

class OORange: 
	  public virtual XInterface,
	  public virtual SheetCell,
	  public virtual SheetCellRange, 
	  
	  // special services - to do actions with columns and rows
	  public virtual XTableColumns,
	  public virtual XTableRows,
	  public virtual TableColumn,
	  public virtual TableRow
{
public:

  OORange();
  virtual ~OORange();     
   
  OORange& operator=( const OORange &obj); 
  OORange& operator=( const XCellRange &obj);
  OORange& operator=( const XCell &obj);
  OORange& operator=( const XTableColumns &obj);
  OORange& operator=( const XTableRows &obj);
   
private:
           
};          

#endif //__UNIOFFICE_OO_WRAP_RANGE_H__
