/*
 * header file - OONamedRange
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

#ifndef __UNIOFFICE_OO_WRAP_NAMED_RANGE_H__
#define __UNIOFFICE_OO_WRAP_NAMED_RANGE_H__

#include "../uno/x_interface.h"
#include "x_named_range.h"
#include "x_cell_range_referrer.h"

using namespace com::sun::star::uno;
using namespace com::sun::star::sheet;

class OONamedRange: 
	  public XInterface,
	  public XNamedRange,
	  public XCellRangeReferrer
{
public:

  OONamedRange();
  virtual ~OONamedRange();  
  
  OONamedRange& operator=( const XBase &obj);  

/* 
Properties' Summary
[ readonly ] long TokenIndex [ OPTIONAL ]
returns the index used to refer to this name in token arrays.  
boolean IsSharedFormula	[ OPTIONAL ]
Determines if this defined name represents a shared formula.  
*/ 

private:
		         
};  





#endif //__UNIOFFICE_OO_WRAP_NAMED_RANGE_H__
