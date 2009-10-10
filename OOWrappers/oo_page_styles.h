/*
 * header file - OOPageStyles
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

#ifndef __UNIOFFICE_OO_WRAP_PAGE_STYLES_H__
#define __UNIOFFICE_OO_WRAP_PAGE_STYLES_H__

#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../Common/tools.h"
#include "../OOWrappers/oo_page_style.h"

class OOPageStyles
{
public:

  OOPageStyles();
  OOPageStyles(const OOPageStyles &);
  virtual ~OOPageStyles();     
   
  OOPageStyles& operator=( const OOPageStyles &); 
  
  void Init( IDispatch* );
  
  bool IsNull();
  
  HRESULT getByName( BSTR, OOPageStyle& );
  
private:
  
   IDispatch*   m_pd_page_styles;
           
};  

#endif // __UNIOFFICE_OO_WRAP_PAGE_STYLES_H__
