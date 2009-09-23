/*
 * header file - WrapPropertyArray
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

#ifndef __UNIOFFICE_WRAP_PROPERTY_ARRAY_H__
#define __UNIOFFICE_WRAP_PROPERTY_ARRAY_H__

#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../OOWrappers/oo_property_value.h"
#include <vector>

using namespace std;

class WrapPropertyArray
{
public:
       
  WrapPropertyArray();
  WrapPropertyArray(const WrapPropertyArray &);
  virtual ~WrapPropertyArray(); 
 
  WrapPropertyArray& operator=(const WrapPropertyArray &); 
  
  SAFEARRAY FAR* Get_SafeArray();
  
  void Clear();
  void Add( OOPropertyValue );
        
private:      
  
  SAFEARRAY FAR*             m_sa_property_values;
  
  vector<OOPropertyValue>    m_properties;
  
  HRESULT Build_SafeArray();
      
};

#endif //__UNIOFFICE_WRAP_PROPERTY_ARRAY_H__
