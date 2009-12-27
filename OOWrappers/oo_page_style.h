/*
 * header file - OOPageStyle
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

#ifndef __UNIOFFICE_OO_WRAP_PAGE_STYLE_H__
#define __UNIOFFICE_OO_WRAP_PAGE_STYLE_H__

#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../Common/tools.h"

#include "./com/sun/star/uno/x_interface.h"

using namespace com::sun::star::uno;

class OOPageStyle: 
	  public virtual XInterface
{
public:

  OOPageStyle();
  virtual ~OOPageStyle(); 
  
  OOPageStyle& operator=( const XBase &obj);    
  
  double  LeftMargin( );
  HRESULT LeftMargin( double );
  
  double  RightMargin( );
  HRESULT RightMargin( double );  

  double  TopMargin( );
  HRESULT TopMargin( double );  

  double  BottomMargin( );
  HRESULT BottomMargin( double );

  VARIANT_BOOL IsLandscape( );
  HRESULT      IsLandscape( VARIANT_BOOL );

  short    PageScale();
  HRESULT  PageScale( short );

  short    ScaleToPagesY();
  HRESULT  ScaleToPagesY( short );

  short    ScaleToPagesX();
  HRESULT  ScaleToPagesX( short );

  double  HeaderHeight( );
  HRESULT HeaderHeight( double );

  double  FooterHeight( );
  HRESULT FooterHeight( double );

  VARIANT_BOOL CenterHorizontally( );
  HRESULT      CenterHorizontally( VARIANT_BOOL );

  VARIANT_BOOL CenterVertically( );
  HRESULT      CenterVertically( VARIANT_BOOL );
  
private:
           
};  

#endif // __UNIOFFICE_OO_WRAP_PAGE_STYLE_H__
