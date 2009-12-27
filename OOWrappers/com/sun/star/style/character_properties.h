/*
 * header file - CharacterProperties
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

#ifndef __UNIOFFICE_CHARACTER_PROPERTIES_H__
#define __UNIOFFICE_CHARACTER_PROPERTIES_H__

#include "../uno/x_base.h"
#include "../awt/font_slant.h"

namespace com
{
    namespace sun
    {
        namespace star
        {
		    namespace style 
			{ 
 		  
			class CharacterProperties: 
				  public virtual com::sun::star::uno::XBase 
			{
			public:
       
			  CharacterProperties( );
			  virtual ~CharacterProperties( );     
         
         	  HRESULT setCharFontName( BSTR );
         	  HRESULT getCharFontName( BSTR& );
         
         	  HRESULT setCharShadowed( bool );
         	  HRESULT getCharShadowed( bool& );
         	  
         	  HRESULT setCharWeight( float );
         	  HRESULT getCharWeight( float& );         	  
         	  
         	  HRESULT setCharPosture( com::sun::star::awt::FontSlant );
         	  HRESULT getCharPosture( com::sun::star::awt::FontSlant& );           	  
         	  
         	  HRESULT setCharColor( long );
         	  HRESULT getCharColor( long& );
         	  
         	  HRESULT setCharHeight( long );
         	  HRESULT getCharHeight( long& );
         	  
			protected:             
								  
			};

            } // namespace style
        } // namespace star
    } // namespace sun
} // namespace com

#endif // __UNIOFFICE_CHARACTER_PROPERTIES_H__
