/*
 * implementation of OOPageStyles
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

#include "../OOWrappers/oo_page_styles.h"


OOPageStyles::OOPageStyles()
{
    TRACE_IN;
                                    
    m_pd_page_styles = NULL;                                   
    
    TRACE_OUT;                   
}

OOPageStyles::OOPageStyles(const OOPageStyles &obj)
{
   TRACE_IN;    
                               
   m_pd_page_styles = obj.m_pd_page_styles;
   if ( m_pd_page_styles != NULL )
       m_pd_page_styles->AddRef();  
       
   TRACE_OUT;                        
}
                       
OOPageStyles::~OOPageStyles()
{
   TRACE_IN;                    
                     
   if ( m_pd_page_styles != NULL )
   {
       m_pd_page_styles->Release();
       m_pd_page_styles = NULL;        
   }
   
   TRACE_OUT;
}
   
OOPageStyles& OOPageStyles::operator=( const OOPageStyles &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_page_styles != NULL )
   {
       m_pd_page_styles->Release();
       m_pd_page_styles = NULL;        
   } 
   
   m_pd_page_styles = obj.m_pd_page_styles;
   if ( m_pd_page_styles != NULL )
       m_pd_page_styles->AddRef();
   
   return ( *this );         
}
  
void OOPageStyles::Init( IDispatch* p_oo_page_styles)
{
   TRACE_IN; 
     
   if ( m_pd_page_styles != NULL )
   {
       m_pd_page_styles->Release();
       m_pd_page_styles = NULL;        
   } 
   
   if ( p_oo_page_styles == NULL )
   {
       ERR( " p_oo_page_styles == NULL \n" );
       return;     
   }
   
   m_pd_page_styles = p_oo_page_styles;
   m_pd_page_styles->AddRef();
   
   TRACE_OUT;
   
   return;          
}
  
bool OOPageStyles::IsNull()
{
    return ( (m_pd_page_styles == NULL) ? true : false );     
}

HRESULT OOPageStyles::getByName( BSTR _name_of_style, OOPageStyle& oo_page_style )
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_disp;
    VARIANT res, param1;
     
    VariantInit(&res);
    VariantInit(&param1);
    
    if ( IsNull() )
    {
        ERR( " IsNull() == true \n" );
        TRACE_OUT;
        return ( E_FAIL );      
    } 
	
	V_VT( &param1 ) = VT_BSTR;
	V_BSTR( &param1 ) = SysAllocString( _name_of_style );
	
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_page_styles, L"getByName", 1, param1);
    
    p_disp = V_DISPATCH( &res );
    	
    if ( FAILED( hr ) ) {
        ERR(" getByName \n ");
        TRACE_OUT;
        return ( hr );
    }
    
    if ( p_disp == NULL )
    {
	    ERR( " p_disp == NULL \n" );
		TRACE_OUT;   	 
	    return ( E_FAIL );
    }
    
    oo_page_style.Init( p_disp );
    
    VariantClear( &res ); 
    VariantClear( &param1 );
    
    TRACE_OUT;
    return ( hr ); 	 		
}
