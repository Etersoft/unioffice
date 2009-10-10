/*
 * implementation of OOStyleFamilies
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

#include "../OOWrappers/oo_style_families.h"


OOStyleFamilies::OOStyleFamilies()
{
    TRACE_IN;
                                    
    m_pd_style_families = NULL;                                   
    
    TRACE_OUT;                   
}

OOStyleFamilies::OOStyleFamilies(const OOStyleFamilies &obj)
{
   TRACE_IN;    
                               
   m_pd_style_families = obj.m_pd_style_families;
   if ( m_pd_style_families != NULL )
       m_pd_style_families->AddRef();  
       
   TRACE_OUT;                        
}
                       
OOStyleFamilies::~OOStyleFamilies()
{
   TRACE_IN;                    
                     
   if ( m_pd_style_families != NULL )
   {
       m_pd_style_families->Release();
       m_pd_style_families = NULL;        
   }
   
   TRACE_OUT;
}
   
OOStyleFamilies& OOStyleFamilies::operator=( const OOStyleFamilies &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_style_families != NULL )
   {
       m_pd_style_families->Release();
       m_pd_style_families = NULL;        
   } 
   
   m_pd_style_families = obj.m_pd_style_families;
   if ( m_pd_style_families != NULL )
       m_pd_style_families->AddRef();
   
   return ( *this );         
}
  
void OOStyleFamilies::Init( IDispatch* p_oo_style_families)
{
   TRACE_IN; 
     
   if ( m_pd_style_families != NULL )
   {
       m_pd_style_families->Release();
       m_pd_style_families = NULL;        
   } 
   
   if ( p_oo_style_families == NULL )
   {
       ERR( " p_oo_style_families == NULL \n" );
       return;     
   }
   
   m_pd_style_families = p_oo_style_families;
   m_pd_style_families->AddRef();
   
   TRACE_OUT;
   
   return;          
}
  
bool OOStyleFamilies::IsNull()
{
    return ( (m_pd_style_families == NULL) ? true : false );     
}

HRESULT OOStyleFamilies::getPageStyles( OOPageStyles& oo_page_styles )
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
	V_BSTR( &param1 ) = SysAllocString( L"PageStyles");
	
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_style_families, L"getByName", 1, param1);
    
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
    
    oo_page_styles.Init( p_disp );
    
    VariantClear( &res ); 
    VariantClear( &param1 );
    
    TRACE_OUT;
    return ( hr ); 		
}

