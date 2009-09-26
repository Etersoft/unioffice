/*
 * implementation of OOSheet
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

#include "../OOWrappers/oo_sheet.h"


OOSheet::OOSheet()
{
    TRACE_IN;
                                    
    m_pd_sheet = NULL;                                   
    
    TRACE_OUT;                   
}

OOSheet::OOSheet(const OOSheet &obj)
{
   TRACE_IN;
         
   if ( m_pd_sheet != NULL )
   {
       m_pd_sheet->Release();
       m_pd_sheet = NULL;        
   }        
                               
   m_pd_sheet = obj.m_pd_sheet;
   if ( m_pd_sheet != NULL )
       m_pd_sheet->AddRef();  
       
   TRACE_OUT;                        
}
                       
OOSheet::~OOSheet()
{
   TRACE_IN;                    
                     
   if ( m_pd_sheet != NULL )
   {
       m_pd_sheet->Release();
       m_pd_sheet = NULL;        
   }
   
   TRACE_OUT;
}
   
OOSheet& OOSheet::operator=( const OOSheet &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_sheet != NULL )
   {
       m_pd_sheet->Release();
       m_pd_sheet = NULL;        
   } 
   
   m_pd_sheet = obj.m_pd_sheet;
   if ( m_pd_sheet != NULL )
       m_pd_sheet->AddRef();
   
   return ( *this );         
}
  
void OOSheet::Init( IDispatch* p_oo_sheet)
{
   TRACE_IN; 
     
   if ( m_pd_sheet != NULL )
   {
       m_pd_sheet->Release();
       m_pd_sheet = NULL;        
   } 
   
   if ( p_oo_sheet == NULL )
   {
       ERR( " p_oo_sheet == NULL \n" );
       return;     
   }
   
   m_pd_sheet = p_oo_sheet;
   m_pd_sheet->AddRef();
   
   TRACE_OUT;
   
   return;          
}
  
bool OOSheet::IsNull()
{
    return ( (m_pd_sheet == NULL) ? true : false );     
}

BSTR OOSheet::getName( )
{
    TRACE_IN;
    HRESULT hr;
    VARIANT res;
    BSTR result;
    
    VariantInit( &res );
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_sheet, L"getName", 0);
    if ( FAILED( hr ) )
    {
        ERR( " getName \n" );     
        result = SysAllocString( L"" );
    } else
    {
        result = SysAllocString( V_BSTR( &res ) );      
    }
    
    VariantClear( &res );
    
    TRACE_OUT;     
    return ( result );
}

HRESULT OOSheet::setName( BSTR bstr_name )
{
    TRACE_IN;
    
    HRESULT hr;
    VARIANT param1, res;
    
    VariantInit( &param1 );
    VariantInit( &res );  
        
    V_VT(&param1)   = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(bstr_name);

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_sheet, L"setName", 1, param1);
    
    if ( FAILED( hr ) )
    {
        ERR( " setName \n" );     
    }    
    
    VariantClear( &res );
    VariantClear(&param1 );
    
    TRACE_OUT;
    return ( hr );      
}