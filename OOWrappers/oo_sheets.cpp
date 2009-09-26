/*
 * implementation of OOSheets
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

#include "../OOWrappers/oo_sheets.h"

OOSheets::OOSheets()
{
    TRACE_IN;
                                    
    m_pd_sheets = NULL;                                   
    
    TRACE_OUT;                    
}

OOSheets::OOSheets(const OOSheets &obj)
{
   TRACE_IN;
         
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   }        
                               
   m_pd_sheets = obj.m_pd_sheets;
   if ( m_pd_sheets != NULL )
       m_pd_sheets->AddRef();  
       
   TRACE_OUT;                         
}

OOSheets::~OOSheets()
{
   TRACE_IN;                    
                     
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   }
   
   TRACE_OUT;                        
}    
   
OOSheets& OOSheets::operator=( const OOSheets &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   } 
   
   m_pd_sheets = obj.m_pd_sheets;
   if ( m_pd_sheets != NULL )
       m_pd_sheets->AddRef();
   
   return ( *this );           
}
  
void OOSheets::Init( IDispatch* p_oo_sheets)
{
   TRACE_IN; 
     
   if ( m_pd_sheets != NULL )
   {
       m_pd_sheets->Release();
       m_pd_sheets = NULL;        
   } 
   
   if ( p_oo_sheets == NULL )
   {
       ERR( " p_oo_sheets == NULL \n" );
       return;     
   }
   
   m_pd_sheets = p_oo_sheets;
   m_pd_sheets->AddRef();
   
   TRACE_OUT;
   
   return;     
}
  
bool OOSheets::IsNull()
{
    return ( (m_pd_sheets == NULL) ? true : false );     
}
  
long OOSheets::getCount( )
{
   TRACE_IN;
   VARIANT res;
   HRESULT hr;
   long count = -1;
   
   VariantInit( &res );
   
   if ( IsNull() )
   {
       ERR( " IsNull() == true \n" );
       VariantClear( &res );
       return ( -1 );     
   }
   
   hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_sheets, L"getCount", 0);
   
   if ( FAILED( hr ) ) 
   {
       ERR(" getCount \n");
       count = -1;
   } else
   {
       count = V_I4( &res );        
   }
   
   VariantClear( &res );
   
   TRACE_OUT;
   return ( count );     
}

OOSheet OOSheets::getByIndex( long _index )
{
    TRACE_IN; 
    OOSheet oo_sheet;
    HRESULT hr;
    VARIANT res, var_index;
    
    VariantInit( &res );
    VariantInit( &var_index );
    
    V_VT( &res ) = VT_I4;
    V_I4( &var_index ) = _index;
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_sheets, L"getByIndex", 1, var_index);
    if ( FAILED( hr ) )
    {
        ERR( " getByIndex \n" );
    } else
    {
        oo_sheet.Init( V_DISPATCH( &res ) );      
    }
    
 
    VariantClear( &res );
    VariantClear( &var_index );
 
    TRACE_OUT;
    return oo_sheet;    
}

OOSheet OOSheets::getByName( BSTR _sheet_name )
{
    TRACE_IN; 
    OOSheet oo_sheet;
    HRESULT hr;
    VARIANT res, var_index;
    
    VariantInit( &res );
    VariantInit( &var_index );
    
    V_VT( &res ) = VT_BSTR;
    V_BSTR( &var_index ) = SysAllocString( _sheet_name );
    
    hr = AutoWrap (DISPATCH_METHOD, &res, m_pd_sheets, L"getByName", 1, var_index);
    if ( FAILED( hr ) )
    {
        ERR( " getByIndex \n" );
    } else
    {
        oo_sheet.Init( V_DISPATCH( &res ) );      
    }
    
 
    VariantClear( &res );
    VariantClear( &var_index );
 
    TRACE_OUT;
    return oo_sheet;           
}

HRESULT OOSheets::insertNewByName( BSTR _name, long _index )
{
    TRACE_IN;
    HRESULT hr;
    VARIANT param1, param2;
    VARIANT res;
    
    VariantInit( &param1 );
    VariantInit( &param2 );
    VariantInit( &res );
    
    V_VT( &param1 ) = VT_BSTR;
    V_BSTR( &param1 ) = SysAllocString( _name );
    
    V_VT( &param2 ) = VT_I4;
    V_I4( &param2 ) = _index;

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_sheets, L"insertNewByName", 2, param2, param1);
    if ( FAILED( hr ) ) 
    {
        ERR(" insertNewByName\n");
    }

    VariantClear( &param1 );
    VariantClear( &param2 );
    VariantClear( &res );
    
    TRACE_OUT;
    return E_NOTIMPL;          
}

HRESULT OOSheets::removeByName( BSTR _name )
{
    TRACE_IN;
    HRESULT hr;
    
    VARIANT param1;
    VARIANT res;
    
    VariantInit( &param1 );
    VariantInit( &res );
    
    V_VT( &param1 ) = VT_BSTR;
    V_BSTR( &param1 ) = SysAllocString( _name );

    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_sheets, L"removeByName", 1, param1);
    if ( FAILED( hr ) ) 
    {
        ERR(" removeByName \n");
    }

    VariantClear( &param1 );
    VariantClear( &res );
       
    TRACE_OUT;
    return ( hr );
}
