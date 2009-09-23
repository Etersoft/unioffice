/*
 * implementation of OODesktop
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

#include "../OOWrappers/oo_desktop.h"


OODesktop::OODesktop()
{    
    TRACE_IN;
                                    
    m_pd_desktop = NULL;                                   
    
    TRACE_OUT;
}


OODesktop::OODesktop(const OODesktop &obj)
{
   TRACE_IN;
                               
   m_pd_desktop = obj.m_pd_desktop;
   if ( m_pd_desktop != NULL )
       m_pd_desktop->AddRef();  
       
   TRACE_OUT;                      
}


OODesktop::~OODesktop()
{
   TRACE_IN;
   
   if ( m_pd_desktop != NULL )
   {
       m_pd_desktop->Release();
       m_pd_desktop = NULL;        
   }                                  
   
   TRACE_OUT;
}



OODesktop& OODesktop::operator=(const OODesktop &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_desktop != NULL )
   {
       m_pd_desktop->Release();
       m_pd_desktop = NULL;        
   } 
   
   m_pd_desktop = obj.m_pd_desktop;
   if ( m_pd_desktop != NULL )
       m_pd_desktop->AddRef();
   
   return ( *this );          
    
}

void OODesktop::Init( IDispatch* p_oo_desktop )
{
    TRACE_IN; 
     
   if ( m_pd_desktop != NULL )
   {
       m_pd_desktop->Release();
       m_pd_desktop = NULL;        
   } 
   
   if ( p_oo_desktop == NULL )
   {
       ERR( " p_oo_desktop == NULL \n" );
       return;     
   }
   
   m_pd_desktop = p_oo_desktop;
   m_pd_desktop->AddRef();
   
   TRACE_OUT;
   
   return;
}

HRESULT OODesktop::terminate()
{
    VARIANT res;
    HRESULT hr = S_OK;
    
    VariantInit( &res );
    
    hr = AutoWrap(DISPATCH_METHOD, &res, m_pd_desktop, L"terminate", 0);

    if ( FAILED( hr ) )
    {
        ERR( " m_pd_desktop -> terminate \n" );
        return S_FALSE;     
    }
    
    return ( hr );
}

OODocument OODesktop::LoadComponentFromURL( BSTR _type_doc, BSTR _template, long _not_used, WrapPropertyArray& _property_array )
{
    HRESULT hr;
    VARIANT param0,param1,param2,param3;
    OODocument document;  
    VARIANT resultDoc;
  
    TRACE_IN;
  
    VariantInit(&param0);
    VariantInit(&param1);
    VariantInit(&param2);
    VariantInit(&param3);
    VariantInit(&resultDoc);
  
    //type of document
    V_VT(&param0) = VT_BSTR;
    V_BSTR(&param0) = SysAllocString(_type_doc);
  
    // template for new document   
    V_VT(&param1) = VT_BSTR;
    V_BSTR(&param1) = SysAllocString(_template);
  
    // not used parameters
    V_VT(&param2) = VT_I2;
    V_I2(&param2) = _not_used;  
  
    V_VT(&param3) = VT_ARRAY | VT_DISPATCH;
    V_ARRAY(&param3) = _property_array.Get_SafeArray();
  
    hr = AutoWrap(DISPATCH_METHOD, &resultDoc, m_pd_desktop, L"loadComponentFromURL", 4, param3, param2, param1, param0);
    
    if ( FAILED(hr) ) {
        ERR( " LoadComponentFromURL \n" ); 
        return ( document );
    }
  
    document.Init( resultDoc.pdispVal );
  
    
    VariantClear(&param0);
    VariantClear(&param1);
    VariantClear(&param2);
    VariantClear(&param3);
    VariantClear(&resultDoc);
  
    TRACE_OUT;
  
    return ( document );      
}

