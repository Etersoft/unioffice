#include "../OOWrappers/oo_property_value.h"

OOPropertyValue::OOPropertyValue()
{
    TRACE_IN;
                                    
    m_pd_property_value = NULL;                                   
    
    TRACE_OUT;                        
}

OOPropertyValue::OOPropertyValue(const OOPropertyValue &obj)
{
   TRACE_IN;
                               
   m_pd_property_value = obj.m_pd_property_value;
   if ( m_pd_property_value != NULL )
       m_pd_property_value->AddRef();  
       
   TRACE_OUT;                      
}

OOPropertyValue::~OOPropertyValue()
{
   TRACE_IN;
   
   if ( m_pd_property_value != NULL )
   {
       m_pd_property_value->Release();
       m_pd_property_value = NULL;        
   }                                  
   
   TRACE_OUT;                         
}

OOPropertyValue& OOPropertyValue::operator=( const OOPropertyValue &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_property_value != NULL )
   {
       m_pd_property_value->Release();
       m_pd_property_value = NULL;        
   } 
   
   m_pd_property_value = obj.m_pd_property_value;
   if ( m_pd_property_value != NULL )
       m_pd_property_value->AddRef();
   
   return ( *this );          
    
}

void OOPropertyValue::Init( IDispatch* p_oo_property_value  )
{
   TRACE_IN; 
     
   if ( m_pd_property_value != NULL )
   {
       m_pd_property_value->Release();
       m_pd_property_value = NULL;        
   } 
   
   if ( p_oo_property_value == NULL )
   {
       ERR( " p_oo_property_value == NULL \n" );
       return;     
   }
   
   m_pd_property_value = p_oo_property_value;
   m_pd_property_value->AddRef();
   
   TRACE_OUT;
   
   return;
}

IDispatch* OOPropertyValue::GetOOProperty()
{
    if ( m_pd_property_value != NULL )
    {
       m_pd_property_value->AddRef();     
    } else
    {
        ERR( " m_pd_property_value is NULL \n" );      
    }           
          
    return ( m_pd_property_value );           
}

HRESULT OOPropertyValue::Set_PropertyName( BSTR _name )
{
    HRESULT hr = S_OK;
    VARIANT param;
    VARIANT res;
    
    TRACE_IN;
    
    VariantInit(&param);
    VariantInit(&res);
    
    V_VT(&param) = VT_BSTR;
    V_BSTR(&param) = _name;
    
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_property_value, L"Name", 1, param);
    
    if ( FAILED( hr ) )
    {
        ERR( " \n " );     
    }
    
    VariantClear(&param);
    VariantClear(&res);
    
    TRACE_OUT;
    
    return ( hr );        
}

HRESULT OOPropertyValue::Set_PropertyValue( BSTR _value )
{
    HRESULT hr = S_OK;
    VARIANT param;
    VARIANT res;
    
    TRACE_IN;
    
    VariantInit(&param);
    VariantInit(&res);
    
    V_VT(&param) = VT_BSTR;
    V_BSTR(&param) = _value;
    
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_property_value, L"Value", 1, param);
    
    if ( FAILED( hr ) )
    {
        ERR( " \n " );     
    }
    
    VariantClear(&param);
    VariantClear(&res);
    
    TRACE_OUT;
    
    return ( hr );         
}
 
HRESULT OOPropertyValue::Set_PropertyValue( VARIANT_BOOL _value)
{
    HRESULT hr = S_OK;
    VARIANT param;
    VARIANT res;
    
    TRACE_IN;
    
    VariantInit(&param);
    VariantInit(&res);
    
    V_VT(&param) = VT_BOOL;
    V_BOOL(&param) = _value;
    
    hr = AutoWrap(DISPATCH_PROPERTYPUT, &res, m_pd_property_value, L"Value", 1, param);
    
    if ( FAILED( hr ) )
    {
        ERR( " \n " );     
    }
    
    VariantClear(&param);
    VariantClear(&res);
    
    TRACE_OUT;
    
    return ( hr );   
      
} 
        
HRESULT OOPropertyValue::Set_Property( BSTR _name, BSTR _value )
{
    HRESULT hr = S_OK;
    
    TRACE_IN;
    
    hr = Set_PropertyName( _name );
    
    if ( FAILED( hr ) )
    {
        ERR( " Set_PropertyName() \n " );     
    }
    
    hr = Set_PropertyValue( _value );
    
    if ( FAILED( hr ) )
    {
        ERR( " Set_PropertyValue() \n " );     
    }
    
    TRACE_OUT;
    
    return ( hr );         
}
