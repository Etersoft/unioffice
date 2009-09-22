#include "../OOWrappers/wrap_property_array.h"


WrapPropertyArray::WrapPropertyArray()
{
    TRACE_IN;                                  
                                      
    m_sa_property_values = NULL;
    
    TRACE_OUT;                  
}

WrapPropertyArray::WrapPropertyArray(const WrapPropertyArray &obj)
{
                       
}

WrapPropertyArray::~WrapPropertyArray()
{
    TRACE_IN;
    
    HRESULT hr;
    
    if ( m_sa_property_values != NULL )
    {
        hr = SafeArrayDestroy( m_sa_property_values ); 
         
        if ( FAILED( hr ) )
        {
            ERR( " SafeArrayDestroy \n" );     
        }  
        
        m_sa_property_values = NULL;   
    }
    
    TRACE_OUT;        
}

WrapPropertyArray& WrapPropertyArray::operator=(const WrapPropertyArray &obj)
{
   HRESULT hr;                
                   
   if ( this == &obj )
   {
       return ( *this );                 
   }  
   
   m_properties = obj.m_properties;   
   
   if ( m_sa_property_values != NULL )
   {
        hr = SafeArrayDestroy( m_sa_property_values ); 
         
        if ( FAILED( hr ) )
        {
            ERR( " SafeArrayDestroy \n" );     
        }  
        
        m_sa_property_values = NULL;   
   }
   
   hr = SafeArrayCopy( obj.m_sa_property_values, &m_sa_property_values);
   
   if ( FAILED( hr ) )
   {
       ERR( " SafeArrayCopy \n" );     
   } 
   
   return ( *this );                 
}

HRESULT WrapPropertyArray::Build_SafeArray()
{
    HRESULT hr = S_OK;
     
    if ( m_sa_property_values != NULL )
    {
        hr = SafeArrayDestroy( m_sa_property_values ); 
         
        if ( FAILED( hr ) )
        {
            ERR( " SafeArrayDestroy \n" );
            return ( hr );     
        }
        
        m_sa_property_values = NULL;     
    } 
    
    int array_size = m_properties.size();
        
    m_sa_property_values = SafeArrayCreateVector( VT_DISPATCH, 0, array_size );
    
    TRACE( " SafeArrayCreateVector size = %i \n", array_size ); 
    
    if ( m_sa_property_values == NULL )
    {
        ERR( " SafeArrayCreateVector \n" ); 
        return ( E_FAIL );     
    }
    
    for (long i = 0; i < array_size; i++)
    {
        hr = SafeArrayPutElement( m_sa_property_values, &i, m_properties[i].GetOOProperty() );
        
        if ( FAILED( hr ) )    
        {
            ERR( " SafeArrayPutElement \n" );
            return ( E_FAIL ); 
        }
        
        TRACE(" SafeArrayPutElement index = %i \n", i);
    }
    
    return ( hr );       
}


SAFEARRAY FAR* WrapPropertyArray::Get_SafeArray()
{
    TRACE_IN;
    HRESULT hr = S_OK;
    SAFEARRAY FAR* ret_value;
            
    hr = Build_SafeArray( );
    
    if ( FAILED( hr ) )
    {
        ERR( " Build_SafeArray \n " ); 
        return ( NULL );     
    }
    
    hr = SafeArrayCopy( m_sa_property_values, &ret_value);
   
    if ( FAILED( hr ) )
    {
       ERR( " SafeArrayCopy \n" );     
    } 
    
    TRACE_OUT;
    return ( ret_value );           
}

void WrapPropertyArray::Clear()
{
    m_properties.clear(); 
}

void WrapPropertyArray::Add( OOPropertyValue _property )
{
   m_properties.push_back( _property );  
}
