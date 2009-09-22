#include "../OOWrappers/oo_document.h"

OODocument::OODocument()
{
    TRACE_IN;
                                    
    m_pd_document = NULL;                                   
    
    TRACE_OUT;                        
}

OODocument::OODocument(const OODocument &obj)
{
   TRACE_IN;
                               
   m_pd_document = obj.m_pd_document;
   if ( m_pd_document != NULL )
       m_pd_document->AddRef();  
       
   TRACE_OUT;                      
}

OODocument::~OODocument()
{
   TRACE_IN;
   
   if ( m_pd_document != NULL )
   {
       m_pd_document->Release();
       m_pd_document = NULL;        
   }                                  
   
   TRACE_OUT;                         
}

OODocument& OODocument::operator=( const OODocument &obj)
{
   if ( this == &obj )
   {
       return ( *this );                 
   }    
   
   if ( m_pd_document != NULL )
   {
       m_pd_document->Release();
       m_pd_document = NULL;        
   } 
   
   m_pd_document = obj.m_pd_document;
   if ( m_pd_document != NULL )
       m_pd_document->AddRef();
   
   return ( *this );          
    
}

void OODocument::Init( IDispatch* p_oo_document )
{
   TRACE_IN; 
     
   if ( m_pd_document != NULL )
   {
       m_pd_document->Release();
       m_pd_document = NULL;        
   } 
   
   if ( p_oo_document == NULL )
   {
       ERR( " p_oo_document == NULL \n" );
       return;     
   }
   
   m_pd_document = p_oo_document;
   m_pd_document->AddRef();
   
   TRACE_OUT;
   
   return;
}

bool OODocument::IsNull()
{
    return ( (m_pd_document == NULL) ? true : false );     
}

HRESULT OODocument::StoreAsURL( BSTR _filename, WrapPropertyArray& _property_array )
{
    HRESULT hr;
    VARIANT param0,param1;
    VARIANT result;
  
    TRACE_IN;
  
    VariantInit(&param0);
    VariantInit(&param1);
    VariantInit(&result); 
     
    //type of document
    V_VT(&param0) = VT_BSTR;
    V_BSTR(&param0) = SysAllocString( _filename );
    
    V_VT(&param1) = VT_ARRAY | VT_DISPATCH;
    V_ARRAY(&param1) = _property_array.Get_SafeArray();
    
    TRACE( " Filename = " );
    int i=0;
    while (*(_filename+i)!=0) {
        WTRACE_HARD(L"%c",*(_filename+i));
        i++;
    }
    WTRACE_HARD(L"\n");
        
    hr = AutoWrap(DISPATCH_METHOD, &result, m_pd_document, L"StoreAsURL", 2, param1, param0);
    
    if ( FAILED( hr ) ) {
        ERR( " StoreAsURL \n" ); 
        return ( E_FAIL );
    }
    
    VariantClear(&param0);
    VariantClear(&param1);
    VariantClear(&result);
  
    TRACE_OUT;
     
    return ( hr );    
} 

HRESULT OODocument::Close( VARIANT_BOOL _hard_close )
{
    HRESULT hr;
    VARIANT res;
    VARIANT hard_close;
    
    TRACE_IN;
    
    VariantInit( &hard_close );
    VariantInit( &res );
    
    V_VT( &hard_close )   = VT_BOOL;
    V_BOOL( &hard_close ) = _hard_close;
    
    hr = AutoWrap( DISPATCH_METHOD, &res, m_pd_document, L"close", 1, hard_close );

    if ( FAILED( hr ) )
    { 
            ERR(" FAILED 1 CLOSE \n"); 
    }
            
    VariantClear( &res );
    VariantClear( &hard_close );       
               
    TRACE_OUT;        
    return ( hr );          
}
