#ifndef __UNIOFFICE_OO_WRAP_PROPERTY_VALUE_H__
#define __UNIOFFICE_OO_WRAP_PROPERTY_VALUE_H__

#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../Common/tools.h"

class OOPropertyValue
{
public:
       
  OOPropertyValue();
  OOPropertyValue(const OOPropertyValue &);
  virtual ~OOPropertyValue();     
   
  OOPropertyValue& operator=( const OOPropertyValue &); 
  
  void Init( IDispatch* );  
  IDispatch* GetOOProperty();
  
  HRESULT Set_PropertyName( BSTR );
  HRESULT Set_PropertyValue( BSTR );
  HRESULT Set_Property( BSTR, BSTR ); 
  
  HRESULT Set_PropertyValue( VARIANT_BOOL );
  
       
private:            
   
   IDispatch*   m_pd_property_value;  
      
};

#endif //__UNIOFFICE_OO_WRAP_PROPERTY_VALUE_H__
