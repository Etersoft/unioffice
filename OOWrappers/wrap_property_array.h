#ifndef __UNIOFFICE_WRAP_PROPERTY_ARRAY_H__
#define __UNIOFFICE_WRAP_PROPERTY_ARRAY_H__

#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../OOWrappers/oo_property_value.h"
#include <vector>

using namespace std;

class WrapPropertyArray
{
public:
       
  WrapPropertyArray();
  WrapPropertyArray(const WrapPropertyArray &);
  virtual ~WrapPropertyArray(); 
 
  WrapPropertyArray& operator=(const WrapPropertyArray &); 
  
  SAFEARRAY FAR* Get_SafeArray();
  
  void Clear();
  void Add( OOPropertyValue );
        
private:      
  
  SAFEARRAY FAR*             m_sa_property_values;
  
  vector<OOPropertyValue>    m_properties;
  
  HRESULT Build_SafeArray();
      
};

#endif //__UNIOFFICE_WRAP_PROPERTY_ARRAY_H__
