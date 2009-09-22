#ifndef __UNIOFFICE_OO_WRAP_DOCUMENT_H__
#define __UNIOFFICE_OO_WRAP_DOCUMENT_H__

#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../Common/tools.h"
#include "../OOWrappers/wrap_property_array.h"

class OODocument
{
public:
       
  OODocument();
  OODocument(const OODocument &);
  virtual ~OODocument();     
   
  OODocument& operator=( const OODocument &); 
  
  void Init( IDispatch* );
  
  bool IsNull();
  
  HRESULT StoreAsURL( BSTR ,WrapPropertyArray& );
  
  HRESULT Close( VARIANT_BOOL );
       
private:            
   
   IDispatch*   m_pd_document;  
      
};

#endif //__UNIOFFICE_OO_WRAP_DOCUMENT_H__
