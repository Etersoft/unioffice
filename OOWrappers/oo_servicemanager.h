#ifndef __UNIOFFICE_OO_WRAP_SERVICE_MANAGER_H__
#define __UNIOFFICE_OO_WRAP_SERVICE_MANAGER_H__


#include <ole2.h>
#include <oaidl.h>
#include "../Common/debug.h"
#include "../Common/tools.h"
#include "../OOWrappers/oo_desktop.h"

class OOServiceManager
{
public:
       
  OOServiceManager();
  virtual ~OOServiceManager();     
  
  IDispatch* CreateInstance( BSTR );
  IDispatch* Bridge_GetStruct( BSTR ); 
   
  OODesktop         Get_Desktop( ); 
  OOPropertyValue   Get_PropertyValue(  );
       
private:            
   
   IDispatch*   m_pd_servicemanager;  
      
};






#endif //__UNIOFFICE_OO_WRAP_SERVICE_MANAGER_H__
