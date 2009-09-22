#ifndef __UNIOFFICE_OO_WRAP_DESKTOP_H__
#define __UNIOFFICE_OO_WRAP_DESKTOP_H__

#include <ole2.h>
#include <oaidl.h>

#include "../Common/debug.h"
#include "../Common/tools.h"

#include "../OOWrappers/oo_document.h"
#include "../OOWrappers/wrap_property_array.h"

class OODesktop
{
public:
       
    OODesktop();
    OODesktop( const OODesktop & );
    virtual ~OODesktop();       
    
    OODesktop &operator=(const OODesktop& );
    
    void Init( IDispatch* p_oo_desktop );  
    
    OODocument LoadComponentFromURL( BSTR, BSTR, long, WrapPropertyArray& );
    
    HRESULT terminate();
       
private:             
      
   IDispatch*   m_pd_desktop;      
      
};




#endif // __UNIOFFICE_OO_WRAP_DESKTOP_H__
