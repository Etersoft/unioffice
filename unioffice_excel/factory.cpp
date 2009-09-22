#include "factory.h"
#include "application.h"

//
// Реализация IUnknown для фабрики класса
//
HRESULT STDMETHODCALLTYPE CFactory::QueryInterface(const IID& iid, void** ppv)
{
        TRACE_IN;
        
        if ((iid == IID_IUnknown) || (iid == IID_IClassFactory))
           {
                 *ppv = static_cast<IClassFactory*>(this);
           }
           else
           {
                 *ppv = NULL;
                 return E_NOINTERFACE;
           }
           
        reinterpret_cast<IUnknown*>(*ppv)->AddRef();
           
        TRACE_OUT;   
           
        return S_OK;
}

ULONG STDMETHODCALLTYPE CFactory::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);
}
ULONG STDMETHODCALLTYPE CFactory::Release()
{
      TRACE( " ref = %i \n", m_cRef );
      
      if (InterlockedDecrement(&m_cRef) == 0)
      {
              delete this;
              return 0;
      }
      
      return m_cRef;
}

//
// Реализация IClassFactory
//
HRESULT STDMETHODCALLTYPE CFactory::CreateInstance(
                  IUnknown* pUnknownOuter,
                  const IID& iid,
                  void** ppv)
{
   HRESULT hr = S_FALSE;
        
   TRACE_IN;     
                  
   // Агрегирование не поддерживается
   if (pUnknownOuter != NULL)
   {
       return CLASS_E_NOAGGREGATION;
   }
   // Создать компонент
   
   Application* pApplication = new Application;
   if (pApplication == NULL)
   {
       return E_OUTOFMEMORY;
   }
   
   // Вернуть запрошенный интерфейс
   hr = pApplication->QueryInterface(iid, ppv);
   // Освободить указатель на IUnknown
   // (При ошибке в QueryInterface компонент разрушит сам себя)
   pApplication->Release();
   
   TRACE_OUT;
   
   return hr;
}
   
// LockServer
HRESULT STDMETHODCALLTYPE CFactory::LockServer(BOOL bLock)
{
        TRACE_IN;
        
        if (bLock)
        {
            InterlockedIncrement(&g_cServerLocks);
        }
        else
        {
            InterlockedDecrement(&g_cServerLocks);
        }
        
        TRACE_OUT;
        
        return S_OK;
}


