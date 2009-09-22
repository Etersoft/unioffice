#ifndef __UNIOFFICE_EXCEL_FACTORY_H__
#define __UNIOFFICE_EXCEL_FACTORY_H__

#include "unioffice_excel_private.h"

///////////////////////////////////////////////////////////
//
// Фабрика класса
//
class CFactory : public IClassFactory
{
public:
       // IUnknown
       virtual HRESULT STDMETHODCALLTYPE QueryInterface(const IID& iid, void** ppv);
       virtual ULONG STDMETHODCALLTYPE AddRef();
       virtual ULONG STDMETHODCALLTYPE Release();
       
       // Интерфейс IClassFactory
       virtual HRESULT STDMETHODCALLTYPE CreateInstance(
               IUnknown* pUnknownOuter,
               const IID& iid,
               void** ppv);
       virtual HRESULT STDMETHODCALLTYPE LockServer(BOOL bLock);
       
       // Конструктор
       CFactory() : m_cRef(1) {CREATE_OBJECT;}
       // Деструктор
       virtual ~CFactory() { DELETE_OBJECT; }
       
private:
        long m_cRef;
};

#endif //__UNIOFFICE_EXCEL_FACTORY_H__
