#ifndef __UNIOFFICE_EXCEL_FACTORY_H__
#define __UNIOFFICE_EXCEL_FACTORY_H__

#include "unioffice_excel_private.h"

///////////////////////////////////////////////////////////
//
// ������� ������
//
class CFactory : public IClassFactory
{
public:
       // IUnknown
       virtual HRESULT STDMETHODCALLTYPE QueryInterface(const IID& iid, void** ppv);
       virtual ULONG STDMETHODCALLTYPE AddRef();
       virtual ULONG STDMETHODCALLTYPE Release();
       
       // ��������� IClassFactory
       virtual HRESULT STDMETHODCALLTYPE CreateInstance(
               IUnknown* pUnknownOuter,
               const IID& iid,
               void** ppv);
       virtual HRESULT STDMETHODCALLTYPE LockServer(BOOL bLock);
       
       // �����������
       CFactory() : m_cRef(1) {CREATE_OBJECT;}
       // ����������
       virtual ~CFactory() { DELETE_OBJECT; }
       
private:
        long m_cRef;
};

#endif //__UNIOFFICE_EXCEL_FACTORY_H__
