/*
 * implementation of CFactory
 *
 * Copyright (C) 2009 Sinitsin Ivan (Etersoft) <ivan@etersoft.ru>
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301, USA
 */

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


