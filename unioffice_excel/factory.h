/*
 * header file - CFactory
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
