/*
 * implementation of Sheets
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

#include "sheets.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CSheets::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(this);
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(this);
    }     
    
    if ( iid == IID_Sheets ) {
        TRACE("_Workbook \n");
        *ppv = static_cast<_Workbook*>(this);
    } 
      
    if ( *ppv != NULL ) 
    {
        reinterpret_cast<IUnknown*>(*ppv)->AddRef();
         
        return S_OK;
    } else
    {    
        WCHAR str_clsid[39];
         
        StringFromGUID2( iid, str_clsid, 39);
        WTRACE(L"(%s) not supported \n", str_clsid);
        
        return E_NOINTERFACE;                          
    }       
}

ULONG STDMETHODCALLTYPE CSheets::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);              
}

ULONG STDMETHODCALLTYPE CSheets::Release()
{
      TRACE( " ref = %i \n", m_cRef );
      
      if (InterlockedDecrement(&m_cRef) == 0)
      {
              delete this;
              return 0;
      }
      
      return m_cRef;          
}
       
       // IDispatch    
HRESULT STDMETHODCALLTYPE CSheets::GetTypeInfoCount( UINT * pctinfo )
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}

HRESULT STDMETHODCALLTYPE CSheets::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}

HRESULT STDMETHODCALLTYPE CSheets::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}

HRESULT STDMETHODCALLTYPE CSheets::Invoke(
               DISPID dispIdMember,
               REFIID riid,
               LCID lcid,
               WORD wFlags,
               DISPPARAMS * pDispParams,
               VARIANT * pVarResult,
               EXCEPINFO * pExcepInfo,
               UINT * puArgErr)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
} 
               
        // Sheets
HRESULT STDMETHODCALLTYPE CSheets::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Add( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [optional][in] */ VARIANT Count,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Copy( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Count( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Delete( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::FillAcrossSheets( 
            /* [in] */ Range	*Range,
            /* [defaultvalue][optional][in] */ XlFillWith Type,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Item( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Move( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::__PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Select( 
            /* [optional][in] */ VARIANT Replace,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_HPageBreaks( 
            /* [retval][out] */ HPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_VPageBreaks( 
            /* [retval][out] */ vPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get__Default( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::_PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [optional][in] */ VARIANT IgnorePrintAreas,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
            
            





HRESULT CSheets::Init()
{
     HRESULT hr = S_OK;   
      
     if (m_pITypeInfo == NULL)
     {
        ITypeLib* pITypeLib = NULL;
        hr = LoadRegTypeLib(LIBID_Office,
                                1, 0, // Номера версии
                                0x00,
                                &pITypeLib); 
        
       if (FAILED(hr))
       {
           ERR( " Typelib not register \n" );
           return hr;
       } 
        
       // Получить информацию типа для интерфейса объекта
       hr = pITypeLib->GetTypeInfoOfGuid(IID_Sheets, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;        
}

HRESULT CSheets::Put_Application( void* p_application )
{
    m_p_application = p_application;
        
    return S_OK;      
}

HRESULT CSheets::Put_Parent( void* p_parent )
{
   m_p_parent = p_parent;
   
   return S_OK;     
}








