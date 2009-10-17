/*
 * implementation of Names
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

#include "names.h"

       // IUnknown
       HRESULT STDMETHODCALLTYPE CNames::QueryInterface(const IID& iid, void** ppv)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       ULONG STDMETHODCALLTYPE CNames::AddRef()
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       ULONG STDMETHODCALLTYPE CNames::Release()
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
       
       // IDispatch    
       HRESULT STDMETHODCALLTYPE CNames::GetTypeInfoCount( UINT * pctinfo )
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       HRESULT STDMETHODCALLTYPE CNames::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       HRESULT STDMETHODCALLTYPE CNames::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       HRESULT STDMETHODCALLTYPE CNames::Invoke(
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
               
        //Names
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CNames::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CNames::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CNames::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CNames::Add( 
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT RefersTo,
            /* [optional][in] */ VARIANT Visible,
            /* [optional][in] */ VARIANT MacroType,
            /* [optional][in] */ VARIANT ShortcutKey,
            /* [optional][in] */ VARIANT Category,
            /* [optional][in] */ VARIANT NameLocal,
            /* [optional][in] */ VARIANT RefersToLocal,
            /* [optional][in] */ VARIANT CategoryLocal,
            /* [optional][in] */ VARIANT RefersToR1C1,
            /* [optional][in] */ VARIANT RefersToR1C1Local,
            /* [retval][out] */ Name	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CNames::Item( 
            /* [optional][in] */ VARIANT Index,
            /* [optional][in] */ VARIANT IndexLocal,
            /* [optional][in] */ VARIANT RefersTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Name	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CNames::_Default( 
            /* [optional][in] */ VARIANT Index,
            /* [optional][in] */ VARIANT IndexLocal,
            /* [optional][in] */ VARIANT RefersTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Name	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CNames::get_Count( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CNames::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
			
HRESULT CNames::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_INames, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;		
}
       
HRESULT CNames::Put_Application( void* p_application )
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;
}

HRESULT CNames::Put_Parent( void* p_parent )
{
   TRACE_IN;  
      
   m_p_parent = p_parent;
   
   TRACE_OUT;
   return S_OK;	
}

		               
