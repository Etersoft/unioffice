/*
 * implementation of Outline
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

#include "outline.h"

#include "application.h"
#include "worksheet.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE COutline::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<IOutline*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<IOutline*>(this));
    }     
    
    if ( iid == IID_IOutline) {
        TRACE("IOutline\n");
        *ppv = static_cast<IOutline*>(this);
    } 
    
    if ( iid == DIID_Outline ) {
        TRACE("Outline \n");
        *ppv = static_cast<Outline*>(this);
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

        ULONG STDMETHODCALLTYPE COutline::AddRef()
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}

        ULONG STDMETHODCALLTYPE COutline::Release()
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
       
       // IDispatch    
        HRESULT STDMETHODCALLTYPE COutline::GetTypeInfoCount( UINT * pctinfo )
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}

        HRESULT STDMETHODCALLTYPE COutline::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}

        HRESULT STDMETHODCALLTYPE COutline::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}

        HRESULT STDMETHODCALLTYPE COutline::Invoke(
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
			
			
		// Outline
         /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE COutline::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE COutline::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE COutline::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE COutline::get_AutomaticStyles( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE COutline::put_AutomaticStyles( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext] */ HRESULT STDMETHODCALLTYPE COutline::ShowLevels( 
            /* [optional][in] */ VARIANT RowLevels,
            /* [optional][in] */ VARIANT ColumnLevels,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE COutline::get_SummaryColumn( 
            /* [retval][out] */ XlSummaryColumn *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE COutline::put_SummaryColumn( 
            /* [in] */ XlSummaryColumn RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE COutline::get_SummaryRow( 
            /* [retval][out] */ XlSummaryRow *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
         /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE COutline::put_SummaryRow( 
            /* [in] */ XlSummaryRow RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
            
            
HRESULT COutline::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_IOutline, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;  		
}
       
HRESULT COutline::Put_Application( void* p_application) 
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK; 
}

HRESULT COutline::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;
   
   TRACE_OUT;
   return S_OK;		
}
