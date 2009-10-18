/*
 * implementation of Name
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

#include "name.h"


       // IUnknown
       HRESULT STDMETHODCALLTYPE CName::QueryInterface(const IID& iid, void** ppv)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       ULONG STDMETHODCALLTYPE CName::AddRef()
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       ULONG STDMETHODCALLTYPE CName::Release()
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
       
       // IDispatch    
       HRESULT STDMETHODCALLTYPE CName::GetTypeInfoCount( UINT * pctinfo )
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       HRESULT STDMETHODCALLTYPE CName::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       HRESULT STDMETHODCALLTYPE CName::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}

       HRESULT STDMETHODCALLTYPE CName::Invoke(
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
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get__Default( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Index( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Category( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Category( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_CategoryLocal( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_CategoryLocal( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CName::Delete( void)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_MacroType( 
            /* [retval][out] */ XlXLMMacroType *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_MacroType( 
            /* [in] */ XlXLMMacroType RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Name( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Name( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersTo( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersTo( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_ShortcutKey( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_ShortcutKey( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Value( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Value( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Visible( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Visible( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_NameLocal( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_NameLocal( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToLocal( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersToLocal( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToR1C1( 
            /* [lcid][in] */ long lcidIn,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersToR1C1( 
            /* [lcid][in] */ long lcidIn,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToR1C1Local( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_RefersToR1C1Local( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_RefersToRange( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_Comment( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_Comment( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_WorkbookParameter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CName::put_WorkbookParameter( 
            /* [in] */ VARIANT_BOOL RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CName::get_ValidWorkbookParameter( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
      
HRESULT CName::Init( )
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
       
HRESULT CName::Put_Application( void* p_application)
{
}

HRESULT CName::Put_Parent( void* p_parent)
{
}

