/*
 * implementation of Interior
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

#include "interior.h"
#include "application.h"

       // IUnknown
       HRESULT STDMETHODCALLTYPE CInterior::QueryInterface(const IID& iid, void** ppv)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       ULONG STDMETHODCALLTYPE CInterior::AddRef()
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       ULONG STDMETHODCALLTYPE CInterior::Release()
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       
       // IDispatch    
       HRESULT STDMETHODCALLTYPE CInterior::GetTypeInfoCount( UINT * pctinfo )
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       HRESULT STDMETHODCALLTYPE CInterior::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       HRESULT STDMETHODCALLTYPE CInterior::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       HRESULT STDMETHODCALLTYPE CInterior::Invoke(
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

			   // IInterior               
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Color( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_Color( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_ColorIndex( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_InvertIfNegative( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_InvertIfNegative( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Pattern( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_Pattern( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternColorIndex( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_ThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_TintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_PatternTintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CInterior::put_PatternTintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CInterior::get_Gradient( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
       
HRESULT CInterior::Init( )
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
       
HRESULT CInterior::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK; 		
} 

HRESULT CInterior::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;       
      
   TRACE_OUT;
   return S_OK; 		
} 
