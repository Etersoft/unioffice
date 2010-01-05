/*
 * implementation of Borders
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

#include "borders.h"
#include "application.h"

       // IUnknown
       HRESULT STDMETHODCALLTYPE CBorders::QueryInterface(const IID& iid, void** ppv)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       ULONG STDMETHODCALLTYPE CBorders::AddRef()       
{
    TRACE_NOTIMPL;
	return 0; 		
} 

       ULONG STDMETHODCALLTYPE CBorders::Release()
{
    TRACE_NOTIMPL;
	return 0; 		
} 
       
       // IDispatch    
       HRESULT STDMETHODCALLTYPE CBorders::GetTypeInfoCount( UINT * pctinfo )
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       HRESULT STDMETHODCALLTYPE CBorders::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       HRESULT STDMETHODCALLTYPE CBorders::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

       HRESULT STDMETHODCALLTYPE CBorders::Invoke(
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

			   // IBorders              
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Color( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_Color( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_ColorIndex( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Count( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Item( 
            /* [in] */ XlBordersIndex Index,
            /* [retval][out] */ Border	**RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_LineStyle( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_LineStyle( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Value( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_Value( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_Weight( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_Weight( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get__Default( 
            /* [in] */ XlBordersIndex Index,
            /* [retval][out] */ Border	**RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_ThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CBorders::get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CBorders::put_TintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 
            
HRESULT CBorders::Init( )
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 	
} 
       
HRESULT CBorders::Put_Application( void* p_application)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
} 

HRESULT CBorders::Put_Parent( void* p_parent)
{
    TRACE_NOTIMPL;
	return E_NOTIMPL; 		
}            
     
	        
