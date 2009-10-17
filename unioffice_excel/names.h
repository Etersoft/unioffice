/*
 * header file - Names
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

#ifndef __UNIOFFICE_EXCEL_NAMES_H__
#define __UNIOFFICE_EXCEL_NAMES_H__

#include "unioffice_excel_private.h"
#include "../OOWrappers/oo_named_ranges.h"

class CNames : public INames, public Names, public IEnumVARIANT
{
public:
	   
       // IUnknown
       virtual HRESULT STDMETHODCALLTYPE QueryInterface(const IID& iid, void** ppv);
       virtual ULONG STDMETHODCALLTYPE AddRef();
       virtual ULONG STDMETHODCALLTYPE Release();
       
       // IDispatch    
       virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount( UINT * pctinfo );
       virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo);
       virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId);
       virtual HRESULT STDMETHODCALLTYPE Invoke(
               DISPID dispIdMember,
               REFIID riid,
               LCID lcid,
               WORD wFlags,
               DISPPARAMS * pDispParams,
               VARIANT * pVarResult,
               EXCEPINFO * pExcepInfo,
               UINT * puArgErr);
		
		
        //Names
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Add( 
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
            /* [retval][out] */ Name	**RHS);
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Item( 
            /* [optional][in] */ VARIANT Index,
            /* [optional][in] */ VARIANT IndexLocal,
            /* [optional][in] */ VARIANT RefersTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Name	**RHS);
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE _Default( 
            /* [optional][in] */ VARIANT Index,
            /* [optional][in] */ VARIANT IndexLocal,
            /* [optional][in] */ VARIANT RefersTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Name	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS);
	  
 	    virtual HRESULT Next ( ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched);
	    virtual HRESULT Skip ( ULONG celt);
	    virtual HRESULT Reset( );
	    virtual HRESULT Clone(IEnumVARIANT** ppEnum);
	  
       CNames()
       {
            CREATE_OBJECT; 
            m_cRef = 1;
            m_pITypeInfo = NULL;
            
            m_p_application = NULL;
            m_p_parent = NULL;
                        
            HRESULT hr = Init();
            
            if ( FAILED(hr) )
            {
                 ERR( " \n " );
            }
            
            InterlockedIncrement(&g_cComponents);         
       } 

       virtual ~CNames()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }	  
	  
       HRESULT Init( ); 
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* );
	   
	   HRESULT InitWrapper( OONamedRanges );	  		
			
private:
	
       long         m_cRef; 
       
       ITypeInfo*   m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;	
       
       int          enum_position;
       
       OONamedRanges  m_oo_named_ranges;
	   	
};


#endif //__UNIOFFICE_EXCEL_NAMES_H__
