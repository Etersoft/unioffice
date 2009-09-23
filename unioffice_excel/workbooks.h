/*
 * header file - Workbooks
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

#ifndef __UNIOFFICE_EXCEL_WORKBOOKS_H__
#define __UNIOFFICE_EXCEL_WORKBOOKS_H__

#include "unioffice_excel_private.h"
#include "workbook.h"
#include <list>

class CWorkbooks : public Workbooks, public IEnumVARIANT
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
       
        //Workbooks
        virtual HRESULT STDMETHODCALLTYPE get_Application( 
             Application	**RHS);
        
        virtual HRESULT STDMETHODCALLTYPE get_Creator( 
             XlCreator *RHS);
        
        virtual HRESULT STDMETHODCALLTYPE get_Parent( 
             IDispatch **RHS);
        
        virtual HRESULT STDMETHODCALLTYPE Add( 
             VARIANT Template,
             long lcid,
             Workbook **RHS);
        
        virtual HRESULT STDMETHODCALLTYPE Close( 
             long lcid);
        
        virtual HRESULT STDMETHODCALLTYPE get_Count( 
             long *RHS);
        
        virtual HRESULT STDMETHODCALLTYPE get_Item( 
             VARIANT Index,
             Workbook **RHS);
        
        virtual HRESULT STDMETHODCALLTYPE get__NewEnum( 
             IUnknown **RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE _Open( 
             BSTR Filename,
             VARIANT UpdateLinks,
             VARIANT ReadOnly,
             VARIANT Format,
             VARIANT Password,
             VARIANT WriteResPassword,
             VARIANT IgnoreReadOnlyRecommended,
             VARIANT Origin,
             VARIANT Delimiter,
             VARIANT Editable,
             VARIANT Notify,
             VARIANT Converter,
             VARIANT AddToMru,
             long lcid,
             Workbook **RHS);
        
        virtual HRESULT STDMETHODCALLTYPE __OpenText( 
             BSTR Filename,
             VARIANT Origin,
             VARIANT StartRow,
             VARIANT DataType,
             XlTextQualifier TextQualifier,
             VARIANT ConsecutiveDelimiter,
             VARIANT Tab,
             VARIANT Semicolon,
             VARIANT Comma,
             VARIANT Space,
             VARIANT Other,
             VARIANT OtherChar,
             VARIANT FieldInfo,
             VARIANT TextVisualLayout,
             long lcid);
        
        virtual  HRESULT STDMETHODCALLTYPE get__Default( 
             VARIANT Index,
             Workbook **RHS);
        
        virtual HRESULT STDMETHODCALLTYPE _OpenText( 
             BSTR Filename,
             VARIANT Origin,
             VARIANT StartRow,
             VARIANT DataType,
             XlTextQualifier TextQualifier,
             VARIANT ConsecutiveDelimiter,
             VARIANT Tab,
             VARIANT Semicolon,
             VARIANT Comma,
             VARIANT Space,
             VARIANT Other,
             VARIANT OtherChar,
             VARIANT FieldInfo,
             VARIANT TextVisualLayout,
             VARIANT DecimalSeparator,
             VARIANT ThousandsSeparator,
             long lcid);
        
        virtual  HRESULT STDMETHODCALLTYPE Open( 
             BSTR Filename,
             VARIANT UpdateLinks,
             VARIANT ReadOnly,
             VARIANT Format,
             VARIANT Password,
             VARIANT WriteResPassword,
             VARIANT IgnoreReadOnlyRecommended,
             VARIANT Origin,
             VARIANT Delimiter,
             VARIANT Editable,
             VARIANT Notify,
             VARIANT Converter,
             VARIANT AddToMru,
             VARIANT Local,
             VARIANT CorruptLoad,
             long lcid,
             Workbook **RHS);
        
        virtual HRESULT STDMETHODCALLTYPE OpenText( 
             BSTR Filename,
             VARIANT Origin,
             VARIANT StartRow,
             VARIANT DataType,
             XlTextQualifier TextQualifier,
             VARIANT ConsecutiveDelimiter,
             VARIANT Tab,
             VARIANT Semicolon,
             VARIANT Comma,
             VARIANT Space,
             VARIANT Other,
             VARIANT OtherChar,
             VARIANT FieldInfo,
             VARIANT TextVisualLayout,
             VARIANT DecimalSeparator,
             VARIANT ThousandsSeparator,
             VARIANT TrailingMinusNumbers,
             VARIANT Local,
             long lcid);
        
        virtual HRESULT STDMETHODCALLTYPE OpenDatabase( 
             BSTR Filename,
             VARIANT CommandText,
             VARIANT CommandType,
             VARIANT BackgroundQuery,
             VARIANT ImportDataAs,
             Workbook **RHS);
        
        virtual HRESULT STDMETHODCALLTYPE CheckOut( 
             BSTR Filename);
        
        virtual HRESULT STDMETHODCALLTYPE CanCheckOut( 
             BSTR Filename,
             VARIANT_BOOL *RHS);
        
        virtual HRESULT STDMETHODCALLTYPE _OpenXML( 
             BSTR Filename,
             VARIANT Stylesheets,
             Workbook **RHS);
        
        virtual  HRESULT STDMETHODCALLTYPE OpenXML( 
             BSTR Filename,
             VARIANT Stylesheets,
             VARIANT LoadOption,
             Workbook **RHS);       
       
       	virtual HRESULT Next ( ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched);
	    virtual HRESULT Skip ( ULONG celt);
	    virtual HRESULT Reset( );
	    virtual HRESULT Clone(IEnumVARIANT** ppEnum);
       
       
       CWorkbooks()
       {
            CREATE_OBJECT; 
            m_cRef = 1;
            m_pITypeInfo = NULL;
            
            m_lst_of_workbook.clear();
            m_it_of_workbook = m_lst_of_workbook.end();
            
            m_p_application = NULL;
            m_p_parent = NULL;
            
            enum_position = 0;
            
            HRESULT hr = Init();
            
            if ( FAILED(hr) )
            {
                 ERR( " \n " );
            }
            
            InterlockedIncrement(&g_cComponents);         
       }
       
       virtual ~CWorkbooks()
       {
            InterlockedDecrement(&g_cComponents);    
            
            std::list< Workbook* >::iterator it_begin = m_lst_of_workbook.begin();
            
            while ( it_begin != m_lst_of_workbook.end() )
            {
                  (*it_begin)->Release();
                  
                  it_begin++;
            }
            
            m_lst_of_workbook.clear();  
              
            m_p_application = NULL;
            m_p_parent = NULL;
                
            DELETE_OBJECT;             
       }
       
       HRESULT Init();
       
       HRESULT Put_Visible( VARIANT_BOOL );
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* );
       
       HRESULT DeleteWorkbookFromVector( Workbook* );
       
private:
        
       long m_cRef; 
       
       ITypeInfo* m_pITypeInfo;    
        
       std::list< Workbook* >             m_lst_of_workbook; 
       std::list< Workbook* >::iterator   m_it_of_workbook; 
       
       void*   m_p_application;
       void*   m_p_parent;
       
       int enum_position;
};


#endif //__UNIOFFICE_EXCEL_WORKBOOKS_H__
