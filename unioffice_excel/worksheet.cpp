/*
 * implementation of Worksheet
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

#include "worksheet.h"
#include "application.h"
#include "sheets.h"


       // IUnknown
HRESULT STDMETHODCALLTYPE Worksheet::QueryInterface(const IID& iid, void** ppv)
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
    
    if ( iid == IID__Worksheet ) {
        TRACE("_Worksheet \n");
        *ppv = static_cast<_Worksheet*>(this);
    } 
      
    if ( iid == CLSID_Worksheet ) {
        TRACE("Worksheet \n");
        *ppv = static_cast<Worksheet*>(this);
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

ULONG STDMETHODCALLTYPE Worksheet::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);                          
}

ULONG STDMETHODCALLTYPE Worksheet::Release()
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
HRESULT STDMETHODCALLTYPE Worksheet::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;                           
}

HRESULT STDMETHODCALLTYPE Worksheet::GetTypeInfo(
               UINT iTInfo,
               LCID lcid,
               ITypeInfo ** ppTInfo)
{
    *ppTInfo = NULL;
    
    if(iTInfo != 0)
    {
        return DISP_E_BADINDEX;
    }
    
    m_pITypeInfo->AddRef();
    *ppTInfo = m_pITypeInfo;
    
    return S_OK;                        
}
               
HRESULT STDMETHODCALLTYPE Worksheet::GetIDsOfNames(
               REFIID riid,
               LPOLESTR * rgszNames,
               UINT cNames,
               LCID lcid,
               DISPID * rgDispId)
{
    if (riid != IID_NULL )
    {
        return DISP_E_UNKNOWNINTERFACE;
    }
    
    HRESULT hr = m_pITypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
    
    if ( FAILED(hr) )
    {
     ERR( " name = %s \n", *rgszNames );     
    }
    
    return hr;                           
}
               
HRESULT STDMETHODCALLTYPE Worksheet::Invoke(
               DISPID dispIdMember,
               REFIID riid,
               LCID lcid,
               WORD wFlags,
               DISPPARAMS * pDispParams,
               VARIANT * pVarResult,
               EXCEPINFO * pExcepInfo,
               UINT * puArgErr)
{
    if ( riid != IID_NULL)
    {
        return DISP_E_UNKNOWNINTERFACE;
    }
    
    HRESULT hr = m_pITypeInfo->Invoke(
                 static_cast<IDispatch*>(this), 
                 dispIdMember, 
                 wFlags, 
                 pDispParams, 
                 pVarResult, 
                 pExcepInfo, 
                 puArgErr);
      
    if ( FAILED(hr) )
    {
     ERR( " dispIdMember = %i \n", dispIdMember );     
    }  
                 
    return hr;                          
} 
       
       
       // _Worksheet
HRESULT STDMETHODCALLTYPE Worksheet::get_Application( 
            /* [retval][out] */ Application	**RHS)
{
   TRACE_IN;             
   
   if ( m_p_application == NULL )
   {
       ERR( " m_p_application == NULL \n " ); 
       TRACE_OUT;
       return ( S_FALSE );    
   }
            
   HRESULT hr = S_OK;
   
   hr = (static_cast<Application*>( m_p_application ))->get_Application( RHS );          
             
   TRACE_OUT;
   return hr;                           
}

HRESULT STDMETHODCALLTYPE Worksheet::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;                             
}
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;             
    
   if ( m_p_parent == NULL )
   {
       ERR( " m_p_parent == NULL \n " ); 
       TRACE_OUT;
       return ( E_FAIL );    
   }    
            
   HRESULT hr = S_OK;
   
   hr = (static_cast<CSheets*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;                           
}
        
HRESULT STDMETHODCALLTYPE Worksheet::Activate( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                         
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Copy( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
HRESULT STDMETHODCALLTYPE Worksheet::Delete( 
            /* [lcid][in] */ long lcid)
{         
   TRACE_IN;
   HRESULT hr;
   
   BSTR name;
   IDispatch* p_disp = NULL;
   
   hr = get_Name( &name );
   if ( FAILED( hr ) )
   {
       ERR( " get_Name \n" );     
   }
   
   hr = get_Parent( &p_disp );
   if ( FAILED( hr ) )
   {
       ERR( " get_Parent \n" );     
   }   
   
   hr = reinterpret_cast<CSheets*>(p_disp)->RemoveWorksheetByName( name );
   if ( FAILED( hr ) )
   {
       ERR( " RemoveWorksheetByName \n" );     
   }    
   
   if ( p_disp != NULL )
   {
       p_disp->Release();
       p_disp = NULL;     
   }
   
   SysFreeString( name );
   
   TRACE_OUT;
   return ( hr );                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_CodeName( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get__CodeName( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put__CodeName( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Index( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Move( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Name( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_IN;
   HRESULT hr;
   BSTR result;
   
   result = SysAllocString( m_oo_sheet.getName( ) );

   if ( lstrlenW( result ) == 0 )
   {
       ERR( " m_oo_sheet.getName \n" );  
       hr = E_FAIL;   
   } else
   {
      *RHS = SysAllocString( result );       
   }

   SysFreeString( result );

   TRACE_OUT;
   return ( hr );                           
}
        
HRESULT STDMETHODCALLTYPE Worksheet::put_Name( 
            /* [in] */ BSTR RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   hr = m_oo_sheet.setName( SysAllocString( RHS ) );
   if ( FAILED( hr ) )
   {
       ERR( " m_oo_sheet.setName \n" );     
   }
   
   TRACE_OUT;
   return ( hr );                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Next( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_PageSetup( 
            /* [retval][out] */ PageSetup	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Previous( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::__PrintOut( 
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
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::_Protect( 
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT DrawingObjects,
            /* [optional][in] */ VARIANT Contents,
            /* [optional][in] */ VARIANT Scenarios,
            /* [optional][in] */ VARIANT UserInterfaceOnly,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ProtectContents( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ProtectDrawingObjects( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ProtectionMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ProtectScenarios( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::_SaveAs( 
            /* [in] */ BSTR Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Select( 
            /* [optional][in] */ VARIANT Replace,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
HRESULT STDMETHODCALLTYPE Worksheet::Unprotect( 
            /* [optional][in] */ VARIANT Password,
            /* [lcid][in] */ long lcid)
{
    TRACE_IN;
    HRESULT hr;
   
    CorrectArg(Password, &Password);
   	
	if ( Is_Variant_Null(Password) ) 
	{
	    hr = m_oo_sheet.unprotect( L"" );   	 
	    if ( FAILED( hr ) )
	    {
		    ERR( " _oo_sheet.unprotect \n" );   	 
        }
	    
    } else
	{
	    hr = m_oo_sheet.unprotect( V_BSTR( &Password ) );   	 
	    if ( FAILED( hr ) )
	    {
		    ERR( " _oo_sheet.unprotect \n" );   	 
        }	  	  
    }     	
			
    TRACE_OUT;
    return ( hr );                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlSheetVisibility *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlSheetVisibility RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Shapes( 
            /* [retval][out] */ Shapes **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_TransitionExpEval( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_TransitionExpEval( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Arcs( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_AutoFilterMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_AutoFilterMode( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::SetBackgroundPicture( 
            /* [in] */ BSTR Filename)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Buttons( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Calculate( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_EnableCalculation( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_EnableCalculation( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Cells( 
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ChartObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::CheckBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::CheckSpelling( 
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [optional][in] */ VARIANT AlwaysSuggest,
            /* [optional][in] */ VARIANT SpellLang,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_CircularReference( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ClearArrows( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Columns( 
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ConsolidationFunction( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlConsolidationFunction *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ConsolidationOptions( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ConsolidationSources( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_DisplayAutomaticPageBreaks( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_DisplayAutomaticPageBreaks( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Drawings( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::DrawingObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::DropDowns( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_EnableAutoFilter( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_EnableAutoFilter( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_EnableSelection( 
            /* [retval][out] */ XlEnableSelection *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_EnableSelection( 
            /* [in] */ XlEnableSelection RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_EnableOutlining( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_EnableOutlining( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_EnablePivotTable( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_EnablePivotTable( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::_Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_FilterMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ResetAllPageBreaks( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::GroupBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::GroupObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Labels( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Lines( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ListBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Names( 
            /* [retval][out] */ Names	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::OLEObjects( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_OnCalculate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_OnCalculate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_OnData( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_OnData( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::OptionButtons( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Outline( 
            /* [retval][out] */ Outline	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Ovals( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Paste( 
            /* [optional][in] */ VARIANT Destination,
            /* [optional][in] */ VARIANT Link,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::_PasteSpecial( 
            /* [optional][in] */ VARIANT Format,
            /* [optional][in] */ VARIANT Link,
            /* [optional][in] */ VARIANT DisplayAsIcon,
            /* [optional][in] */ VARIANT IconFileName,
            /* [optional][in] */ VARIANT IconIndex,
            /* [optional][in] */ VARIANT IconLabel,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Pictures( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::PivotTables( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::PivotTableWizard( 
            /* [optional][in] */ VARIANT SourceType,
            /* [optional][in] */ VARIANT SourceData,
            /* [optional][in] */ VARIANT TableDestination,
            /* [optional][in] */ VARIANT TableName,
            /* [optional][in] */ VARIANT RowGrand,
            /* [optional][in] */ VARIANT ColumnGrand,
            /* [optional][in] */ VARIANT SaveData,
            /* [optional][in] */ VARIANT HasAutoFormat,
            /* [optional][in] */ VARIANT AutoPage,
            /* [optional][in] */ VARIANT Reserved,
            /* [optional][in] */ VARIANT BackgroundQuery,
            /* [optional][in] */ VARIANT OptimizeCache,
            /* [optional][in] */ VARIANT PageFieldOrder,
            /* [optional][in] */ VARIANT PageFieldWrapCount,
            /* [optional][in] */ VARIANT ReadData,
            /* [optional][in] */ VARIANT Connection,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ PivotTable **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Range( 
            /* [in] */ VARIANT Cell1,
            /* [optional][in] */ VARIANT Cell2,
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Rectangles( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Rows( 
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Scenarios( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ScrollArea( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_ScrollArea( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ScrollBars( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ShowAllData( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ShowDataForm( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Spinners( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_StandardHeight( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_StandardWidth( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_StandardWidth( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::TextBoxes( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_TransitionFormEntry( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_TransitionFormEntry( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Type( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlSheetType *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_UsedRange( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_HPageBreaks( 
            /* [retval][out] */ HPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_VPageBreaks( 
            /* [retval][out] */ vPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_QueryTables( 
            /* [retval][out] */ QueryTables **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_DisplayPageBreaks( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_DisplayPageBreaks( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Comments( 
            /* [retval][out] */ Comments **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Hyperlinks( 
            /* [retval][out] */ HyperLinks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ClearCircles( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::CircleInvalid( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get__DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put__DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_AutoFilter( 
            /* [retval][out] */ AutoFilter **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_DisplayRightToLeft( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Scripts( 
            /* [retval][out] */ Scripts **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::_PrintOut( 
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
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::_CheckSpelling( 
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [optional][in] */ VARIANT AlwaysSuggest,
            /* [optional][in] */ VARIANT SpellLang,
            /* [optional][in] */ VARIANT IgnoreFinalYaa,
            /* [optional][in] */ VARIANT SpellScript,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Tab( 
            /* [retval][out] */ Tab **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_MailEnvelope( 
            /* [retval][out] */ MsoEnvelope **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::SaveAs( 
            /* [in] */ BSTR Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [optional][in] */ VARIANT Local)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_CustomProperties( 
            /* [retval][out] */ CustomProperties **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_SmartTags( 
            /* [retval][out] */ SmartTags **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Protection( 
            /* [retval][out] */ Protection **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::PasteSpecial( 
            /* [optional][in] */ VARIANT Format,
            /* [optional][in] */ VARIANT Link,
            /* [optional][in] */ VARIANT DisplayAsIcon,
            /* [optional][in] */ VARIANT IconFileName,
            /* [optional][in] */ VARIANT IconIndex,
            /* [optional][in] */ VARIANT IconLabel,
            /* [optional][in] */ VARIANT NoHTMLFormatting,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Protect( 
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT DrawingObjects,
            /* [optional][in] */ VARIANT Contents,
            /* [optional][in] */ VARIANT Scenarios,
            /* [optional][in] */ VARIANT UserInterfaceOnly,
            /* [optional][in] */ VARIANT AllowFormattingCells,
            /* [optional][in] */ VARIANT AllowFormattingColumns,
            /* [optional][in] */ VARIANT AllowFormattingRows,
            /* [optional][in] */ VARIANT AllowInsertingColumns,
            /* [optional][in] */ VARIANT AllowInsertingRows,
            /* [optional][in] */ VARIANT AllowInsertingHyperlinks,
            /* [optional][in] */ VARIANT AllowDeletingColumns,
            /* [optional][in] */ VARIANT AllowDeletingRows,
            /* [optional][in] */ VARIANT AllowSorting,
            /* [optional][in] */ VARIANT AllowFiltering,
            /* [optional][in] */ VARIANT AllowUsingPivotTables)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_ListObjects( 
            /* [retval][out] */ ListObjects **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::XmlDataQuery( 
            /* [in] */ BSTR XPath,
            /* [optional][in] */ VARIANT SelectionNamespaces,
            /* [optional][in] */ VARIANT Map,
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::XmlMapQuery( 
            /* [in] */ BSTR XPath,
            /* [optional][in] */ VARIANT SelectionNamespaces,
            /* [optional][in] */ VARIANT Map,
            /* [retval][out] */ Range	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::PrintOut( 
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
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_EnableFormatConditionsCalculation( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Worksheet::put_EnableFormatConditionsCalculation( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Worksheet::get_Sort( 
            /* [retval][out] */ Sort **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
        /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Worksheet::ExportAsFixedFormat( 
            /* [in] */ XlFixedFormatType Type,
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT Quality,
            /* [optional][in] */ VARIANT IncludeDocProperties,
            /* [optional][in] */ VARIANT IgnorePrintAreas,
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT OpenAfterPublish,
            /* [optional][in] */ VARIANT FixedFormatExtClassPtr)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}

HRESULT Worksheet::Init()
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID__Worksheet, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
    }
    
    return ( hr );     
}

HRESULT Worksheet::Put_Application( void* p_application )
{
    m_p_application = p_application;
            
    return S_OK;      
}

HRESULT Worksheet::Put_Parent( void* p_parent )
{
   m_p_parent = p_parent;
   
   return S_OK;     
}

HRESULT Worksheet::InitWrapper( OOSheet _oo_sheet)
{
    m_oo_sheet = _oo_sheet;      
}
