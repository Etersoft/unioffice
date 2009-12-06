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
#include "page_setup.h"
#include "workbook.h"
#include "outline.h"
#include "names.h"
#include "range.h"
#include "../OOWrappers/oo_document.h"
#include "../OOWrappers/oo_page_style.h"
#include "../OOWrappers/oo_page_styles.h"
#include "../OOWrappers/oo_style_families.h"
#include "../OOWrappers/oo_named_ranges.h"
#include "../OOWrappers/oo_controller.h"
#include "../OOWrappers/oo_range.h"

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
        
HRESULT STDMETHODCALLTYPE Worksheet::get_PageSetup( 
            /* [retval][out] */ PageSetup	**RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   CPageSetup* p_page_setup = new CPageSetup;
   
   p_page_setup->Put_Application( m_p_application );
   p_page_setup->Put_Parent( this );

   Workbook* workbook = NULL;
   
   hr = (static_cast<CSheets*>( m_p_parent ))->get_Parent( (IDispatch**)&workbook );   
         
   OODocument oo_document = workbook->GetOODocument( ); 

   if ( workbook != NULL )
   {
       workbook->Release( );
	   workbook = NULL;   	  	
   }

   BSTR style_name = m_oo_sheet.PageStyle( );

   OOStyleFamilies oo_style_families; 
   hr = oo_document.StyleFamilies( oo_style_families );  
   if ( FAILED( hr ) )
   {
       ERR( " oo_document.StyleFamilies \n" );
       if ( p_page_setup != NULL )
           p_page_setup->Release();
           
       SysFreeString( style_name );          
	   TRACE_OUT;
	   return ( hr );	  	
   }          
   
   OOPageStyles oo_page_styles;
   hr = oo_style_families.getPageStyles( oo_page_styles );
   if ( FAILED( hr ) )
   {
       ERR( " oo_style_families.getPageStyles \n" );
       if ( p_page_setup != NULL )
           p_page_setup->Release();
           
       SysFreeString( style_name );          
	   TRACE_OUT;
	   return ( hr );	  	
   }

   OOPageStyle oo_page_style;
   
   hr = S_OK;
   
   oo_page_style = oo_page_styles.getByName( style_name );
  
   if ( oo_page_style.IsNull() )
   {
   	   hr = E_FAIL;	
       ERR( " oo_page_styles.getByName \n" );
       if ( p_page_setup != NULL )
           p_page_setup->Release();
           
       SysFreeString( style_name );          
	   TRACE_OUT;
	   return ( hr );	  	
   }
   
   p_page_setup->InitWrapper( oo_page_style );
             
   hr = p_page_setup->QueryInterface( DIID_PageSetup, (void**)RHS );
             
   if ( FAILED( hr ) )
   {
       ERR( " page_setup.QueryInterface \n" );     
   }
             
   if ( p_page_setup != NULL )
       p_page_setup->Release();
   
   SysFreeString( style_name );
   
   TRACE_OUT;
   return ( hr );                           
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
        
HRESULT STDMETHODCALLTYPE Worksheet::Select( 
            /* [optional][in] */ VARIANT Replace,
            /* [lcid][in] */ long lcid)
{
   TRACE_IN;
   HRESULT hr = S_OK;
   
   CorrectArg(Replace, &Replace);
   
   hr = VariantChangeTypeEx(&Replace, &Replace, 0, 0, VT_BOOL);
   if ( FAILED( hr ) ) {
       ERR(" VariantChangeTypeEx   %08x\n", hr);
       TRACE_OUT;
       return ( hr );
   }

   if (V_BOOL(&Replace)==VARIANT_TRUE) 
   {  
       Workbook* workbook = NULL;
   
       hr = (static_cast<CSheets*>( m_p_parent ))->get_Parent( (IDispatch**)&workbook );   
       if ( FAILED ( hr ) )
       {
	       ERR( " get_Parent \n" );	  	
       }
	            
       OODocument oo_document = workbook->GetOODocument( ); 
              
       if ( workbook != NULL )
       {
           workbook->Release( );
	       workbook = NULL;   	  	
       }
       
       OOController  oo_controller;
       
       hr = oo_document.getCurrentController( oo_controller );
       if ( FAILED ( hr ) )
       {
	       ERR( " getCurrentController \n" );	  	
       }
       
       
       
   	  								   
   } else
   {
      ERR( " Get VARIANT_FALSE as parameter \n" );   	 	 
      // TODO:  
      // now set S_OK 
      hr = S_OK;
   }
   
   TRACE_OUT;
   return ( hr );                           
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
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlSheetVisibility *RHS)
{
   TRACE_IN;
   HRESULT hr = S_OK;
    
   VARIANT_BOOL b_visible;
    
   b_visible = m_oo_sheet.isVisible( );
   
   switch ( b_visible ) {
       case VARIANT_TRUE:
           *RHS = xlSheetVisible;
           break;
       case VARIANT_FALSE:
           *RHS = xlSheetHidden;
           break;
   }
   
   TRACE_OUT
   return ( hr );                           
}
        
HRESULT STDMETHODCALLTYPE Worksheet::put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlSheetVisibility RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   VARIANT_BOOL b_visible;
   
   switch ( RHS )
   {
        case xlSheetVeryHidden:
        case xlSheetHidden:
            b_visible = VARIANT_FALSE;
            break;
        case xlSheetVisible:
            b_visible = VARIANT_TRUE;
            break;
   }
   
   hr = m_oo_sheet.isVisible( b_visible );
   
   if ( FAILED( hr ) )
   {
       ERR( " m_oo_sheet.isVisible \n" );	  	
   }
   
   TRACE_OUT;
   return ( hr );                           
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
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Cells( 
            /* [retval][out] */ Range	**RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   CRange* p_range = new CRange;
   
   p_range->Put_Application( m_p_application );
   p_range->Put_Parent( this );
     
   OORange    oo_range;
   
   if ( OOVersion == VER_3 )
   {
   	  	oo_range = m_oo_sheet.getCellRangeByPosition( 0, 0, 1023, 65535 );
   } else 
   {
   	    oo_range = m_oo_sheet.getCellRangeByPosition( 0, 0, 255, 65535 ); 	  
   }
      
   if ( oo_range.IsNull() )
   {
       ERR( " m_oo_sheet.getCellRangeByPosition \n" );	  
	   
       if ( p_range != NULL )
           p_range->Release();	 
		     
	   TRACE_OUT;
	   return ( hr );	
   }
   
   p_range->InitWrapper( oo_range );
             
   hr = p_range->QueryInterface( DIID_Range, (void**)RHS );
             
   if ( FAILED( hr ) )
   {
       ERR( " p_range.QueryInterface \n" );     
   }
             
   if ( p_range != NULL )
       p_range->Release();
   
   TRACE_OUT;
   return ( hr );                          
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
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Names( 
            /* [retval][out] */ Names	**RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   CNames* p_names = new CNames;
   
   p_names->Put_Application( m_p_application );
   p_names->Put_Parent( this );
   
   Workbook* workbook = NULL;
   
   hr = (static_cast<CSheets*>( m_p_parent ))->get_Parent( (IDispatch**)&workbook );   
         
   OODocument oo_document = workbook->GetOODocument( ); 

   if ( workbook != NULL )
   {
       workbook->Release( );
	   workbook = NULL;   	  	
   } 
  
   OONamedRanges    oo_named_ranges;
   
   hr = oo_document.NamedRanges( oo_named_ranges );
   if ( FAILED( hr ) )
   {
       ERR( " oo_document.GetNamedRanges \n" );	  
	   
       if ( p_names != NULL )
           p_names->Release();	 
		     
	   TRACE_OUT;
	   return ( hr );	
   }
   
   p_names->InitWrapper( oo_named_ranges );
             
   hr = p_names->QueryInterface( DIID_Names, (void**)RHS );
             
   if ( FAILED( hr ) )
   {
       ERR( " p_names.QueryInterface \n" );     
   }
             
   if ( p_names != NULL )
       p_names->Release();
   
   TRACE_OUT;
   return ( hr );                          
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
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Outline( 
            /* [retval][out] */ Outline	**RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   COutline* p_outline = new COutline;
   
   p_outline->Put_Application( m_p_application );
   p_outline->Put_Parent( this );
   
   p_outline->InitWrapper( m_oo_sheet );
             
   hr = p_outline->QueryInterface( DIID_Outline, (void**)RHS );
             
   if ( FAILED( hr ) )
   {
       ERR( " p_outline.QueryInterface \n" );     
   }
             
   if ( p_outline != NULL )
       p_outline->Release();
   
   TRACE_OUT;
   return ( hr );                          
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
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Range( 
            /* [in] */ VARIANT Cell1,
            /* [optional][in] */ VARIANT Cell2,
            /* [retval][out] */ Range	**RHS)
{
   TRACE_IN;
   HRESULT hr;
   VARIANT vNull;   
   
   VariantInit(&vNull);
    
   CorrectArg(Cell1, &Cell1);
   CorrectArg(Cell2, &Cell2);
    
   if ( Is_Variant_Null( Cell2 ) ) 
   {
   	   VARIANT     var_range;	
   	   Range*      p_range = NULL;
	   IRange*     p_irange = NULL;	  	
   	  	
   	   VariantClear( &var_range );	
   	  	
       hr = get_Cells( &p_range );
       if ( FAILED( hr ) ) {
           ERR( " get_Cells \n" );
           TRACE_OUT;
           return ( hr );
       }
       
       hr = p_range->QueryInterface( IID_IRange, (void**)(&p_irange));
       if ( FAILED( hr ) ) {
           ERR( " p_range->QueryInterface \n" );
           
           if ( p_range != NULL )
           {
               p_range->Release();
               p_range = NULL;
		   }
		              
           TRACE_OUT;
           return ( hr );
       }
       
       hr = p_irange->get__Default( Cell1, vNull, 0, &var_range );
       if ( FAILED( hr ) ) {
           ERR(" get__Default \n" );
           
           if ( p_range != NULL )
           {
               p_range->Release();
               p_range = NULL;
		   }

           if ( p_irange != NULL )
           {
               p_irange->Release();
               p_irange = NULL;
		   }
		   
           TRACE_OUT;
           return ( hr );
       }
       
       hr = (V_DISPATCH( &var_range ))->QueryInterface( DIID_Range, (void**)RHS);
       if ( FAILED( hr ) )
       {
	       ERR( " QueryInterface \n ");
	       VariantClear( &var_range );
	       
	       if ( p_range != NULL )
 	       {
 	           p_range->Release();
 	           p_range = NULL;
		   }
		   
		   if ( p_irange != NULL )
 	       {
 	           p_irange->Release();
 	           p_irange = NULL;
		   }
		   
		   TRACE_OUT;
		   return ( hr );   	 
       }
       
       VariantClear( &var_range );
       
       if ( p_range != NULL )
 	   {
 	       p_range->Release();
 	       p_range = NULL;
	   }
      
	   if ( p_irange != NULL )
 	   {
 	       p_irange->Release();
 	       p_irange = NULL;
	   }
	   
       TRACE_OUT;
       return ( hr );
   } // if ( Is_Variant_Null( Cell2 ) )      
    
    
    
    
    
    
    
    
   
   TRACE_OUT;
   return ( hr );                           
}
        
        /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Worksheet::Rectangles( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
HRESULT STDMETHODCALLTYPE Worksheet::get_Rows( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_IN;
    HRESULT hr;
	Range*      p_range = NULL;
	IRange*     p_irange = NULL;	  	
   	  	
    hr = get_Cells( &p_range );
    if ( FAILED( hr ) ) 
	{
        ERR( " get_Cells \n" );
        TRACE_OUT;
        return ( hr );
    }
       
    hr = p_range->QueryInterface( IID_IRange, (void**)(&p_irange));
    if ( FAILED( hr ) ) 
	{
        ERR( " p_range->QueryInterface \n" );
        
        if ( p_range != NULL )
        {
            p_range->Release();
            p_range = NULL;
	    }
		              
        TRACE_OUT;
        return ( hr );
    }
	 
	hr = p_irange->get_Rows( RHS );
    if ( FAILED( hr ) ) {
        ERR(" get_Rows \n" );
          
        if ( p_range != NULL )
        {
            p_range->Release();
            p_range = NULL;
	    }

        if ( p_irange != NULL )
        {
            p_irange->Release();
            p_irange = NULL;
	    }
		   
        TRACE_OUT;
        return ( hr );
    }

    if ( p_range != NULL )
 	{
 	    p_range->Release();
 	    p_range = NULL;
	}
      
	if ( p_irange != NULL )
 	{
 	    p_irange->Release();
 	    p_irange = NULL;
    }	 
	 
    TRACE_OUT;
    return ( hr );                            
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
        
HRESULT STDMETHODCALLTYPE Worksheet::Protect( 
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
    TRACE_IN;
	HRESULT hr; 			
 	
    CorrectArg(Password,                 &Password);
    CorrectArg(DrawingObjects,           &DrawingObjects);
    CorrectArg(Contents,                 &Contents);
    CorrectArg(Scenarios,                &Scenarios);
    CorrectArg(UserInterfaceOnly,        &UserInterfaceOnly);
    CorrectArg(AllowFormattingCells,     &AllowFormattingCells);
    CorrectArg(AllowFormattingColumns,   &AllowFormattingColumns);
    CorrectArg(AllowFormattingRows,      &AllowFormattingRows);
    CorrectArg(AllowInsertingColumns,    &AllowInsertingColumns);
    CorrectArg(AllowInsertingRows,       &AllowInsertingRows);
    CorrectArg(AllowInsertingHyperlinks, &AllowInsertingHyperlinks);
    CorrectArg(AllowDeletingColumns,     &AllowDeletingColumns);
    CorrectArg(AllowDeletingRows,        &AllowDeletingRows);
    CorrectArg(AllowSorting,             &AllowSorting);
    CorrectArg(AllowFiltering,           &AllowFiltering);
    CorrectArg(AllowUsingPivotTables,    &AllowUsingPivotTables); 
	
	if ( Is_Variant_Null( Password ) ) 
	{
	    hr = m_oo_sheet.protect( L"" );   	 
	    if ( FAILED( hr ) )
	    {
		    ERR( " _oo_sheet.protect \n" );   	 
        }
	    
    } else
	{
	    hr = m_oo_sheet.protect( V_BSTR( &Password ) );   	 
	    if ( FAILED( hr ) )
	    {
		    ERR( " _oo_sheet.protect \n" );   	 
        }	  	  
    }  
		 		
    TRACE_OUT;
    return ( hr );                           
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
