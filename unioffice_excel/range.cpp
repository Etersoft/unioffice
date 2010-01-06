/*
 * implementation of Range
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

#include "range.h"

#include "application.h"
#include "worksheet.h"
#include "font.h"
#include "interior.h"
#include "borders.h"

#include "../OOWrappers/com/sun/star/table/cell_range_address.h"

using namespace com::sun::star::table;


const long CC_VALUE 	    = 1;
const long CC_DATETIME 		= 2;
const long CC_STRING 	    = 4;
const long CC_ANNOTATION 	= 8;
const long CC_FORMULA 	    = 16;
const long CC_HARDATTR 		= 32;
const long CC_STYLES 	    = 64;
const long CC_OBJECTS 	    = 128;
const long CC_EDITATTR 		= 256;
const long CC_FORMATTED 	= 512;

       // IUnknown
HRESULT STDMETHODCALLTYPE CRange::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<IRange*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<IRange*>(this));
    }     
    
    if ( iid == IID_IRange) {
        TRACE("IRange\n");
        *ppv = static_cast<IRange*>(this);
    } 
    
    if ( iid == DIID_Range ) {
        TRACE("Range \n");
        *ppv = static_cast<Range*>(this);
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
        
ULONG STDMETHODCALLTYPE CRange::AddRef( )
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef); 		
}
        
ULONG STDMETHODCALLTYPE CRange::Release( )
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
HRESULT STDMETHODCALLTYPE CRange::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK; 		
}
        
HRESULT STDMETHODCALLTYPE CRange::GetTypeInfo(
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
        
HRESULT STDMETHODCALLTYPE CRange::GetIDsOfNames(
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
    
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CRange::Invoke(
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
                 static_cast<IDispatch*>(static_cast<IRange*>(this)), 
                 dispIdMember, 
                 wFlags, 
                 pDispParams, 
                 pVarResult, 
                 pExcepInfo, 
                 puArgErr);       
            
    if ( FAILED(hr) )
    { 
        ERR( " dispIdMember = %i   hr = %08x \n", dispIdMember, hr ); 
	    ERR( " wFlags = %i  \n", wFlags );   
	    ERR( " pDispParams->cArgs = %i \n", pDispParams->cArgs );
    }  
	             
    return ( hr );  		
}
         
               
        // IRange       
HRESULT STDMETHODCALLTYPE CRange::get_Application( 
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
   
   _Application* p_application = NULL;
   
   hr = (static_cast<IUnknown*>( m_p_application ))->QueryInterface( IID__Application,(void**)(&p_application) ); 
   if ( FAILED( hr ) )
   {
       ERR( " IUnknown->QueryInterface \n" );
	   TRACE_OUT;
	   return ( hr );	  	
   }
   
   hr = p_application->get_Application( RHS );          
   
   if ( p_application != NULL )
   {
       p_application->Release();
	   p_application = NULL;	  	
   }
             
   TRACE_OUT;
   return hr;  		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Parent( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;             
    
   if ( m_p_parent == NULL )
   {
       ERR( " m_p_parent == NULL \n " ); 
       TRACE_OUT;
       return ( S_FALSE );    
   }    
            
   HRESULT hr = S_OK;
   
   hr = (static_cast<IUnknown*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;  		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Activate( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_AddIndent( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_AddIndent( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Address( 
            /* [optional][in] */ VARIANT RowAbsolute,
            /* [optional][in] */ VARIANT ColumnAbsolute,
            /* [defaultvalue][optional][in] */ XlReferenceStyle ReferenceStyle,
            /* [optional][in] */ VARIANT External,
            /* [optional][in] */ VARIANT RelativeTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_AddressLocal( 
            /* [optional][in] */ VARIANT RowAbsolute,
            /* [optional][in] */ VARIANT ColumnAbsolute,
            /* [defaultvalue][optional][in] */ XlReferenceStyle ReferenceStyle,
            /* [optional][in] */ VARIANT External,
            /* [optional][in] */ VARIANT RelativeTo,
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::AdvancedFilter( 
            /* [in] */ XlFilterAction Action,
            /* [optional][in] */ VARIANT CriteriaRange,
            /* [optional][in] */ VARIANT CopyToRange,
            /* [optional][in] */ VARIANT Unique,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ApplyNames( 
            /* [optional][in] */ VARIANT Names,
            /* [optional][in] */ VARIANT IgnoreRelativeAbsolute,
            /* [optional][in] */ VARIANT UseRowColumnNames,
            /* [optional][in] */ VARIANT OmitColumn,
            /* [optional][in] */ VARIANT OmitRow,
            /* [defaultvalue][optional][in] */ XlApplyNamesOrder Order,
            /* [optional][in] */ VARIANT AppendLast,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ApplyOutlineStyles( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Areas( 
            /* [retval][out] */ Areas **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::AutoComplete( 
            /* [in] */ BSTR String,
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::AutoFill( 
            /* [in] */ Range	*Destination,
            /* [defaultvalue][optional][in] */ XlAutoFillType Type,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::AutoFilter( 
            /* [optional][in] */ VARIANT Field,
            /* [optional][in] */ VARIANT Criteria1,
            /* [defaultvalue][optional][in] */ XlAutoFilterOperator Operator,
            /* [optional][in] */ VARIANT Criteria2,
            /* [optional][in] */ VARIANT VisibleDropDown,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::AutoFit( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE CRange::AutoFormat( 
            /* [defaultvalue][optional][in] */ XlRangeAutoFormat Format,
            /* [optional][in] */ VARIANT Number,
            /* [optional][in] */ VARIANT Font,
            /* [optional][in] */ VARIANT Alignment,
            /* [optional][in] */ VARIANT Border,
            /* [optional][in] */ VARIANT Pattern,
            /* [optional][in] */ VARIANT Width,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::AutoOutline( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::BorderAround( 
            /* [optional][in] */ VARIANT LineStyle,
            /* [defaultvalue][optional][in] */ XlBorderWeight Weight,
            /* [defaultvalue][optional][in] */ XlColorIndex ColorIndex,
            /* [optional][in] */ VARIANT Color,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Borders( 
            /* [retval][out] */ Borders	**RHS)
{
    TRACE_IN;
    HRESULT hr;
    
	CBorders* p_borders = new CBorders;
   
   	p_borders->Put_Application( m_p_application );
	p_borders->Put_Parent( this );
   
    OOBorders    oo_borders;
   	 					
	oo_borders = m_oo_range;						
										     
	p_borders->InitWrapper( oo_borders );
             
   	hr = p_borders->QueryInterface( DIID_Borders, (void**)(RHS) );
             
    if ( FAILED( hr ) )
	{
	    ERR( " p_borders.QueryInterface \n" );     
	}
             
	if ( p_borders != NULL )
	{
	    p_borders->Release();
	    p_borders = NULL;
	}
	 
    TRACE_OUT;
    return ( hr );				
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Calculate( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Cells( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Characters( 
            /* [optional][in] */ VARIANT Start,
            /* [optional][in] */ VARIANT Length,
            /* [retval][out] */ Characters **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::CheckSpelling( 
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [optional][in] */ VARIANT AlwaysSuggest,
            /* [optional][in] */ VARIANT SpellLang,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::Clear( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    VariantInit( RHS );
    
    hr = m_oo_range.clearContents( 
	   CC_VALUE + 
	   CC_DATETIME + 
	   CC_STRING + 
	   CC_ANNOTATION + 
	   CC_FORMULA + 
	   CC_HARDATTR + 
	   CC_STYLES + 
	   CC_OBJECTS + 
	   CC_EDITATTR + 
	   CC_FORMATTED );
    
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.clearContents \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 			
}
        
HRESULT STDMETHODCALLTYPE CRange::ClearContents( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    VariantInit( RHS );
    
    hr = m_oo_range.clearContents( 
	   CC_VALUE + 
	   CC_DATETIME + 
	   CC_STRING + 
	   CC_FORMULA + 
	   CC_OBJECTS );
    
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.clearContents \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::ClearFormats( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    VariantInit( RHS );
    
    hr = m_oo_range.clearContents( 
	   CC_HARDATTR + 
	   CC_STYLES + 
	   CC_EDITATTR + 
	   CC_FORMATTED );
    
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.clearContents \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::ClearNotes( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    VariantInit( RHS );
    
    hr = m_oo_range.clearContents( CC_ANNOTATION  );
    
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.clearContents \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::ClearOutline( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    VariantInit( RHS );
    
    hr = m_oo_range.clearContents( CC_STYLES  );
    
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.clearContents \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Column( 
            /* [retval][out] */ long *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    CellRangeAddress  cell_range_address;
    
    cell_range_address = m_oo_range.getRangeAddress();
    if ( cell_range_address.IsNull() )
    {
	    ERR( " getRangeAddress \n" );  
		TRACE_OUT;
		return ( E_FAIL ); 	 
    }
    
    hr = S_OK;
    *RHS = cell_range_address.StartColumn() + 1;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ColumnDifferences( 
            /* [in] */ VARIANT Comparison,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Columns( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_IN;
    HRESULT hr;
    
	CRange* p_range = new CRange;
   
   	p_range->Put_Application( m_p_application );
	p_range->Put_Parent( m_p_parent );
   	 								     
	p_range->InitWrapper( m_oo_range );
             
   	hr = p_range->QueryInterface( DIID_Range, (void**)(RHS) );
             
    if ( FAILED( hr ) )
	{
	    ERR( " p_range.QueryInterface \n" );     
	}
             
	if ( p_range != NULL )
	{
	    p_range->Release();
	    p_range = NULL;
	}
	 
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_ColumnWidth( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT  hr;
    long width = 0;
    OORange oo_columns;
    
    oo_columns = m_oo_range.getColumns();
    if ( oo_columns.IsNull( ) )
    {
	    ERR( " getColumns \n" ); 
		TRACE_OUT;
		return ( E_FAIL );  	 
    }
    
    hr = oo_columns.getWidth( width );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_columns.getWidth \n" );
		TRACE_OUT;
		return ( hr );   	 
	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = width / 200;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_ColumnWidth( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT  hr;
    long width = 0;
    OORange oo_columns;
        
    CorrectArg(RHS, &RHS);    
        
    oo_columns = m_oo_range.getColumns();
    if ( oo_columns.IsNull( ) )
    {
	    ERR( " getColumns \n" ); 
		TRACE_OUT;
		return ( E_FAIL );  	 
    }
    
    VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    
    width = V_I4( &RHS ) * 210;
    
    hr = oo_columns.setWidth( width );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_columns.setWidth \n" );
		TRACE_OUT;
		return ( hr );   	 
	}
    
    TRACE_OUT;
    return ( hr );		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Consolidate( 
            /* [optional][in] */ VARIANT Sources,
            /* [optional][in] */ VARIANT Function,
            /* [optional][in] */ VARIANT TopRow,
            /* [optional][in] */ VARIANT LeftColumn,
            /* [optional][in] */ VARIANT CreateLinks,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Copy( 
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::CopyFromRecordset( 
            /* [in] */ IUnknown *Data,
            /* [optional][in] */ VARIANT MaxRows,
            /* [optional][in] */ VARIANT MaxColumns,
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::CopyPicture( 
            /* [defaultvalue][optional][in] */ XlPictureAppearance Appearance,
            /* [defaultvalue][optional][in] */ XlCopyPictureFormat Format,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Count( 
            /* [retval][out] */ long *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    CellRangeAddress  cell_range_address;
    
    cell_range_address = m_oo_range.getRangeAddress();
    if ( cell_range_address.IsNull() )
    {
	    ERR( " getRangeAddress \n" );  
		TRACE_OUT;
		return ( E_FAIL ); 	 
    }
    
    hr = S_OK;
    
    long width = cell_range_address.EndColumn() - cell_range_address.StartColumn() + 1;
    long height = cell_range_address.EndRow() - cell_range_address.StartRow() + 1;
    
    *RHS = width * height;
    
    TRACE_OUT;
    return ( hr );  		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::CreateNames( 
            /* [optional][in] */ VARIANT Top,
            /* [optional][in] */ VARIANT Left,
            /* [optional][in] */ VARIANT Bottom,
            /* [optional][in] */ VARIANT Right,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE CRange::CreatePublisher( 
            /* [optional][in] */ VARIANT Edition,
            /* [defaultvalue][optional][in] */ XlPictureAppearance Appearance,
            /* [optional][in] */ VARIANT ContainsPICT,
            /* [optional][in] */ VARIANT ContainsBIFF,
            /* [optional][in] */ VARIANT ContainsRTF,
            /* [optional][in] */ VARIANT ContainsVALU,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_CurrentArray( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_CurrentRegion( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Cut( 
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::DataSeries( 
            /* [optional][in] */ VARIANT Rowcol,
            /* [defaultvalue][optional][in] */ XlDataSeriesType Type,
            /* [defaultvalue][optional][in] */ XlDataSeriesDate Date,
            /* [optional][in] */ VARIANT Step,
            /* [optional][in] */ VARIANT Stop,
            /* [optional][in] */ VARIANT Trend,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get__Default( 
            /* [optional][in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr = E_NOTIMPL;
    int parameters_amount = 0;
    
    CorrectArg(RowIndex, &RowIndex);
    CorrectArg(ColumnIndex, &ColumnIndex);
    
    if ( !Is_Variant_Null( RowIndex ) )
        parameters_amount++;
        
    if ( !Is_Variant_Null( ColumnIndex ) )
        parameters_amount++;        
        
    switch ( parameters_amount )
    {
	    case 0:
			 {
			     ERR( " parameters_amount == 0 \n" );
				 hr = E_FAIL;  				   
	         }   
	         break;
	         
	    case 1:
			 {
			     if ( V_VT(&RowIndex) == VT_BSTR ) 
		 	     {
				     CRange* p_range = new CRange;
   
		 			 p_range->Put_Application( m_p_application );
   					 p_range->Put_Parent( m_p_parent );
     
   	 				 OORange    oo_range;
   	 				 
					 oo_range = m_oo_range.getCellRangeByName( V_BSTR( &RowIndex ) );
					 
					 hr = S_OK;
					 
					 if ( m_oo_range.IsNull() )
					 {
					  	  ERR( " failed m_oo_range.getCellRangeByName \n" );
						  hr = E_FAIL;	  
				     } 
				     
				     p_range->InitWrapper( oo_range );
             
                     V_VT( RHS ) = VT_DISPATCH;
   			 		 hr = p_range->QueryInterface( DIID_Range, (void**)(&(V_DISPATCH( RHS ))) );
             
   			 		 if ( FAILED( hr ) )
				 	 {
       				  	 ERR( " p_range.QueryInterface \n" );     
	                 }
             
   			 		 if ( p_range != NULL )
   			 		 {
      		  		     p_range->Release();
      		  		     p_range = NULL;
					 } 				 
                 } else
                 {
			         ERR( " now not realized VT(RowIndex) = %i \n", V_VT( &RowIndex ) );
				     hr = E_FAIL; 	   
				 }
				 				  	   
             }     
             break;
             
        case 2:
		     {
			   	 hr = VariantChangeTypeEx(&RowIndex, &RowIndex, 0, 0, VT_I4);  
			 	 hr = VariantChangeTypeEx(&ColumnIndex, &ColumnIndex, 0, 0, VT_I4);
			 	 
			 	 if ( ( V_VT( &RowIndex ) == VT_I4 ) && ( V_VT( &ColumnIndex ) == VT_I4 ) )
			 	 {
				  	 // we need to sub 1, because
				  	 // OpenOffice start numeration from 0, but MSOffice from 1
                     long row = V_I4( &RowIndex ) - 1;
					 long column = V_I4( &ColumnIndex ) - 1;
					 
					 CRange* p_range = new CRange;
   
		 			 p_range->Put_Application( m_p_application );
   					 p_range->Put_Parent( m_p_parent );
     
   	 				 OORange    oo_range;
   	 				 
					 oo_range = m_oo_range.getCellRangeByPosition( column, row, column, row );
					 
					 hr = S_OK;
					 
					 if ( m_oo_range.IsNull() )
					 {
					  	  ERR( " failed m_oo_range.getCellByPosition \n" );
						  hr = E_FAIL;	  
				     } 
				     
				     p_range->InitWrapper( oo_range );
             
                     V_VT( RHS ) = VT_DISPATCH;
   			 		 hr = p_range->QueryInterface( DIID_Range, (void**)(&(V_DISPATCH( RHS ))) );
             
   			 		 if ( FAILED( hr ) )
				 	 {
       				  	 ERR( " p_range.QueryInterface \n" );     
	                 }
             
   			 		 if ( p_range != NULL )
      		  		 {
      		  		     p_range->Release();
      		  		     p_range = NULL;
					 }

			     } else
				 {
				     ERR( " not supported type of parameter V_VT(RowIndex) == %i \n ", V_VT(&RowIndex) );
				     ERR( " not supported type of parameter V_VT(ColumnIndex) == %i \n ", V_VT(&ColumnIndex) );
				     hr = E_FAIL;   	   
		         }			   			   
		     }
		     break;
    } // switch ( parameters_amount ) 
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put__Default( 
            /* [optional][in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    VARIANT vNull;
    HRESULT hr;
    
    CorrectArg(RowIndex, &RowIndex);
    CorrectArg(ColumnIndex, &ColumnIndex);
    
    IRange   *p_range = NULL;
 	VARIANT  var_range;
    
    VariantClear( &var_range );
    
    hr = get__Default( RowIndex, ColumnIndex, lcid, &var_range );
    if ( FAILED( hr ) )
    {
	    ERR( " get__Default \n" );
	    VariantClear( &var_range );
		TRACE_OUT;
		return ( hr );   	 
	}
    
    hr = (V_DISPATCH( &var_range ))->QueryInterface( IID_IRange, (void**)(&p_range));
    if ( FAILED( hr ) )
    {
	    ERR( " QueryInterface \n ");
	    VariantClear( &var_range );
	    if ( p_range != NULL )
 	    {
 	        p_range->Release();
 	        p_range = NULL;
		}
		TRACE_OUT;
		return ( hr );   	 
    }
    
    VariantClear( &var_range );
    
    VariantInit( &vNull );
    V_VT( &vNull ) = VT_NULL;
    
    hr = p_range->put_Value( vNull, 0, RHS );
    if ( FAILED( hr ) )
    {
	   ERR( " put_Value \n" );	 
	}
    
    if ( p_range != NULL )
    {
 	    p_range->Release();
 	    p_range = NULL;
	}
	    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Delete( 
            /* [optional][in] */ VARIANT Shift,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Dependents( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::DialogBox( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_DirectDependents( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_DirectPrecedents( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::EditionOptions( 
            /* [in] */ XlEditionType Type,
            /* [in] */ XlEditionOptionsOption Option,
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT Reference,
            /* [defaultvalue][optional][in] */ XlPictureAppearance Appearance,
            /* [defaultvalue][optional][in] */ XlPictureAppearance ChartSize,
            /* [optional][in] */ VARIANT Format,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_End( 
            /* [in] */ XlDirection Direction,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_EntireColumn( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_EntireRow( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    OORange oo_rows;
    
    oo_rows = m_oo_range.getRows();
    if ( oo_rows.IsNull() )
    {
 	    ERR(" m_oo_range.getRows \n");
		TRACE_OUT;
		return ( E_FAIL );  
	}
    
    CellRangeAddress   cell_range_address;
    
    cell_range_address = m_oo_range.getRangeAddress();
    if ( cell_range_address.IsNull() )
    {
	    ERR( " getRangeAddress \n" );  
		TRACE_OUT;
		return ( E_FAIL ); 	 
    }

	long start_row = cell_range_address.StartRow();
	long end_row = cell_range_address.EndRow();

	if ( (start_row < 0) || (end_row < 0) )
	{
	    ERR( " (start_row < 0) || (end_row < 0) \n" ); 
		TRACE_OUT;
		return ( E_FAIL );  	 
    }

    OOSheet oo_sheet = getParentOOSheet();
	if ( oo_sheet.IsNull() )
	{
	    ERR( " oo_sheet.IsNull() \n" );  
		TRACE_OUT;
		return ( E_FAIL ); 	 
    }

	OORange oo_range;
	
	switch ( OOVersion )
    {
       case VER_3:
	   		{
			    oo_range = oo_sheet.getCellRangeByPosition( 0, start_row, 1023, end_row ); 	  
		 	}
		 	break;
		 	
       case VER_2:
	   		{
			    oo_range = oo_sheet.getCellRangeByPosition( 0, start_row, 255, end_row );			 	  
		 	}
		 	break;		 	
		 	
	   default:
	   		{
			    oo_range = oo_sheet.getCellRangeByPosition( 0, start_row, 255, end_row );			   				
			}	 	
   		    break;
   		    
    } // switch ( OOVersion )

	CRange* p_range = new CRange;
   
   	p_range->Put_Application( m_p_application );
	p_range->Put_Parent( m_p_parent );
				     
	p_range->InitWrapper( oo_range );
             
   	hr = p_range->QueryInterface( DIID_Range, (void**)(RHS) );
             
    if ( FAILED( hr ) )
	{
	    ERR( " p_range.QueryInterface \n" );     
	}
             
	if ( p_range != NULL )
	{
	    p_range->Release();
	    p_range = NULL;
	}

    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::FillDown( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::FillLeft( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::FillRight( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::FillUp( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Find( 
            /* [in] */ VARIANT What,
            /* [optional][in] */ VARIANT After,
            /* [optional][in] */ VARIANT LookIn,
            /* [optional][in] */ VARIANT LookAt,
            /* [optional][in] */ VARIANT SearchOrder,
            /* [defaultvalue][optional][in] */ XlSearchDirection SearchDirection,
            /* [optional][in] */ VARIANT MatchCase,
            /* [optional][in] */ VARIANT MatchByte,
            /* [optional][in] */ VARIANT SearchFormat,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::FindNext( 
            /* [optional][in] */ VARIANT After,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::FindPrevious( 
            /* [optional][in] */ VARIANT After,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Font( 
            /* [retval][out] */ /* external definition not present */ Font **RHS)
{
    TRACE_IN;
    HRESULT hr;
    
	CFont* p_font = new CFont;
   
   	p_font->Put_Application( m_p_application );
	p_font->Put_Parent( this );
   
    OOFont    oo_font;
   	 					
	oo_font = m_oo_range;						
										     
	p_font->InitWrapper( oo_font );
             
   	hr = p_font->QueryInterface( DIID_Font, (void**)(RHS) );
             
    if ( FAILED( hr ) )
	{
	    ERR( " p_font.QueryInterface \n" );     
	}
             
	if ( p_font != NULL )
	{
	    p_font->Release();
	    p_font = NULL;
	}
	 
    TRACE_OUT;
    return ( hr );		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Formula( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
        
    OORange  oo_range;
		
    // get first cell in range
    oo_range = m_oo_range.getCellByPosition( 0, 0 );
     								   
	hr = oo_range.getFormula( RHS );
	
	if ( FAILED ( hr ) )
	{
		ERR( " failed oo_range.getFormula \n" ); 	  
	}			  												   
    
    TRACE_OUT;
    return ( hr );	
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_Formula( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    VARIANT vNull;
    HRESULT hr;
    
    VariantInit( &vNull );
    V_VT( &vNull ) = VT_NULL;
    
    hr = put_Value( vNull, 0, RHS );
    if ( FAILED( hr ) )
    {
	   ERR( " put_Value \n" );	 
	}
    
    TRACE_OUT;
    return ( hr );		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_FormulaArray( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_FormulaArray( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_FormulaLabel( 
            /* [retval][out] */ XlFormulaLabel *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_FormulaLabel( 
            /* [in] */ XlFormulaLabel RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_FormulaHidden( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_FormulaHidden( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_FormulaLocal( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_FormulaLocal( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_FormulaR1C1( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_FormulaR1C1( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_FormulaR1C1Local( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_FormulaR1C1Local( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::FunctionWizard( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE CRange::GoalSeek( 
            /* [in] */ VARIANT Goal,
            /* [in] */ Range	*ChangingCell,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Group( 
            /* [optional][in] */ VARIANT Start,
            /* [optional][in] */ VARIANT End,
            /* [optional][in] */ VARIANT By,
            /* [optional][in] */ VARIANT Periods,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_HasArray( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_HasFormula( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Height( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::awt::Size  oo_size;
    long count = 0;
    long value = 0;
    
    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
	    ERR( " get_Count \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }    
    
    if ( count == 1 )
        oo_size = (static_cast<SheetCell>(m_oo_range)).getSize();
    else
	    oo_size = (static_cast<SheetCellRange>(m_oo_range)).getSize();
     
    if ( oo_size.IsNull() )
    {
	    ERR( " oo_size.IsNull \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }
    
    hr = oo_size.getHeight( value );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_size.getHeight \n" );
		TRACE_OUT;
		return ( E_FAIL ); 	     	 
	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = value;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Hidden( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Hidden( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_HorizontalAlignment( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::table::CellHoriJustify value = com::sun::star::table::HORI_STANDARD;
    		
	hr = m_oo_range.getHoriJustify( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_range.getHoriJustify \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
        
    switch ( value ) {
    case com::sun::star::table::HORI_STANDARD: 
		 V_I4( RHS ) = xlHAlignGeneral;
		 break;
		 
    case com::sun::star::table::HORI_LEFT:
		 V_I4( RHS ) = xlHAlignLeft;
		 break;
		 
    case com::sun::star::table::HORI_CENTER:
		 V_I4( RHS ) = xlHAlignCenter;
		 break;
		 
    case com::sun::star::table::HORI_RIGHT:
		 V_I4( RHS ) = xlHAlignRight;
		 break;
		 
    case com::sun::star::table::HORI_BLOCK:
		 V_I4( RHS ) = xlHAlignJustify;
		 break;
		 
    case com::sun::star::table::HORI_REPEAT:
		 V_I4( RHS ) =xlHAlignFill;
		 break;
		 
    default:
			V_I4( RHS ) = xlHAlignGeneral;
			break;
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_HorizontalAlignment( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::table::CellHoriJustify value = com::sun::star::table::HORI_STANDARD;
    
    CorrectArg(RHS, &RHS);
    
    VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    
    switch ( V_I4( &RHS ) ) {
	case xlHAlignGeneral: 
		 value = com::sun::star::table::HORI_STANDARD;
		 break;
		 
    case xlHAlignLeft:
		 value = com::sun::star::table::HORI_LEFT;
		 break;
		 
    case xlHAlignCenter:
		 value = com::sun::star::table::HORI_CENTER;
		 break;
		 
    case xlHAlignRight:
		 value = com::sun::star::table::HORI_RIGHT;
		 break;
		 
    case xlHAlignJustify:
		 value = com::sun::star::table::HORI_BLOCK;
		 break;
		 
    case xlHAlignFill:
		 value = com::sun::star::table::HORI_REPEAT;
		 break;
		 
    default:
 			value = com::sun::star::table::HORI_STANDARD;	
			break; 		
    } // switch
    
    
    hr = m_oo_range.setHoriJustify( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_range.setHoriJustify \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_IndentLevel( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_IndentLevel( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::InsertIndent( 
            /* [in] */ long InsertAmount)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Insert( 
            /* [optional][in] */ VARIANT Shift,
            /* [optional][in] */ VARIANT CopyOrigin,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Interior( 
            /* [retval][out] */ Interior	**RHS)
{
    TRACE_IN;
    HRESULT hr;
    
	CInterior* p_interior = new CInterior;
   
   	p_interior->Put_Application( m_p_application );
	p_interior->Put_Parent( this );
   
    OOInterior    oo_interior;
   	 					
	oo_interior = m_oo_range;						
										     
	p_interior->InitWrapper( oo_interior );
             
   	hr = p_interior->QueryInterface( DIID_Interior, (void**)(RHS) );
             
    if ( FAILED( hr ) )
	{
	    ERR( " p_interior.QueryInterface \n" );     
	}
             
	if ( p_interior != NULL )
	{
	    p_interior->Release();
	    p_interior = NULL;
	}
	 
    TRACE_OUT;
    return ( hr );		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Item( 
            /* [in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Item( 
            /* [in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Justify( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Left( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::awt::Point  oo_point;
    long count = 0;
    long value = 0;
    
    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
	    ERR( " get_Count \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }    
    
    if ( count == 1 )
        oo_point = (static_cast<SheetCell>(m_oo_range)).getPosition();
    else
	    oo_point = (static_cast<SheetCellRange>(m_oo_range)).getPosition();
     
    if ( oo_point.IsNull() )
    {
	    ERR( " oo_point.IsNull \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }
    
    hr = oo_point.getX( value );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_point.getX \n" );
		TRACE_OUT;
		return ( E_FAIL ); 	     	 
	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = value;
    
    TRACE_OUT;
    return ( hr );		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ListHeaderRows( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ListNames( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_LocationInTable( 
            /* [retval][out] */ XlLocationInTable *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Locked( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Locked( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::Merge( 
            /* [optional][in] */ VARIANT Across)
{
    TRACE_IN;
    HRESULT hr;
    
    CorrectArg( Across, &Across );
    
    VariantChangeTypeEx(&Across, &Across, 0, 0, VT_BOOL);
    
    if ( V_BOOL( &Across ) == VARIANT_FALSE )
	{
	    hr = m_oo_range.merge( true ); 
		if ( FAILED( hr ) )
		{
		    ERR( " m_oo_range.merge \n" );   	 
		    TRACE_OUT;
		    return ( hr );
	    }  	 
    } else
	{
	    CellRangeAddress  cell_range_address;
    
        cell_range_address = m_oo_range.getRangeAddress();
    	if ( cell_range_address.IsNull() )
    	{
	       ERR( " getRangeAddress \n" );  
		   TRACE_OUT;
		   return ( E_FAIL ); 	 
	    }
    
    	long startrow    = cell_range_address.StartRow();
		long endrow      = cell_range_address.EndRow();
		long startcolumn = cell_range_address.StartColumn();
		long endcolumn   = cell_range_address.EndColumn();  
		
		if (   
		   ( startrow < 0 ) ||
		   ( endrow < 0 ) ||
		   ( startcolumn < 0 ) ||
		   ( endcolumn < 0 )
		) 
		{
		    ERR( " failed when get start and end parameters \n " );
			TRACE_OUT;
			return ( E_FAIL );   	 
	    }
		
		for ( int i = 0; i <= endrow-startrow; i++ )
		{
		 	OORange oo_range;
		 	
		 	oo_range = m_oo_range.getCellRangeByPosition( 0, i, endcolumn - startcolumn, i );	
		 	
		 	if ( oo_range.IsNull() )
		 	{
			    ERR( " oo_range.IsNull() i = %i \n", i );
				TRACE_OUT;
				return ( E_FAIL );   	 
			}
		 	
		 	hr = oo_range.merge( true );
		 	
		 	if ( FAILED( hr ) )
			{
		       ERR( " m_oo_range.merge i = %i \n", i );   	 
		       TRACE_OUT;
		       return ( hr );
	    	} 
	 	
		} // for ( int i = 0; i <= endrow-startrow; i++ )	  
    }       
    
    TRACE_OUT;
    return ( S_OK ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::UnMerge( void)
{
    TRACE_IN;
    HRESULT hr;
    
    hr = m_oo_range.merge( false );
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.merge \n" );      	 
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_MergeArea( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_MergeCells( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    bool value = false;
    
    hr = m_oo_range.getIsMerged( value );
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.getIsMerged \n " ); 
		TRACE_OUT;
		return ( hr );  	 
    }
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_BOOL;
    
    if ( value )
        V_BOOL( RHS ) = VARIANT_TRUE;
    else
        V_BOOL( RHS ) = VARIANT_FALSE;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_MergeCells( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    CorrectArg(RHS, &RHS);
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    
    if ( V_BOOL( &RHS ) == VARIANT_FALSE) 
	{
        hr = UnMerge();
        
        if ( FAILED( hr ) )
    	{
	       ERR( " UnMerge \n " );   	 
		}
	
    } else 
	{
	    VARIANT param1;
		
		VariantInit( &param1 );
		
		V_VT( &param1 ) = VT_BOOL;
		V_BOOL( &param1 ) = VARIANT_TRUE; 	   
		
        hr = Merge( param1 );
        
        if ( FAILED( hr ) )
    	{
	       ERR( " Merge \n " );   	 
		}       
    }
        
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Name( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Name( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::NavigateArrow( 
            /* [optional][in] */ VARIANT TowardPrecedent,
            /* [optional][in] */ VARIANT ArrowNumber,
            /* [optional][in] */ VARIANT LinkNumber,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Next( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::NoteText( 
            /* [optional][in] */ VARIANT Text,
            /* [optional][in] */ VARIANT Start,
            /* [optional][in] */ VARIANT Length,
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_NumberFormat( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_NumberFormat( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_NumberFormatLocal( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_NumberFormatLocal( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Offset( 
            /* [optional][in] */ VARIANT RowOffset,
            /* [optional][in] */ VARIANT ColumnOffset,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Orientation( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Orientation( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_OutlineLevel( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_OutlineLevel( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_PageBreak( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_PageBreak( 
            /* [in] */ long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Parse( 
            /* [optional][in] */ VARIANT ParseLine,
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE CRange::_PasteSpecial( 
            /* [defaultvalue][optional][in] */ XlPasteType Paste,
            /* [defaultvalue][optional][in] */ XlPasteSpecialOperation Operation,
            /* [optional][in] */ VARIANT SkipBlanks,
            /* [optional][in] */ VARIANT Transpose,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_PivotField( 
            /* [retval][out] */ PivotField **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_PivotItem( 
            /* [retval][out] */ PivotItem **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_PivotTable( 
            /* [retval][out] */ PivotTable **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Precedents( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_PrefixCharacter( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Previous( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE CRange::__PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_QueryTable( 
            /* [retval][out] */ QueryTable **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Range( 
            /* [in] */ VARIANT Cell1,
            /* [optional][in] */ VARIANT Cell2,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::RemoveSubtotal( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Replace( 
            /* [in] */ VARIANT What,
            /* [in] */ VARIANT Replacement,
            /* [optional][in] */ VARIANT LookAt,
            /* [optional][in] */ VARIANT SearchOrder,
            /* [optional][in] */ VARIANT MatchCase,
            /* [optional][in] */ VARIANT MatchByte,
            /* [optional][in] */ VARIANT SearchFormat,
            /* [optional][in] */ VARIANT ReplaceFormat,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Resize( 
            /* [optional][in] */ VARIANT RowSize,
            /* [optional][in] */ VARIANT ColumnSize,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Row( 
            /* [retval][out] */ long *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    CellRangeAddress  cell_range_address;
    
    cell_range_address = m_oo_range.getRangeAddress();
    if ( cell_range_address.IsNull() )
    {
	    ERR( " getRangeAddress \n" );  
		TRACE_OUT;
		return ( E_FAIL ); 	 
    }
    
    hr = S_OK;
    *RHS = cell_range_address.StartRow() + 1;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::RowDifferences( 
            /* [in] */ VARIANT Comparison,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_RowHeight( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT  hr;
    long height = 0;
    OORange oo_rows;
    
    oo_rows = m_oo_range.getRows();
    if ( oo_rows.IsNull( ) )
    {
	    ERR( " getRows \n" ); 
		TRACE_OUT;
		return ( E_FAIL );  	 
    }
    
    hr = oo_rows.getHeight( height );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_rows.getHeight \n" );
		TRACE_OUT;
		return ( hr );   	 
	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = height / 1000 * 28;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_RowHeight( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT  hr;
    long height = 0;
    OORange oo_rows;
    
    CorrectArg(RHS, &RHS);
    
    oo_rows = m_oo_range.getRows();
    if ( oo_rows.IsNull( ) )
    {
	    ERR( " getRows \n" ); 
		TRACE_OUT;
		return ( E_FAIL );  	 
    }
    
    VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    
	height = V_I4( &RHS ) / 28 * 1000;
    
    hr = oo_rows.setHeight( height );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_rows.setHeight \n" );
		TRACE_OUT;
		return ( hr );   	 
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Rows( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_IN;
    HRESULT hr;
    
	CRange* p_range = new CRange;
   
   	p_range->Put_Application( m_p_application );
	p_range->Put_Parent( m_p_parent );
				     
	p_range->InitWrapper( m_oo_range );
             
   	hr = p_range->QueryInterface( DIID_Range, (void**)(RHS) );
             
    if ( FAILED( hr ) )
	{
	    ERR( " p_range.QueryInterface \n" );     
	}
             
	if ( p_range != NULL )
	{
	    p_range->Release();
	    p_range = NULL;
	}
	 
    TRACE_OUT;
    return ( hr ); 			
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Run( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [optional][in] */ VARIANT Arg9,
            /* [optional][in] */ VARIANT Arg10,
            /* [optional][in] */ VARIANT Arg11,
            /* [optional][in] */ VARIANT Arg12,
            /* [optional][in] */ VARIANT Arg13,
            /* [optional][in] */ VARIANT Arg14,
            /* [optional][in] */ VARIANT Arg15,
            /* [optional][in] */ VARIANT Arg16,
            /* [optional][in] */ VARIANT Arg17,
            /* [optional][in] */ VARIANT Arg18,
            /* [optional][in] */ VARIANT Arg19,
            /* [optional][in] */ VARIANT Arg20,
            /* [optional][in] */ VARIANT Arg21,
            /* [optional][in] */ VARIANT Arg22,
            /* [optional][in] */ VARIANT Arg23,
            /* [optional][in] */ VARIANT Arg24,
            /* [optional][in] */ VARIANT Arg25,
            /* [optional][in] */ VARIANT Arg26,
            /* [optional][in] */ VARIANT Arg27,
            /* [optional][in] */ VARIANT Arg28,
            /* [optional][in] */ VARIANT Arg29,
            /* [optional][in] */ VARIANT Arg30,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Select( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Show( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ShowDependents( 
            /* [optional][in] */ VARIANT Remove,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ShowDetail( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_ShowDetail( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ShowErrors( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ShowPrecedents( 
            /* [optional][in] */ VARIANT Remove,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ShrinkToFit( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_ShrinkToFit( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Sort( 
            /* [optional][in] */ VARIANT Key1,
            /* [defaultvalue][optional][in] */ XlSortOrder Order1,
            /* [optional][in] */ VARIANT Key2,
            /* [optional][in] */ VARIANT Type,
            /* [defaultvalue][optional][in] */ XlSortOrder Order2,
            /* [optional][in] */ VARIANT Key3,
            /* [defaultvalue][optional][in] */ XlSortOrder Order3,
            /* [defaultvalue][optional][in] */ XlYesNoGuess Header,
            /* [optional][in] */ VARIANT OrderCustom,
            /* [optional][in] */ VARIANT MatchCase,
            /* [defaultvalue][optional][in] */ XlSortOrientation Orientation,
            /* [defaultvalue][optional][in] */ XlSortMethod SortMethod,
            /* [defaultvalue][optional][in] */ XlSortDataOption DataOption1,
            /* [defaultvalue][optional][in] */ XlSortDataOption DataOption2,
            /* [defaultvalue][optional][in] */ XlSortDataOption DataOption3,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::SortSpecial( 
            /* [defaultvalue][optional][in] */ XlSortMethod SortMethod,
            /* [optional][in] */ VARIANT Key1,
            /* [defaultvalue][optional][in] */ XlSortOrder Order1,
            /* [optional][in] */ VARIANT Type,
            /* [optional][in] */ VARIANT Key2,
            /* [defaultvalue][optional][in] */ XlSortOrder Order2,
            /* [optional][in] */ VARIANT Key3,
            /* [defaultvalue][optional][in] */ XlSortOrder Order3,
            /* [defaultvalue][optional][in] */ XlYesNoGuess Header,
            /* [optional][in] */ VARIANT OrderCustom,
            /* [optional][in] */ VARIANT MatchCase,
            /* [defaultvalue][optional][in] */ XlSortOrientation Orientation,
            /* [defaultvalue][optional][in] */ XlSortDataOption DataOption1,
            /* [defaultvalue][optional][in] */ XlSortDataOption DataOption2,
            /* [defaultvalue][optional][in] */ XlSortDataOption DataOption3,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_SoundNote( 
            /* [retval][out] */ SoundNote **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::SpecialCells( 
            /* [in] */ XlCellType Type,
            /* [optional][in] */ VARIANT Value,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Style( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Style( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::SubscribeTo( 
            /* [in] */ BSTR Edition,
            /* [defaultvalue][optional][in] */ XlSubscribeToFormat Format,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Subtotal( 
            /* [in] */ long GroupBy,
            /* [in] */ XlConsolidationFunction Function,
            /* [in] */ VARIANT TotalList,
            /* [optional][in] */ VARIANT Replace,
            /* [optional][in] */ VARIANT PageBreaks,
            /* [defaultvalue][optional][in] */ XlSummaryRow SummaryBelowData,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Summary( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Table( 
            /* [optional][in] */ VARIANT RowInput,
            /* [optional][in] */ VARIANT ColumnInput,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Text( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::TextToColumns( 
            /* [optional][in] */ VARIANT Destination,
            /* [defaultvalue][optional][in] */ XlTextParsingType DataType,
            /* [defaultvalue][optional][in] */ XlTextQualifier TextQualifier,
            /* [optional][in] */ VARIANT ConsecutiveDelimiter,
            /* [optional][in] */ VARIANT Tab,
            /* [optional][in] */ VARIANT Semicolon,
            /* [optional][in] */ VARIANT Comma,
            /* [optional][in] */ VARIANT Space,
            /* [optional][in] */ VARIANT Other,
            /* [optional][in] */ VARIANT OtherChar,
            /* [optional][in] */ VARIANT FieldInfo,
            /* [optional][in] */ VARIANT DecimalSeparator,
            /* [optional][in] */ VARIANT ThousandsSeparator,
            /* [optional][in] */ VARIANT TrailingMinusNumbers,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Top( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::awt::Point  oo_point;
    long count = 0;
    long value = 0;
    
    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
	    ERR( " get_Count \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }    
    
    if ( count == 1 )
        oo_point = (static_cast<SheetCell>(m_oo_range)).getPosition();
    else
	    oo_point = (static_cast<SheetCellRange>(m_oo_range)).getPosition();
     
    if ( oo_point.IsNull() )
    {
	    ERR( " oo_point.IsNull \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }
    
    hr = oo_point.getY( value );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_point.getY \n" );
		TRACE_OUT;
		return ( E_FAIL ); 	     	 
	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = value;
    
    TRACE_OUT;
    return ( hr );		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Ungroup( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_UseStandardHeight( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_UseStandardHeight( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_UseStandardWidth( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_UseStandardWidth( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Validation( 
            /* [retval][out] */ Validation **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Value( 
            /* [optional][in] */ VARIANT RangeValueDataType,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    CorrectArg(RangeValueDataType, &RangeValueDataType);
        
    OORange  oo_range;
		
    // get first cell in range
    oo_range = m_oo_range.getCellByPosition( 0, 0 );
    
    switch ( oo_range.getType( ) )
    {
	    case com::sun::star::table::FORMULA:	   
	    case com::sun::star::table::VALUE:
			 {
			     hr = oo_range.getValue( RHS );
				 if ( FAILED ( hr ) )
				 {
				     ERR( " failed oo_range.getValue \n" ); 	  
				 } 											   
	         }
			 break;
			 
	    case com::sun::star::table::EMPTY:	   
			 {
			     V_VT( RHS ) = VT_EMPTY;
                 hr = S_OK; 												   
	         }
			 break;			 
			 
	    case com::sun::star::table::TEXT:	   
	    default:
			 {		  								   
			     hr = oo_range.getFormula( RHS );
				 if ( FAILED ( hr ) )
				 {
				     ERR( " failed oo_range.getFormula \n" ); 	  
				 }			  												   
	         }
			 break;				 
			  	   
    } // switch ( oo_range.getType( ) )
    
    TRACE_OUT;
    return ( hr );	
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_Value( 
            /* [optional][in] */ VARIANT RangeValueDataType,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    CorrectArg(RangeValueDataType, &RangeValueDataType);
    CorrectArg(RHS, &RHS);
    
    if ( V_VT( &RHS ) & VT_ARRAY )
    {
	    int arr_dim;
        VARIANT *pvar;
		
		arr_dim = SafeArrayGetDim( V_ARRAY( &RHS ) );
		
		if (arr_dim == 1) 
		{
	        /*TODO*/
            ERR( " 1 Demension array NOT REALIZE NOW \n" );
        } // if (arr_dim == 1) 
		
		if (arr_dim == 2) 
		{
  		    long startrow, endrow, startcolumn, endcolumn;
            VARIANT row,col;
            VARIANT vNull;
            VariantInit(&vNull);

            hr = SafeArrayAccessData( V_ARRAY( &RHS ), (void **)&pvar);
            if ( FAILED( hr )) 
            {
			   	ERR( " failed SafeArrayAccessData \n" ); 
			   	TRACE_OUT;
			    return ( hr );
			}

			CellRangeAddress   cell_range_address;

			cell_range_address = m_oo_range.getRangeAddress();
			if ( cell_range_address.IsNull() )
			{
			    ERR( " failed getRangeAddress \n" );  
				TRACE_OUT;
				return ( E_FAIL ); 	 
		    }
			
			startcolumn = cell_range_address.StartColumn( );
			startrow = cell_range_address.StartRow( );
			endcolumn = cell_range_address.EndColumn( );
			endrow = cell_range_address.EndRow( );
			
            int maxj = ( V_ARRAY( &RHS ))->rgsabound[0].cElements;
            int maxi = ( V_ARRAY( &RHS ))->rgsabound[1].cElements;

            for ( int i = 0; i < maxi; i++) 
			{
                for ( int j = 0; j < maxj; j++) 
				{
                    V_VT( &row ) = VT_I4;
                    V_I4( &row ) = i + 1;
                    V_VT( &col ) = VT_I4;
                    V_I4( &col ) = j + 1;

                    if (
					    ( i <= ( endrow - startrow ) ) &&
						( j <= ( endcolumn - startcolumn ) )
						) 
					{
                    
                        hr = put__Default( row, col, 0, pvar[ j * maxi + i ]  );
                        if ( FAILED( hr ) )
						{
						    ERR( " put__Default \n" ); 
							TRACE_OUT;
							return ( hr );  	 
					    }               
                    
                    } // if (( i <= ( endrow - startrow ) ) && ( j <= ( endcolumn - startcolumn ) )) 
                    
                } // for (j=0; j<maxj; j++) 
            } // for (i=0; i<maxi; i++) 
        
            hr = SafeArrayUnaccessData( V_ARRAY( &RHS ) );
        
            if ( FAILED( hr ) ) {
                ERR("Error when SafeArrayUnaccessData \n");
            }
            
            TRACE_OUT;
            return ( hr );
        } // if (arr_dim == 2) 				
	   					
    } //  if ( V_VT( &RHS ) & VT_ARRAY )
    else
    {
    	OORange  oo_range;
		
		// get first cell in range
		oo_range = m_oo_range.getCellByPosition( 0, 0 ); 	
	 	
	 	if ( oo_range.IsNull() )
	 	{
		    hr = E_FAIL;
			
			TRACE_OUT;
			return ( hr );   	 
	    }
	 	
	 	switch V_VT( &RHS ) 
		{
            case VT_I8:/*hack to convert VT_I8 to VT_I4*/
                {
				    long tmp = (long) V_I8(&RHS);
                    VariantClear(&RHS);
                    V_VT(&RHS) = VT_I4;
                    V_I4(&RHS) = tmp;
				}
				break;
        } // switch V_VT( &RHS ) 
	 	
	 	switch V_VT( &RHS ) 
		{
	 	    case VT_BSTR:
	            {
				    if ( lstrlenW( V_BSTR( &RHS ) ) != 0 ) 
					{
					    if ( V_BSTR( &RHS )[0] == L'=' ) 
						{
						    hr = oo_range.setFormula( V_BSTR( &RHS ) );
							
							if ( FAILED( hr ) )
				            {
					             ERR( " oo_range.setFormula \n" );   	 
				            }
							
							TRACE_OUT;
							return ( hr );
							   								
						} // if ( V_BSTR( &RHS )[0] == L'=' )   									
					} // if ( lstrlenW( V_BSTR( &RHS ) ) != 0 ) 
					
                    hr = oo_range.setString( V_BSTR( &RHS ) );
							
					if ( FAILED( hr ) )
                    {
					    ERR( " oo_range.setString \n" );	 
                    } 	
					 	 
				}
				break;
				
	 	    default:
	            {
				    hr = oo_range.setValue( RHS );
					
					if ( FAILED( hr ) )
					{
					    ERR( " oo_range.setValue \n" );   	 
				    }  	 
				}
				break;				
	 	
		} // switch V_VT( &RHS ) 
	
    } //  if ( V_VT( &RHS ) & VT_ARRAY )
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Value2( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Value2( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_VerticalAlignment( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::table::CellVertJustify value = com::sun::star::table::VERT_STANDARD;
    		
	hr = m_oo_range.getVertJustify( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_range.getVertJustify \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
        
    switch ( value ) {
    case com::sun::star::table::VERT_STANDARD: 
		 V_I4( RHS ) = xlVAlignJustify;
		 break;
		 
    case com::sun::star::table::VERT_TOP:
		 V_I4( RHS ) = xlVAlignTop;
		 break;
		 
    case com::sun::star::table::VERT_CENTER:
		 V_I4( RHS ) = xlVAlignCenter;
		 break;
		 
    case com::sun::star::table::VERT_BOTTOM:
		 V_I4( RHS ) = xlVAlignBottom;
		 break;
	 		 
    default:
			V_I4( RHS ) = xlVAlignDistributed;
			break;
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_VerticalAlignment( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::table::CellVertJustify value = com::sun::star::table::VERT_STANDARD;
    
    CorrectArg(RHS, &RHS);
    
    VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
        
    switch (  V_I4( &RHS ) ) {
    case xlVAlignJustify: 
		 value = com::sun::star::table::VERT_STANDARD;
		 break;
		 
    case xlVAlignTop:
		 value = com::sun::star::table::VERT_TOP;
		 break;
		 
    case xlVAlignCenter:
		 value = com::sun::star::table::VERT_CENTER;
		 break;
		 
    case xlVAlignBottom:
		 value = com::sun::star::table::VERT_BOTTOM;
		 break;
	 		 
    default:
			value = com::sun::star::table::VERT_STANDARD;
			break;
    }
    
    hr = m_oo_range.setVertJustify( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_range.setVertJustify \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }	
    
    TRACE_OUT;
    return ( hr );		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Width( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::awt::Size  oo_size;
    long count = 0;
    long value = 0;
    
    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
	    ERR( " get_Count \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }    
    
    if ( count == 1 )
        oo_size = (static_cast<SheetCell>(m_oo_range)).getSize();
    else
	    oo_size = (static_cast<SheetCellRange>(m_oo_range)).getSize();
     
    if ( oo_size.IsNull() )
    {
	    ERR( " oo_size.IsNull \n" );
		TRACE_OUT;
		return ( E_FAIL );   	 
    }
    
    hr = oo_size.getWidth( value );
    if ( FAILED( hr ) )
    {
	    ERR( " oo_size.getWidth \n" );
		TRACE_OUT;
		return ( E_FAIL ); 	     	 
	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = value;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_Worksheet( 
            /* [retval][out] */ Worksheet	**RHS)
{
    TRACE_IN;
    HRESULT hr;
    IDispatch* p_dispatch = NULL;
	 
    hr = get_Parent( &p_dispatch );
    if ( FAILED( hr ) )
    {
	    ERR( " get_Parent \n " ); 
		
  		if ( p_dispatch != NULL )
    	{
	       p_dispatch->Release();
		   p_dispatch = NULL;  
    	}
		 
		TRACE_OUT;
		return ( hr ); 	 
	}
    
    hr = p_dispatch->QueryInterface( CLSID_Worksheet, (void**)(RHS) );
    if ( FAILED( hr ) )
    {
	    ERR( " QueryInterface \n " );  
	} 
    
    if ( p_dispatch != NULL )
    {
	    p_dispatch->Release();
		p_dispatch = NULL;  
    }
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::get_WrapText( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    bool value = false;
    		
	hr = m_oo_range.getisTextWrapped( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_range.getisTextWrapped \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_BOOL;
    if ( value )
	    V_BOOL( RHS ) = VARIANT_TRUE;
	else
	    V_BOOL( RHS ) = VARIANT_FALSE;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::put_WrapText( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    bool value = false;
    
    CorrectArg(RHS, &RHS);
    
    VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    
    if ( V_BOOL(&RHS) == VARIANT_TRUE )
	    value = true;
	else
	    value = false;
		
	hr = m_oo_range.setisTextWrapped( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_range.setisTextWrapped \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::AddComment( 
            /* [optional][in] */ VARIANT Text,
            /* [retval][out] */ Comment **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Comment( 
            /* [retval][out] */ Comment **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
HRESULT STDMETHODCALLTYPE CRange::ClearComments( void)
{
    TRACE_IN;
    HRESULT hr;
    
    hr = m_oo_range.clearContents( CC_ANNOTATION  );
    
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_range.clearContents \n" );   	 
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Phonetic( 
            /* [retval][out] */ Phonetic **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_FormatConditions( 
            /* [retval][out] */ FormatConditions **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ReadingOrder( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_ReadingOrder( 
            /* [in] */ long RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Hyperlinks( 
            /* [retval][out] */ HyperLinks **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Phonetics( 
            /* [retval][out] */ Phonetics **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::SetPhonetic( void)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ID( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_ID( 
            /* [in] */ BSTR RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE CRange::_PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_PivotCell( 
            /* [retval][out] */ PivotCell **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Dirty( void)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Errors( 
            /* [retval][out] */ Errors **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_SmartTags( 
            /* [retval][out] */ SmartTags **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Speak( 
            /* [optional][in] */ VARIANT SpeakDirection,
            /* [optional][in] */ VARIANT SpeakFormulas)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::PasteSpecial( 
            /* [defaultvalue][optional][in] */ XlPasteType Paste,
            /* [defaultvalue][optional][in] */ XlPasteSpecialOperation Operation,
            /* [optional][in] */ VARIANT SkipBlanks,
            /* [optional][in] */ VARIANT Transpose,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_AllowEdit( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ListObject( 
            /* [retval][out] */ ListObject **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_XPath( 
            /* [retval][out] */ XPath **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ServerActions( 
            /* [retval][out] */ Actions **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::RemoveDuplicates( 
            /* [optional][in] */ VARIANT Columns,
            /* [defaultvalue][optional][in] */ XlYesNoGuess Header)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_MDX( 
            /* [retval][out] */ BSTR *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ExportAsFixedFormat( 
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_CountLarge( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::CalculateRowMajorOrder( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
 
 
HRESULT CRange::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_IRange, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr; 		
}
         
HRESULT CRange::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;   		
}
        
HRESULT CRange::Put_Parent( void* p_parent)
{
   TRACE_IN;  
   
   if ( m_p_parent != NULL )
   {
   	  (static_cast<Worksheet*>( m_p_parent ))->Release( );	
       m_p_parent = NULL;	  	
   }
      
   m_p_parent = p_parent;
   
   if ( m_p_parent != NULL )
   {
   	  (static_cast<Worksheet*>( m_p_parent ))->AddRef( );		
   }        
      
   TRACE_OUT;
   return S_OK; 		
}
        
HRESULT CRange::InitWrapper( OORange _oo_range)
{
    m_oo_range = _oo_range;      
}

OOSheet CRange::getParentOOSheet()
{
 	TRACE_IN;
	HRESULT hr;	
 	Worksheet* p_worksheet = NULL;
 	OOSheet oo_sheet;
 	
 	hr = get_Worksheet( &p_worksheet );
 	if ( FAILED( hr ) )
 	{
	    ERR( " get_Worksheet \n" );
	 	if ( p_worksheet != NULL )
 		{
	   	   p_worksheet->Release();
	   	   p_worksheet = NULL;   	 
	   	}
		TRACE_OUT;
		return ( oo_sheet );  	   	 
	}
 	
 	oo_sheet = p_worksheet->getWrapper();
 	
 	if ( p_worksheet != NULL )
 	{
	   p_worksheet->Release();
	   p_worksheet = NULL;   	 
	}
 	
	return ( oo_sheet ); 		 
}
