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


       // IUnknown
HRESULT STDMETHODCALLTYPE CRange::CRange::QueryInterface(const IID& iid, void** ppv)
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
    
    return hr;  		
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
     ERR( " dispIdMember = %i \n", dispIdMember );     
    }  
                 
    return hr;  		
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
   
   hr = (static_cast<Application*>( m_p_application ))->get_Application( RHS );          
             
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
   
   hr = (static_cast<Worksheet*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Borders( 
            /* [retval][out] */ Borders	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Clear( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ClearContents( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ClearFormats( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ClearNotes( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ClearOutline( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Column( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ColumnDifferences( 
            /* [in] */ VARIANT Comparison,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Columns( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_ColumnWidth( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_ColumnWidth( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Count( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get__Default( 
            /* [optional][in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put__Default( 
            /* [optional][in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_EntireRow( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Font( 
            /* [retval][out] */ /* external definition not present */ Font **RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Formula( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Formula( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Height( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_HorizontalAlignment( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_HorizontalAlignment( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Interior( 
            /* [retval][out] */ Interior	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Left( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::Merge( 
            /* [optional][in] */ VARIANT Across)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::UnMerge( void)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_MergeArea( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_MergeCells( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_MergeCells( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Row( 
            /* [retval][out] */ long *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::RowDifferences( 
            /* [in] */ VARIANT Comparison,
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_RowHeight( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_RowHeight( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Rows( 
            /* [retval][out] */ Range	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Top( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Value( 
            /* [optional][in] */ VARIANT RangeValueDataType,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_Value( 
            /* [optional][in] */ VARIANT RangeValueDataType,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_VerticalAlignment( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_VerticalAlignment( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Width( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_Worksheet( 
            /* [retval][out] */ Worksheet	**RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CRange::get_WrapText( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
}
        
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CRange::put_WrapText( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
        
        
        /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CRange::ClearComments( void)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL; 		
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
      
   m_p_parent = p_parent;
   
   TRACE_OUT;
   return S_OK; 		
}
        


