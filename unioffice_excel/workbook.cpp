#include "workbook.h"

#include "application.h"
#include "workbooks.h"
#include "../OOWrappers/wrap_property_array.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE Workbook::QueryInterface(const IID& iid, void** ppv)
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
    
    if ( iid == IID__Workbook ) {
        TRACE("_Workbook \n");
        *ppv = static_cast<_Workbook*>(this);
    } 
      
    if ( iid == CLSID_Workbook ) {
        TRACE("Workbook \n");
        *ppv = static_cast<Workbook*>(this);
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

ULONG STDMETHODCALLTYPE Workbook::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);        
}

ULONG STDMETHODCALLTYPE Workbook::Release()
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
HRESULT STDMETHODCALLTYPE Workbook::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;        
}

HRESULT STDMETHODCALLTYPE Workbook::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE Workbook::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE Workbook::Invoke(
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
       
       
HRESULT STDMETHODCALLTYPE Workbook::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE Workbook::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Parent( 
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
   
   hr = (static_cast<CWorkbooks*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_AcceptLabelsInFormulas( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_AcceptLabelsInFormulas( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::Activate( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ActiveChart( 
            /* [retval][out] */ Chart **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ActiveSheet( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Author( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_Author( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_AutoUpdateFrequency( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_AutoUpdateFrequency( 
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_AutoUpdateSaveChanges( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_AutoUpdateSaveChanges( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ChangeHistoryDuration( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_ChangeHistoryDuration( 
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_BuiltinDocumentProperties( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::ChangeFileAccess( 
            /* [in] */ XlFileAccess Mode,
            /* [optional][in] */ VARIANT WritePassword,
            /* [optional][in] */ VARIANT Notify,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::ChangeLink( 
            /* [in] */ BSTR Name,
            /* [in] */ BSTR NewName,
            /* [defaultvalue][optional][in] */ XlLinkType Type,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Charts( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::Close( 
            /* [optional][in] */ VARIANT SaveChanges,
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT RouteWorkbook,
            /* [lcid][in] */ long lcid)
{
    TRACE_IN;
    
    HRESULT hr;
    VARIANT_BOOL  _save_changes;
    VARIANT_BOOL  _routeworkbook;
    BSTR          _filename;   
     
    VARIANT_BOOL  hard_close = VARIANT_TRUE;
            
    CorrectArg(SaveChanges, &SaveChanges);
    CorrectArg(Filename, &Filename);
    CorrectArg(RouteWorkbook, &RouteWorkbook);
    
    if ( V_VT(&Filename) != VT_BSTR ) 
    {
        _filename = SysAllocString(L"");
    } else 
    {
        _filename = SysAllocString( V_BSTR(&Filename) );          
    }
    
    if ( (Is_Variant_Null( SaveChanges ))||
         (V_VT(&SaveChanges) != VT_BOOL) )
    {
        _save_changes = VARIANT_FALSE;         
    } else
    {
        _save_changes = V_BOOL( &Filename );     
    }
    
    if ( (Is_Variant_Null( RouteWorkbook ))||
         (V_VT(&RouteWorkbook) != VT_BOOL) )
    {
        _routeworkbook = VARIANT_FALSE;         
    } else
    {
        _routeworkbook = V_BOOL( &RouteWorkbook );     
    }    
    
    //////////////////////////////////////
    // if no filename - only close.
    //////////////////////////////////////
    if ( !lstrcmpiW( _filename, L"" ) ) 
    {
        hr = m_oo_document.Close( hard_close );

        if ( FAILED( hr ) )
        { 
            ERR(" FAILED 1 CLOSE \n"); 
        }
        
    } else
    {
        /////////////////////////////////
        // firs save document, then close
        /////////////////////////////////   
          
        VARIANT FileFormat;
        VARIANT Password;
        VARIANT WriteResPassword;
        VARIANT ReadOnlyRecommended;
        VARIANT CreateBackup;
        XlSaveAsAccessMode AccessMode = xlNoChange;
        VARIANT ConflictResolution;
        VARIANT AddToMru;
        VARIANT TextCodepage;
        VARIANT TextVisualLayout;
        VARIANT Local;  
         
        VariantInit( &FileFormat          );
        VariantInit( &Password            );
        VariantInit( &WriteResPassword    );
        VariantInit( &ReadOnlyRecommended );
        VariantInit( &CreateBackup        );
        VariantInit( &ConflictResolution  );
        VariantInit( &AddToMru            );
        VariantInit( &TextCodepage        );
        VariantInit( &TextVisualLayout    );
        VariantInit( &Local               );                 
            
        hr = SaveAs( Filename,
            FileFormat,
            Password,
            WriteResPassword,
            ReadOnlyRecommended,
            CreateBackup,
            AccessMode,
            ConflictResolution,
            AddToMru,
            TextCodepage,
            TextVisualLayout,
            Local,
            lcid);
            
        if ( FAILED( hr ) )
        { 
            ERR(" FAILED SaveAs \n"); 
        }    
        
        VariantClear( &FileFormat          );
        VariantClear( &Password            );
        VariantClear( &WriteResPassword    );
        VariantClear( &ReadOnlyRecommended );
        VariantClear( &CreateBackup        );
        VariantClear( &ConflictResolution  );
        VariantClear( &AddToMru            );
        VariantClear( &TextCodepage        );
        VariantClear( &TextVisualLayout    );
        VariantClear( &Local               );  
            
        // close
            
        hr = m_oo_document.Close( hard_close );

        if ( FAILED( hr ) )
        { 
            ERR(" FAILED CLOSE \n"); 
        }          
          
    }
   
    SysFreeString( _filename );
    
    if ( FAILED( hr ) )
    {
        ERR( " call Close \n" ); 
        return ( hr );     
    }
    
    // delete from common vector
    
    reinterpret_cast<CWorkbooks*>(m_p_parent)->DeleteWorkbookFromVector( this );
          
    TRACE_OUT;
    return ( hr );        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_CodeName( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get__CodeName( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put__CodeName( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Colors( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_Colors( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_CommandBars( 
            /* [retval][out] */ CommandBars **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Comments( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_Comments( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ConflictResolution( 
            /* [retval][out] */ XlSaveConflictResolution *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_ConflictResolution( 
            /* [in] */ XlSaveConflictResolution RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Container( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_CreateBackup( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_CustomDocumentProperties( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Date1904( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_Date1904( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::DeleteNumberFormat( 
            /* [in] */ BSTR NumberFormat,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_DialogSheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_DisplayDrawingObjects( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlDisplayDrawingObjects *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_DisplayDrawingObjects( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlDisplayDrawingObjects RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::ExclusiveAccess( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_FileFormat( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlFileFormat *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::ForwardMailer( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_FullName( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_HasMailer( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_HasMailer( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_HasPassword( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_HasRoutingSlip( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_HasRoutingSlip( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_IsAddin( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_IsAddin( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Keywords( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_Keywords( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::LinkInfo( 
            /* [in] */ BSTR Name,
            /* [in] */ XlLinkInfo LinkInfo,
            /* [optional][in] */ VARIANT Type,
            /* [optional][in] */ VARIANT EditionRef,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::LinkSources( 
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Mailer( 
            /* [retval][out] */ Mailer **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::MergeWorkbook( 
            /* [in] */ VARIANT Filename)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Modules( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_MultiUserEditing( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Name( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Names( 
            /* [retval][out] */ Names **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::NewWindow( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Window **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_OnSave( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_OnSave( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::OpenLinks( 
            /* [in] */ BSTR Name,
            /* [optional][in] */ VARIANT ReadOnly,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Path( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_PersonalViewListSettings( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_PersonalViewListSettings( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_PersonalViewPrintSettings( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_PersonalViewPrintSettings( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::Pivotcaches( 
            /* [retval][out] */ PivotCaches	**RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::Post( 
            /* [optional][in] */ VARIANT DestName,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_PrecisionAsDisplayed( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_PrecisionAsDisplayed( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Workbook::__PrintOut( 
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
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Workbook::_Protect( 
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT Structure,
            /* [optional][in] */ VARIANT Windows)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Workbook::_ProtectSharing( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT SharingPassword)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ProtectStructure( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ProtectWindows( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ReadOnly( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get__ReadOnlyRecommended( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::RefreshAll( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::Reply( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::ReplyAll( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::RemoveUser( 
            /* [in] */ long Index)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_RevisionNumber( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Workbook::Route( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Routed( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_RoutingSlip( 
            /* [retval][out] */ RoutingSlip **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::RunAutoMacros( 
            /* [in] */ XlRunAutoMacro Which,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::Save( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;
}
        
HRESULT STDMETHODCALLTYPE Workbook::_SaveAs( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [defaultvalue][optional][in] */ XlSaveAsAccessMode AccessMode,
            /* [optional][in] */ VARIANT ConflictResolution,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [lcid][in] */ long lcid)
{    
/*    TRACE_IN;
    
    HRESULT hr;
    BSTR    FilenameURL;
    Application*        application;
    WrapPropertyArray   wrap_property_array;
    
    CorrectArg(Filename, &Filename);
        
    if (V_VT(&Filename)!=VT_BSTR) {
        ERR(" Filename != BSTR \n");
        return E_FAIL;
    }

        
    hr = get_Application( &application );
    
    if ( FAILED( hr ) )
    {
        ERR( " Get_Application \n" );     
    }
    
    ////////////////////////////
    // Fill properties
    OOPropertyValue property_1 = application->m_oo_service_manager.Get_PropertyValue();
    property_1.Set_Property( SysAllocString(L"FilterName"), SysAllocString(L"MS Excel 97") );
    
    wrap_property_array.Clear();
    wrap_property_array.Add( property_1 );
    // Fill properties
    ////////////////////////////
    
    MakeURLFromFilename(V_BSTR(&Filename), &FilenameURL);
    
    hr = m_oo_document.StoreAsURL( FilenameURL,                 
                                  wrap_property_array );
        
                                                                                                     
    if ( FAILED( hr ) )
    {
        ERR( " StoreAsURL \n " );
        hr = E_FAIL;     
    }                                                    
    
    SysFreeString( FilenameURL );
        
    TRACE_OUT;    
    return ( hr );  
    */
    
    TRACE_NOTIMPL;
    return ( E_NOTIMPL ); 
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::SaveCopyAs( 
            /* [optional][in] */ VARIANT Filename,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Saved( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_Saved( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_SaveLinkValues( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_SaveLinkValues( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::SendMail( 
            /* [in] */ VARIANT Recipients,
            /* [optional][in] */ VARIANT Subject,
            /* [optional][in] */ VARIANT ReturnReceipt,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::SendMailer( 
            /* [optional][in] */ VARIANT FileFormat,
            /* [defaultvalue][optional][in] */ XlPriority Priority,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::SetLinkOnData( 
            /* [in] */ BSTR Name,
            /* [optional][in] */ VARIANT Procedure,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Sheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_ShowConflictHistory( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_ShowConflictHistory( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Styles( 
            /* [retval][out] */ Styles **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Subject( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Workbook::put_Subject( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Title( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_Title( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::Unprotect( 
            /* [optional][in] */ VARIANT Password,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::UnprotectSharing( 
            /* [optional][in] */ VARIANT SharingPassword)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::UpdateFromFile( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::UpdateLink( 
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_UpdateRemoteReferences( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_UpdateRemoteReferences( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_UserControl( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_UserControl( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_UserStatus( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_CustomViews( 
            /* [retval][out] */ CustomViews **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Windows( 
            /* [retval][out] */ Windows **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Worksheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_WriteReserved( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_WriteReservedBy( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Excel4IntlMacroSheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Excel4MacroSheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_TemplateRemoveExtData( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_TemplateRemoveExtData( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::HighlightChangesOptions( 
            /* [optional][in] */ VARIANT When,
            /* [optional][in] */ VARIANT Who,
            /* [optional][in] */ VARIANT Where)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_HighlightChangesOnScreen( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_HighlightChangesOnScreen( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_KeepChangeHistory( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_KeepChangeHistory( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ListChangesOnNewSheet( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_ListChangesOnNewSheet( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::PurgeChangeHistoryNow( 
            /* [in] */ long Days,
            /* [optional][in] */ VARIANT SharingPassword)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::AcceptAllChanges( 
            /* [optional][in] */ VARIANT When,
            /* [optional][in] */ VARIANT Who,
            /* [optional][in] */ VARIANT Where)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::RejectAllChanges( 
            /* [optional][in] */ VARIANT When,
            /* [optional][in] */ VARIANT Who,
            /* [optional][in] */ VARIANT Where)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::PivotTableWizard( 
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
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::ResetColors( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_VBProject( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::FollowHyperlink( 
            /* [in] */ BSTR Address,
            /* [optional][in] */ VARIANT SubAddress,
            /* [optional][in] */ VARIANT NewWindow,
            /* [optional][in] */ VARIANT AddHistory,
            /* [optional][in] */ VARIANT ExtraInfo,
            /* [optional][in] */ VARIANT Method,
            /* [optional][in] */ VARIANT HeaderInfo)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::AddToFavorites( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_IsInplace( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::_PrintOut( 
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
        
HRESULT STDMETHODCALLTYPE Workbook::WebPagePreview( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_PublishObjects( 
            /* [retval][out] */ PublishObjects **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_WebOptions( 
            /* [retval][out] */ WebOptions **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::ReloadAs( 
            /* [in] */ MsoEncoding Encoding)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_HTMLProject( 
            /* [retval][out] */ HTMLProject **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_EnvelopeVisible( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_EnvelopeVisible( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_CalculationVersion( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::Dummy17( 
            /* [in] */ long calcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::sblt( 
            /* [in] */ BSTR s)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_VBASigned( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ShowPivotTableFieldList( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_ShowPivotTableFieldList( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_UpdateLinks( 
            /* [retval][out] */ XlUpdateLinks *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_UpdateLinks( 
            /* [in] */ XlUpdateLinks RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::BreakLink( 
            /* [in] */ BSTR Name,
            /* [in] */ XlLinkType Type)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::Dummy16( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::SaveAs( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT FileFormat,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [defaultvalue][optional][in] */ XlSaveAsAccessMode AccessMode,
            /* [optional][in] */ VARIANT ConflictResolution,
            /* [optional][in] */ VARIANT AddToMru,
            /* [optional][in] */ VARIANT TextCodepage,
            /* [optional][in] */ VARIANT TextVisualLayout,
            /* [optional][in] */ VARIANT Local,
            /* [lcid][in] */ long lcid)
{
    TRACE_IN;
    
    HRESULT hr;
    BSTR    FilenameURL;
    Application*        application;
    WrapPropertyArray   wrap_property_array;
    
    CorrectArg(Filename, &Filename);
        
    if ( V_VT(&Filename) != VT_BSTR ) {
        ERR(" Filename != BSTR \n");
        return E_FAIL;
    }
        
    hr = get_Application( &application );
    
    if ( FAILED( hr ) )
    {
        ERR( " Get_Application \n" );     
    }
    
    ////////////////////////////
    // Fill properties
    OOPropertyValue property_1 = application->m_oo_service_manager.Get_PropertyValue();
    hr = property_1.Set_Property( SysAllocString(L"FilterName"), 
                                  SysAllocString(L"MS Excel 97") );
    
    if ( FAILED( hr ) )
    {
        ERR( " Set_Property \n" );     
    }
    
    wrap_property_array.Clear();
    wrap_property_array.Add( property_1 );
    // Fill properties
    ////////////////////////////
    
    MakeURLFromFilename(V_BSTR(&Filename), &FilenameURL);
    
    hr = m_oo_document.StoreAsURL( FilenameURL,                 
                                   wrap_property_array );
        
                                                                                                     
    if ( FAILED( hr ) )
    {
        ERR( " StoreAsURL \n " );
        hr = E_FAIL;     
    }                                                    
    
    reinterpret_cast<IDispatch*>(application)->Release();
    
    SysFreeString( FilenameURL );
        
    TRACE_OUT;    
    return ( hr );        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_EnableAutoRecover( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_EnableAutoRecover( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_RemovePersonalInformation( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_RemovePersonalInformation( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_FullNameURLEncoded( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::CheckIn( 
            /* [optional][in] */ VARIANT SaveChanges,
            /* [optional][in] */ VARIANT Comments,
            /* [optional][in] */ VARIANT MakePublic)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::CanCheckIn( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::SendForReview( 
            /* [optional][in] */ VARIANT Recipients,
            /* [optional][in] */ VARIANT Subject,
            /* [optional][in] */ VARIANT ShowMessage,
            /* [optional][in] */ VARIANT IncludeAttachment)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::ReplyWithChanges( 
            /* [optional][in] */ VARIANT ShowMessage)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::EndReview( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Password( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_Password( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_WritePassword( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_WritePassword( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_PasswordEncryptionProvider( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_PasswordEncryptionAlgorithm( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_PasswordEncryptionKeyLength( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::SetPasswordEncryptionOptions( 
            /* [optional][in] */ VARIANT PasswordEncryptionProvider,
            /* [optional][in] */ VARIANT PasswordEncryptionAlgorithm,
            /* [optional][in] */ VARIANT PasswordEncryptionKeyLength,
            /* [optional][in] */ VARIANT PasswordEncryptionFileProperties)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_PasswordEncryptionFileProperties( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ReadOnlyRecommended( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_ReadOnlyRecommended( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::Protect( 
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT Structure,
            /* [optional][in] */ VARIANT Windows)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_SmartTagOptions( 
            /* [retval][out] */ SmartTagOptions **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::RecheckSmartTags( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Workbook::get_Permission( 
            /* [retval][out] */ Permission **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_SharedWorkspace( 
            /* [retval][out] */ SharedWorkspace **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Sync( 
            /* [retval][out] */ Sync **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::SendFaxOverInternet( 
            /* [optional][in] */ VARIANT Recipients,
            /* [optional][in] */ VARIANT Subject,
            /* [optional][in] */ VARIANT ShowMessage)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_XmlNamespaces( 
            /* [retval][out] */ XmlNamespaces **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_XmlMaps( 
            /* [retval][out] */ XmlMaps **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::XmlImport( 
            /* [in] */ BSTR Url,
            /* [out] */ XmlMap **ImportMap,
            /* [optional][in] */ VARIANT Overwrite,
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ XlXmlImportResult *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_SmartDocument( 
            /* [retval][out] */ SmartDocument **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_DocumentLibraryVersions( 
            /* [retval][out] */ DocumentLibraryVersions **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_InactiveListBorderVisible( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_InactiveListBorderVisible( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_DisplayInkComments( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_DisplayInkComments( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
  /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Workbook::XmlImportXml( 
            /* [in] */ BSTR Data,
            /* [out] */ XmlMap **ImportMap,
            /* [optional][in] */ VARIANT Overwrite,
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ XlXmlImportResult *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::SaveAsXMLData( 
            /* [in] */ BSTR Filename,
            /* [in] */ XmlMap *Map)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::ToggleFormsDesign( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ContentTypeProperties( 
            /* [retval][out] */ MetaProperties **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Connections( 
            /* [retval][out] */ Connections **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::RemoveDocumentInformation( 
            /* [in] */ XlRemoveDocInfoType RemoveDocInfoType)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Signatures( 
            /* [retval][out] */ SignatureSet **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::CheckInWithVersion( 
            /* [optional][in] */ VARIANT SaveChanges,
            /* [optional][in] */ VARIANT Comments,
            /* [optional][in] */ VARIANT MakePublic,
            /* [optional][in] */ VARIANT VersionType)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ServerPolicy( 
            /* [retval][out] */ ServerPolicy **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::LockServerFile( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_DocumentInspectors( 
            /* [retval][out] */ DocumentInspectors **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::GetWorkflowTasks( 
            /* [retval][out] */ WorkflowTasks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::GetWorkflowTemplates( 
            /* [retval][out] */ WorkflowTemplates **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::PrintOut( 
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
        
HRESULT STDMETHODCALLTYPE Workbook::get_ServerViewableItems( 
            /* [retval][out] */ ServerViewableItems **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_TableStyles( 
            /* [retval][out] */ TableStyles **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_DefaultTableStyle( 
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_DefaultTableStyle( 
            /* [in] */ VARIANT RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_DefaultPivotTableStyle( 
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_DefaultPivotTableStyle( 
            /* [in] */ VARIANT RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_CheckCompatibility( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_CheckCompatibility( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_HasVBProject( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_CustomXMLParts( 
            /* [retval][out] */ CustomXMLParts **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Final( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_Final( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Research( 
            /* [retval][out] */ Research **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Theme( 
            /* [retval][out] */ OfficeTheme **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::ApplyTheme( 
            /* [in] */ BSTR Filename)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_Excel8CompatibilityMode( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ConnectionsDisabled( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::EnableConnections( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ShowPivotChartActiveFields( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_ShowPivotChartActiveFields( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::ExportAsFixedFormat( 
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
        
HRESULT STDMETHODCALLTYPE Workbook::get_IconSets( 
            /* [retval][out] */ IconSets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_EncryptionProvider( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_EncryptionProvider( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_DoNotPromptForConvert( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_DoNotPromptForConvert( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::get_ForceFullCalculation( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::put_ForceFullCalculation( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}
        
HRESULT STDMETHODCALLTYPE Workbook::ProtectSharing( 
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT WriteResPassword,
            /* [optional][in] */ VARIANT ReadOnlyRecommended,
            /* [optional][in] */ VARIANT CreateBackup,
            /* [optional][in] */ VARIANT SharingPassword,
            /* [optional][in] */ VARIANT FileFormat)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;        
}


HRESULT Workbook::Init()
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID__Workbook, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;        
}

HRESULT Workbook::Put_Visible( VARIANT_BOOL RHS )
{     
   TRACE_STUB;
   return S_OK;         
}

HRESULT Workbook::Put_Application( void* p_application )
{
    m_p_application = p_application;
        
    return S_OK;      
}

HRESULT Workbook::Put_Parent( void* p_parent )
{
   m_p_parent = p_parent;
   
   return S_OK;     
}

HRESULT Workbook::NewDocument( )
{
    TRACE_IN;
    
    HRESULT hr;
    
    Application*        application;
    WrapPropertyArray   wrap_property_array;
    
    hr = get_Application( &application );
    
    if ( FAILED( hr ) )
    {
        ERR( " Get_Application \n" );     
    }
    
    ////////////////////////////
    // Fill properties
    OOPropertyValue property_1 = application->m_oo_service_manager.Get_PropertyValue();
    property_1.Set_PropertyName( SysAllocString(L"Hidden") );
    VARIANT_BOOL _visible;
    hr = application->get_Visible( /*lcid*/ 0, &_visible );
    if ( FAILED( hr ) )
    {
        ERR( "  \n" );     
    }
    property_1.Set_PropertyValue( _visible );
    
    wrap_property_array.Clear();
    wrap_property_array.Add( property_1 );
    // Fill properties
    ////////////////////////////
    
    m_oo_document = application->m_oo_desktop.LoadComponentFromURL( 
                                                    SysAllocString(L"private:factory/scalc"), 
                                                    SysAllocString(L"_blank"), 
                                                    0,
                                                    wrap_property_array );
      
    hr = S_OK;  
                                                                                                     
    if ( m_oo_document.IsNull() )
    {
        ERR( " m_oo_document not load \n" );
        hr = E_FAIL;     
    }                                                    
    
    reinterpret_cast<IDispatch*>(application)->Release();
        
    TRACE_OUT;    
    return ( hr );        
}

HRESULT Workbook::NewDocumentAsTemplate( BSTR template_name )
{
    TRACE_IN;
    
    HRESULT hr;
    
    Application* application;
    
    hr = get_Application( &application );
    
    if ( FAILED( hr ) )
    {
        ERR( " Get_Application \n" );     
    }
       
    
    TRACE_NOTIMPL;
    hr = E_NOTIMPL;    
        
    TRACE_OUT;    
    return ( hr );       
} 
