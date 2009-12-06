/*
 * implementation of Application
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

#include "application.h"
#include "worksheet.h"


       // IUnknown
HRESULT STDMETHODCALLTYPE Application::QueryInterface(const IID& iid, void** ppv)
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
    
    if ( iid == IID__Application ) {
        TRACE("_Application \n");
        *ppv = static_cast<_Application*>(this);
    } 
      
    if ( iid == CLSID_Application ) {
        TRACE("Application \n");
        *ppv = static_cast<Application*>(this);
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

ULONG STDMETHODCALLTYPE Application::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);                        
}

ULONG STDMETHODCALLTYPE Application::Release()
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
HRESULT STDMETHODCALLTYPE Application::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;                      
}

HRESULT STDMETHODCALLTYPE Application::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE Application::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE Application::Invoke(
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
                  
        // _Application
HRESULT STDMETHODCALLTYPE Application::get_Application( 
            Application	**RHS)
{
   HRESULT hr = S_OK;         
   TRACE_IN;
            
   hr = QueryInterface( CLSID_Application, ( void** ) RHS );         
   
   TRACE_OUT;         
   return hr;                         
}
        
HRESULT STDMETHODCALLTYPE Application::get_Creator( 
             XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;                          
}
        
HRESULT STDMETHODCALLTYPE Application::get_Parent( 
             Application	**RHS)
{
   TRACE_IN;          
          
   HRESULT hr = get_Application( RHS );
   if ( FAILED( hr ) )
   {
       ERR( " \n " );     
   }          
             
   TRACE_OUT;
   return hr;                          
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActiveCell( 
             Range **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                          
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActiveChart( 
             Chart **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                         
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActiveDialog( 
             DialogSheet **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActiveMenuBar( 
             MenuBar **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                           
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActivePrinter( 
             long lcid,
             BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                         
}
        
HRESULT STDMETHODCALLTYPE Application::put_ActivePrinter( 
             long lcid,
             BSTR RHS) 
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;                          
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActiveSheet( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   Workbook* p_workbook = NULL;
   
   hr = get_ActiveWorkbook( &p_workbook );
   if ( FAILED( hr ) || ( p_workbook == NULL))
   {
       ERR( " get_ActiveWorkbook \n" );     
       p_workbook = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }
   
   hr = p_workbook->get_ActiveSheet( RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " p_workbook->get_ActiveSheet \n" );     
       p_workbook->Release();
       p_workbook = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }

   p_workbook->Release();
   p_workbook = NULL;
   
   TRACE_OUT;
   return ( hr );           
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActiveWindow( 
            /* [retval][out] */ Window **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL; 
}
        
HRESULT STDMETHODCALLTYPE Application::get_ActiveWorkbook( 
            /* [retval][out] */ Workbook **RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   hr = m_workbooks.GetActiveWorkbook( RHS );
   if ( FAILED( hr ) )
   {
       ERR( " m_workbooks.GetActiveWorkbook \n" );
   }
   
   TRACE_OUT;
   return hr;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_AddIns( 
            /* [retval][out] */ AddIns **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_Assistant( 
            /* [retval][out] */ Assistant **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::Calculate( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_Cells( 
            /* [retval][out] */ Range **RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   Worksheet* p_worksheet = NULL;
   
   hr = get_ActiveSheet( reinterpret_cast<IDispatch**>(&p_worksheet) );
   if ( FAILED( hr ) || ( p_worksheet == NULL))
   {
       ERR( " get_ActiveSheet \n" );     
       p_worksheet = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }
   
   hr = p_worksheet->get_Cells( RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " p_worksheet->get_Cells \n" );     
       p_worksheet->Release();
       p_worksheet = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }

   p_worksheet->Release();
   p_worksheet = NULL;
   
   TRACE_OUT;
   return ( hr );     			
}
        
HRESULT STDMETHODCALLTYPE Application::get_Charts( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_Columns( 
            /* [retval][out] */ Range **RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   Worksheet* p_worksheet = NULL;
   
   hr = get_ActiveSheet( reinterpret_cast<IDispatch**>(&p_worksheet) );
   if ( FAILED( hr ) || ( p_worksheet == NULL))
   {
       ERR( " get_ActiveSheet \n" );     
       p_worksheet = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }
   
   hr = p_worksheet->get_Columns( RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " p_worksheet->get_Columns \n" );     
       p_worksheet->Release();
       p_worksheet = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }

   p_worksheet->Release();
   p_worksheet = NULL;
   
   TRACE_OUT;
   return ( hr );           
}
        
HRESULT STDMETHODCALLTYPE Application::get_CommandBars( 
            /* [retval][out] */ CommandBars **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_DDEAppReturnCode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::DDEExecute( 
            /* [in] */ long Channel,
            /* [in] */ BSTR String,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::DDEInitiate( 
            /* [in] */ BSTR App,
            /* [in] */ BSTR Topic,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::DDEPoke( 
            /* [in] */ long Channel,
            /* [in] */ VARIANT Item,
            /* [in] */ VARIANT Data,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::DDERequest( 
            /* [in] */ long Channel,
            /* [in] */ BSTR Item,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE Application::DDETerminate( 
            /* [in] */ long Channel,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_DialogSheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::_Evaluate( 
            /* [in] */ VARIANT Name,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::ExecuteExcel4Macro( 
            /* [in] */ BSTR String,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Intersect( 
            /* [in] */ Range *Arg1,
            /* [in] */ Range *Arg2,
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
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MenuBars( 
            /* [retval][out] */ MenuBars **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Modules( 
            /* [retval][out] */ Modules **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Names( 
            /* [retval][out] */ Names **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_Range( 
            /* [in] */ VARIANT Cell1,
            /* [optional][in] */ VARIANT Cell2,
            /* [retval][out] */ Range **RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   CorrectArg(Cell1, &Cell1);
   CorrectArg(Cell2, &Cell2);
   
   Worksheet* p_worksheet = NULL;
   
   hr = get_ActiveSheet( reinterpret_cast<IDispatch**>(&p_worksheet) );
   if ( FAILED( hr ) || ( p_worksheet == NULL))
   {
       ERR( " get_ActiveSheet \n" );     
       p_worksheet = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }
   
   hr = p_worksheet->get_Range( Cell1, Cell2, RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " p_worksheet->get_Range \n" );     
       p_worksheet->Release();
       p_worksheet = NULL;
       
       TRACE_OUT;
       return ( hr ); 
   }

   p_worksheet->Release();
   p_worksheet = NULL;
   
   TRACE_OUT;
   return ( hr );             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Rows( 
            /* [retval][out] */ Range **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Run( 
            /* [optional][in] */ VARIANT Macro,
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
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::_Run2( 
            /* [optional][in] */ VARIANT Macro,
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
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Selection( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::SendKeys( 
            /* [in] */ VARIANT Keys,
            /* [optional][in] */ VARIANT Wait,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_Sheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_IN;
   HRESULT hr;
   
   Workbook* p_workbook = NULL;
   
   hr = get_ActiveWorkbook( &p_workbook );
   if ( FAILED( hr ) || ( p_workbook == NULL))
   {
       ERR( " get_ActiveWorkbook \n" );     
       p_workbook = NULL;
       return ( hr ); 
   }
   
   hr = p_workbook->get_Sheets( RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " p_workbook->get_Sheets \n" );     
       p_workbook->Release();
       p_workbook = NULL;
       return ( hr ); 
   }
   
   p_workbook->Release();
   p_workbook = NULL;
   
   TRACE_OUT;
   return ( hr );             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShortcutMenus( 
            /* [in] */ long Index,
            /* [retval][out] */ Menu **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ThisWorkbook( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Workbook **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Toolbars( 
            /* [retval][out] */ Toolbars **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Union( 
            /* [in] */ Range *Arg1,
            /* [in] */ Range *Arg2,
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
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Range **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Windows( 
            /* [retval][out] */ Windows **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_Workbooks( 
             Workbooks **RHS)
{
    HRESULT hr = m_workbooks.QueryInterface( IID_Workbooks, (void**) RHS );         

    if ( FAILED( hr ) )
    {
        ERR( " QueryInterface \n" );     
    }

    return ( hr );             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_WorksheetFunction( 
            /* [retval][out] */ WorksheetFunction **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Worksheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Excel4IntlMacroSheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Excel4MacroSheets( 
            /* [retval][out] */ Sheets **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::ActivateMicrosoftApp( 
            /* [in] */ XlMSApplication Index,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::AddChartAutoFormat( 
            /* [in] */ VARIANT Chart,
            /* [in] */ BSTR Name,
            /* [optional][in] */ VARIANT Description,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::AddCustomList( 
            /* [in] */ VARIANT ListArray,
            /* [optional][in] */ VARIANT ByRow,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AlertBeforeOverwriting( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_AlertBeforeOverwriting( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AltStartupPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_AltStartupPath( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AskToUpdateLinks( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_AskToUpdateLinks( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableAnimations( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableAnimations( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AutoCorrect( 
            /* [retval][out] */ AutoCorrect **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Build( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CalculateBeforeSave( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_CalculateBeforeSave( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Calculation( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlCalculation *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Calculation( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlCalculation RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Caller( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CanPlaySounds( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CanRecordSounds( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Caption( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Caption( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CellDragAndDrop( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_CellDragAndDrop( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::CentimetersToPoints( 
            /* [in] */ double Centimeters,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::CheckSpelling( 
            /* [in] */ BSTR Word,
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ClipboardFormats( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayClipboardWindow( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayClipboardWindow( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ColorButtons( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ColorButtons( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CommandUnderlines( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlCommandUnderlines *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_CommandUnderlines( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlCommandUnderlines RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ConstrainNumeric( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ConstrainNumeric( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::ConvertFormula( 
            /* [in] */ VARIANT Formula,
            /* [in] */ XlReferenceStyle FromReferenceStyle,
            /* [optional][in] */ VARIANT ToReferenceStyle,
            /* [optional][in] */ VARIANT ToAbsolute,
            /* [optional][in] */ VARIANT RelativeTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CopyObjectsWithCells( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_CopyObjectsWithCells( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Cursor( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlMousePointer *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Cursor( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlMousePointer RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CustomListCount( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CutCopyMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlCutCopyMode *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_CutCopyMode( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlCutCopyMode RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DataEntryMode( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DataEntryMode( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy1( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy2( 
            /* [optional][in] */ VARIANT Arg1,
            /* [optional][in] */ VARIANT Arg2,
            /* [optional][in] */ VARIANT Arg3,
            /* [optional][in] */ VARIANT Arg4,
            /* [optional][in] */ VARIANT Arg5,
            /* [optional][in] */ VARIANT Arg6,
            /* [optional][in] */ VARIANT Arg7,
            /* [optional][in] */ VARIANT Arg8,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy3( 
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy4( 
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
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy5( 
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
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy6( 
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy7( 
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy8( 
            /* [optional][in] */ VARIANT Arg1,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy9( 
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy10( 
            /* [optional][in] */ VARIANT arg,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy11( void){}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get__Default( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DefaultFilePath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DefaultFilePath( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::DeleteChartAutoFormat( 
            /* [in] */ BSTR Name,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::DeleteCustomList( 
            /* [in] */ long ListNum,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Dialogs( 
            /* [retval][out] */ Dialogs **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_DisplayAlerts( 
             long lcid,
             VARIANT_BOOL *RHS)
{      
    TRACE_IN;
             
    *RHS = m_b_displayalerts;  
                    
    TRACE_OUT;
    return S_OK;             
}
        
HRESULT STDMETHODCALLTYPE Application::put_DisplayAlerts( 
             long lcid,
             VARIANT_BOOL RHS)
{
    TRACE_IN;
    
    m_b_displayalerts = RHS;     
      
    TRACE_OUT;         
    return S_OK;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayFormulaBar( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayFormulaBar( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayFullScreen( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayFullScreen( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayNoteIndicator( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayNoteIndicator( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayCommentIndicator( 
            /* [retval][out] */ XlCommentDisplayMode *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayCommentIndicator( 
            /* [in] */ XlCommentDisplayMode RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayExcel4Menus( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayExcel4Menus( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayRecentFiles( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayRecentFiles( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayScrollBars( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayScrollBars( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayStatusBar( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayStatusBar( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::DoubleClick( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EditDirectlyInCell( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EditDirectlyInCell( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableAutoComplete( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableAutoComplete( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableCancelKey( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlEnableCancelKey *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableCancelKey( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlEnableCancelKey RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableSound( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableSound( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableTipWizard( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableTipWizard( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FileConverters( 
            /* [optional][in] */ VARIANT Index1,
            /* [optional][in] */ VARIANT Index2,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FileSearch( 
            /* [retval][out] */ FileSearch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FileFind( 
            /* [retval][out] */ IFind **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::_FindFile( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FixedDecimal( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_FixedDecimal( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FixedDecimalPlaces( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_FixedDecimalPlaces( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::GetCustomListContents( 
            /* [in] */ long ListNum,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::GetCustomListNum( 
            /* [in] */ VARIANT ListArray,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::GetOpenFilename( 
            /* [optional][in] */ VARIANT FileFilter,
            /* [optional][in] */ VARIANT FilterIndex,
            /* [optional][in] */ VARIANT Title,
            /* [optional][in] */ VARIANT ButtonText,
            /* [optional][in] */ VARIANT MultiSelect,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::GetSaveAsFilename( 
            /* [optional][in] */ VARIANT InitialFilename,
            /* [optional][in] */ VARIANT FileFilter,
            /* [optional][in] */ VARIANT FilterIndex,
            /* [optional][in] */ VARIANT Title,
            /* [optional][in] */ VARIANT ButtonText,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Goto( 
            /* [optional][in] */ VARIANT Reference,
            /* [optional][in] */ VARIANT Scroll,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Height( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Height( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Help( 
            /* [optional][in] */ VARIANT HelpFile,
            /* [optional][in] */ VARIANT HelpContextID,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_IgnoreRemoteRequests( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_IgnoreRemoteRequests( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::InchesToPoints( 
            /* [in] */ double Inches,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::InputBox( 
            /* [in] */ BSTR Prompt,
            /* [optional][in] */ VARIANT Title,
            /* [optional][in] */ VARIANT Default,
            /* [optional][in] */ VARIANT Left,
            /* [optional][in] */ VARIANT Top,
            /* [optional][in] */ VARIANT HelpFile,
            /* [optional][in] */ VARIANT HelpContextID,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Interactive( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Interactive( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_International( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Iteration( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Iteration( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_LargeButtons( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_LargeButtons( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Left( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Left( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_LibraryPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::MacroOptions( 
            /* [optional][in] */ VARIANT Macro,
            /* [optional][in] */ VARIANT Description,
            /* [optional][in] */ VARIANT HasMenu,
            /* [optional][in] */ VARIANT MenuText,
            /* [optional][in] */ VARIANT HasShortcutKey,
            /* [optional][in] */ VARIANT ShortcutKey,
            /* [optional][in] */ VARIANT Category,
            /* [optional][in] */ VARIANT StatusBar,
            /* [optional][in] */ VARIANT HelpContextID,
            /* [optional][in] */ VARIANT HelpFile,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::MailLogoff( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::MailLogon( 
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT Password,
            /* [optional][in] */ VARIANT DownloadNewMail,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MailSession( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MailSystem( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlMailSystem *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MathCoprocessorAvailable( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MaxChange( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_MaxChange( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MaxIterations( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_MaxIterations( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MemoryFree( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MemoryTotal( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MemoryUsed( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MouseAvailable( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MoveAfterReturn( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_MoveAfterReturn( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MoveAfterReturnDirection( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlDirection *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_MoveAfterReturnDirection( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlDirection RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_RecentFiles( 
            /* [retval][out] */ RecentFiles **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Name( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::NextLetter( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ Workbook **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_NetworkTemplatesPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ODBCErrors( 
            /* [retval][out] */ ODBCErrors **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ODBCTimeout( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ODBCTimeout( 
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OnCalculate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_OnCalculate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OnData( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_OnData( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_OnDoubleClick( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_OnEntry( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::OnKey( 
            /* [in] */ BSTR Key,
            /* [optional][in] */ VARIANT Procedure,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::OnRepeat( 
            /* [in] */ BSTR Text,
            /* [in] */ BSTR Procedure,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_OnSheetActivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_OnSheetDeactivate( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::OnTime( 
            /* [in] */ VARIANT EarliestTime,
            /* [in] */ BSTR Procedure,
            /* [optional][in] */ VARIANT LatestTime,
            /* [optional][in] */ VARIANT Schedule,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::OnUndo( 
            /* [in] */ BSTR Text,
            /* [in] */ BSTR Procedure,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OnWindow( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_OnWindow( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OperatingSystem( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OrganizationName( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Path( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_PathSeparator( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_PreviousSelections( 
            /* [optional][in] */ VARIANT Index,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_PivotTableSelection( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_PivotTableSelection( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_PromptForSummaryInfo( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_PromptForSummaryInfo( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::Quit( void)
{
    TRACE_IN;
    
    HRESULT hr;

    /*Close WorkBooks*/
    hr = m_workbooks.Close( 0 );

    if ( FAILED( hr ) )
    {
        ERR( " workbooks close all \n" );     
    }

       
    hr = m_oo_desktop.terminate();

    if ( FAILED( hr ) )
    {
        ERR( " m_oo_desktop.terminate() \n" );     
    }

    TRACE_OUT;
    return hr;               
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::RecordMacro( 
            /* [optional][in] */ VARIANT BasicCode,
            /* [optional][in] */ VARIANT XlmCode,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_RecordRelative( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ReferenceStyle( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlReferenceStyle *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ReferenceStyle( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlReferenceStyle RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_RegisteredFunctions( 
            /* [optional][in] */ VARIANT Index1,
            /* [optional][in] */ VARIANT Index2,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::RegisterXLL( 
            /* [in] */ BSTR Filename,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Repeat( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::ResetTipWizard( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_RollZoom( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_RollZoom( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Save( 
            /* [optional][in] */ VARIANT Filename,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::SaveWorkspace( 
            /* [optional][in] */ VARIANT Filename,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ScreenUpdating( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ScreenUpdating( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::SetDefaultChart( 
            /* [optional][in] */ VARIANT FormatName,
            /* [optional][in] */ VARIANT Gallery)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_SheetsInNewWorkbook( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_IN;
   
   *RHS =  m_l_sheetsinnewworkbook;
   
   TRACE_OUT;
   return S_OK;             
}
        
HRESULT STDMETHODCALLTYPE Application::put_SheetsInNewWorkbook( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_IN;
   
   m_l_sheetsinnewworkbook = RHS;
   
   TRACE_OUT;
   return S_OK;            
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowChartTipNames( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowChartTipNames( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowChartTipValues( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowChartTipValues( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_StandardFont( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_StandardFont( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_StandardFontSize( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_StandardFontSize( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_StartupPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_StatusBar( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_StatusBar( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_TemplatesPath( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowToolTips( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowToolTips( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Top( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Top( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DefaultSaveFormat( 
            /* [retval][out] */ XlFileFormat *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DefaultSaveFormat( 
            /* [in] */ XlFileFormat RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_TransitionMenuKey( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_TransitionMenuKey( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_TransitionMenuKeyAction( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_TransitionMenuKeyAction( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_TransitionNavigKeys( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_TransitionNavigKeys( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Undo( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UsableHeight( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UsableWidth( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UserControl( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_UserControl( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UserName( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_UserName( 
            /* [lcid][in] */ long lcid,
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Value( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_VBE( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Version( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
HRESULT STDMETHODCALLTYPE Application::get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_IN;
   
   *RHS = m_b_visible;
   
   TRACE_OUT;
   return S_OK;             
}
        
HRESULT STDMETHODCALLTYPE Application::put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_IN;
   
   m_b_visible = RHS;
   
   HRESULT hr = m_workbooks.Put_Visible( RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " \n" );     
   }
   
   TRACE_OUT;
   return hr;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Volatile( 
            /* [optional][in] */ VARIANT Volatile,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::_Wait( 
            /* [in] */ VARIANT Time,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Width( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ double *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_Width( 
            /* [lcid][in] */ long lcid,
            /* [in] */ double RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_WindowsForPens( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_WindowState( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlWindowState *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_WindowState( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlWindowState RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UILanguage( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_UILanguage( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DefaultSheetDirection( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DefaultSheetDirection( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CursorMovement( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_CursorMovement( 
            /* [lcid][in] */ long lcid,
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ControlCharacters( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ControlCharacters( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::_WSFunction( 
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
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableEvents( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableEvents( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayInfoWindow( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayInfoWindow( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::Wait( 
            /* [in] */ VARIANT Time,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ExtendList( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ExtendList( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_OLEDBErrors( 
            /* [retval][out] */ OLEDBErrors **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::GetPhonetic( 
            /* [optional][in] */ VARIANT Text,
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_COMAddIns( 
            /* [retval][out] */ COMAddIns **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DefaultWebOptions( 
            /* [retval][out] */ DefaultWebOptions **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ProductCode( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UserLibraryPath( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AutoPercentEntry( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_AutoPercentEntry( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_LanguageSettings( 
            /* [retval][out] */ LanguageSettings **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Dummy101( 
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy12( 
            /* [in] */ PivotTable *p1,
            /* [in] */ PivotTable *p2)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AnswerWizard( 
            /* [retval][out] */ AnswerWizard **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::CalculateFull( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::FindFile( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CalculationVersion( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowWindowsInTaskbar( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowWindowsInTaskbar( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FeatureInstall( 
            /* [retval][out] */ MsoFeatureInstall *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_FeatureInstall( 
            /* [in] */ MsoFeatureInstall RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Ready( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy13( 
            /* [in] */ VARIANT Arg1,
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
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FindFormat( 
            /* [retval][out] */ CellFormat **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propputref][id] */ HRESULT STDMETHODCALLTYPE Application::putref_FindFormat( 
            /* [in] */ CellFormat *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ReplaceFormat( 
            /* [retval][out] */ CellFormat **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propputref][id] */ HRESULT STDMETHODCALLTYPE Application::putref_ReplaceFormat( 
            /* [in] */ CellFormat *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UsedObjects( 
            /* [retval][out] */ UsedObjects **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CalculationState( 
            /* [retval][out] */ XlCalculationState *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_CalculationInterruptKey( 
            /* [retval][out] */ XlCalculationInterruptKey *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_CalculationInterruptKey( 
            /* [in] */ XlCalculationInterruptKey RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Watches( 
            /* [retval][out] */ Watches **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayFunctionToolTips( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayFunctionToolTips( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AutomationSecurity( 
            /* [retval][out] */ MsoAutomationSecurity *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_AutomationSecurity( 
            /* [in] */ MsoAutomationSecurity RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FileDialog( 
            /* [in] */ MsoFileDialogType fileDialogType,
            /* [retval][out] */ FileDialog **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy14( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::CalculateFullRebuild( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayPasteOptions( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayPasteOptions( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayInsertOptions( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayInsertOptions( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_GenerateGetPivotData( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_GenerateGetPivotData( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AutoRecover( 
            /* [retval][out] */ AutoRecover **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Hwnd( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Hinstance( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::CheckAbort( 
            /* [optional][in] */ VARIANT KeepAbort)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ErrorCheckingOptions( 
            /* [retval][out] */ ErrorCheckingOptions **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AutoFormatAsYouTypeReplaceHyperlinks( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_AutoFormatAsYouTypeReplaceHyperlinks( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_SmartTagRecognizers( 
            /* [retval][out] */ SmartTagRecognizers **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_NewWorkbook( 
            /* [retval][out] */ NewFile **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_SpellingOptions( 
            /* [retval][out] */ SpellingOptions **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Speech( 
            /* [retval][out] */ Speech **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MapPaperSize( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_MapPaperSize( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowStartupDialog( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowStartupDialog( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DecimalSeparator( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DecimalSeparator( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ThousandsSeparator( 
            /* [retval][out] */ BSTR *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ThousandsSeparator( 
            /* [in] */ BSTR RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_UseSystemSeparators( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_UseSystemSeparators( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ThisCell( 
            /* [retval][out] */ Range **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_RTD( 
            /* [retval][out] */ RTD **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayDocumentActionTaskPane( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayDocumentActionTaskPane( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::DisplayXMLSourcePane( 
            /* [optional][in] */ VARIANT XmlMap)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ArbitraryXMLSupportAvailable( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Support( 
            /* [in] */ IDispatch *Object,
            /* [in] */ long ID,
            /* [optional][in] */ VARIANT arg,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][hidden][id] */ HRESULT STDMETHODCALLTYPE Application::Dummy20( 
            /* [in] */ long grfCompareFunctions,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_MeasurementUnit( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_MeasurementUnit( 
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowSelectionFloaties( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowSelectionFloaties( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowMenuFloaties( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowMenuFloaties( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_ShowDevTools( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_ShowDevTools( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableLivePreview( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableLivePreview( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayDocumentInformationPanel( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayDocumentInformationPanel( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_AlwaysUseClearType( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_AlwaysUseClearType( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_WarnOnFunctionNameConflict( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_WarnOnFunctionNameConflict( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_FormulaBarHeight( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_FormulaBarHeight( 
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DisplayFormulaAutoComplete( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_DisplayFormulaAutoComplete( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_GenerateTableRefs( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ XlGenerateTableRefs *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_GenerateTableRefs( 
            /* [lcid][in] */ long lcid,
            /* [in] */ XlGenerateTableRefs RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_Assistance( 
            /* [retval][out] */ IAssistance **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][id] */ HRESULT STDMETHODCALLTYPE Application::CalculateUntilAsyncQueriesDone( void)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_EnableLargeOperationAlert( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_EnableLargeOperationAlert( 
            /* [in] */ VARIANT_BOOL RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_LargeOperationCellThousandCount( 
            /* [retval][out] */ long *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propput][id] */ HRESULT STDMETHODCALLTYPE Application::put_LargeOperationCellThousandCount( 
            /* [in] */ long RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;             
}
        
         /* [helpcontext][propget][id] */ HRESULT STDMETHODCALLTYPE Application::get_DeferAsyncQueries( 
            /* [retval][out] */ VARIANT_BOOL *RHS)
{
                          
}
        
  HRESULT STDMETHODCALLTYPE Application::put_DeferAsyncQueries( 
             VARIANT_BOOL RHS)
{
                          
}
        
  HRESULT STDMETHODCALLTYPE Application::get_MultiThreadedCalculation( 
             MultiThreadedCalculation **RHS)
{
                          
}
        
 HRESULT STDMETHODCALLTYPE Application::SharePointVersion( 
             BSTR bstrUrl,
             long *RHS)
{
                          
}
        
 HRESULT STDMETHODCALLTYPE Application::get_ActiveEncryptionSession( 
             long *RHS)
{
                          
}
        
HRESULT STDMETHODCALLTYPE Application::get_HighQualityModeForGraphics( 
             VARIANT_BOOL *RHS)
{
                          
}
        
HRESULT STDMETHODCALLTYPE Application::put_HighQualityModeForGraphics( 
             VARIANT_BOOL RHS)
{
                          
}

HRESULT Application::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID__Application, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }
        
     m_b_screenupdating      = VARIANT_TRUE; 
     m_b_displayalerts       = VARIANT_TRUE;
     m_b_visible             = VARIANT_FALSE;
     m_l_sheetsinnewworkbook = 1;  
        
     m_workbooks.Put_Application( (void*)this );
     m_workbooks.Put_Parent( (void*)this );   
        
     // Start OpenOffice
        
     hr = m_oo_service_manager.Get_Desktop( m_oo_desktop );    
     if ( FAILED( hr ) )
     {
	     ERR( " m_oo_service_manager.Get_Desktop \n" ); 	  
     }
     
     /*
     TODO:  сделать определение версии OpenOffice
     */   
     OOVersion = VER_3;           
                
     return ( hr );
}

OOServiceManager& Application::GetOOServiceManager( )
{
    return ( m_oo_service_manager ); 	  
}

OODesktop& Application::GetOODesktop( )
{
    return ( m_oo_desktop ); 	  
}
