/*
 * implementation of Sheets
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

#include "sheets.h"

#include "application.h"
#include "workbook.h"
#include "worksheet.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CSheets::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<Sheets*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<Sheets*>(this));
    }     
    
    if ( iid == IID_Sheets ) {
        TRACE("Sheets \n");
        *ppv = static_cast<Sheets*>(this);
    } 

    if ( iid == IID_IWorksheets ) {
        TRACE("IWorksheets \n");
        *ppv = static_cast<IWorksheets*>(this);
    } 
    
    if ( iid == DIID_Worksheets ) {
        TRACE("Worksheets \n");
        *ppv = static_cast<Worksheets*>(this);
    } 
    
    if ( iid == IID_IEnumVARIANT ) {
        TRACE(" IEnumVARIANT \n");
        *ppv = static_cast<IEnumVARIANT*>(this);
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

ULONG STDMETHODCALLTYPE CSheets::AddRef()
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);              
}

ULONG STDMETHODCALLTYPE CSheets::Release()
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
HRESULT STDMETHODCALLTYPE CSheets::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;          
}

HRESULT STDMETHODCALLTYPE CSheets::GetTypeInfo(
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

HRESULT STDMETHODCALLTYPE CSheets::GetIDsOfNames(
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

HRESULT STDMETHODCALLTYPE CSheets::Invoke(
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
                 static_cast<IDispatch*>(static_cast<Sheets*>(this)), 
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
               
        // Sheets
HRESULT STDMETHODCALLTYPE CSheets::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE CSheets::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Parent( 
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
   
   hr = (static_cast<Workbook*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;          
}
        
HRESULT STDMETHODCALLTYPE CSheets::Add( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [optional][in] */ VARIANT Count,
            /* [optional][in] */ VARIANT Type,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;
   HRESULT hr;
   long index;
   BSTR new_name;
   BSTR tmp_str;
   long count;
   
   enum
   {
       none,
       before,
       after,    
   } e_type_add = none;
   
      
   CorrectArg(Before, &Before);
   CorrectArg(After, &After);
   CorrectArg(Count, &Count);
   CorrectArg(Type, &Type);
   
   if ( Is_Variant_Null(Before) ) {
       VariantClear(&Before);
   } else {
      // Convert to VT_I4
      hr = VariantChangeTypeEx(&Before, &Before, 0, 0, VT_I4);
      
      e_type_add = before;
   }
   
   if ( Is_Variant_Null(After) ) {
       VariantClear(&After);
   } else {
       // Convert to VT_I4
       hr = VariantChangeTypeEx(&After, &After, 0, 0, VT_I4);
 
       e_type_add = after;
   }
   
   if ( Is_Variant_Null(Count) ) {
       VariantClear(&Count);
       V_VT(&Count) = VT_I4;
       V_I4(&Count) = 1;
   } else {
       // Convert to VT_I4
       hr = VariantChangeTypeEx(&Count, &Count, 0, 0, VT_I4);
       if ( FAILED( hr ) ) {
           ERR(" VariantChangeTypeEx -Count- \n");
       }
   }
   
   if ( Is_Variant_Null( Type ) ) {
       VariantClear( &Type );
       V_VT(&Type) = VT_I4;
       V_I4(&Type) = xlWorksheet;
   } else {
       // Convert to VT_I4
       hr = VariantChangeTypeEx(&Type, &Type, 0, 0, VT_I4);
       if (FAILED(hr)) {
           ERR(" VariantChangeTypeEx -Type-\n");
       }
       // only xlWorksheet are supported
       switch ( V_I4( &Type ) ) 
       {
       case xlWorksheet: 
            break;
       default :
           ERR(" This Type not implemented type = %i \n", V_I4(&Type) );
           return E_FAIL;
       }
   }
   
   index = 0;
   
   // get Count of worksheets
   hr = get_Count( &count );
   if ( FAILED( hr ) )
   {
       ERR( " get_Count \n" );     
   }
   
   switch ( e_type_add )
   {
   case before:
        {
            WTRACE(L" before element %s \n",V_BSTR(&Before) );
            if ( V_VT(&Before) == VT_I4 ) {
                index = V_I4(&Before) - 1;
            } else {                   
                int i = FindIndexWorksheetByName( V_BSTR(&Before) );
                
                if ( i >= 0 ) 
                    index = i; 
                else 
                    index = 0;
            }   
        }
        break;
        
   case after:
        {
            WTRACE( L"after element %s\n", V_BSTR( &After ) );
            if ( V_VT(&After) == VT_I4 ) {
               index = V_I4(&After);
            } else {
               int i = FindIndexWorksheetByName( V_BSTR(&After) );
               
               if ( i >= 0 ) 
                   index = i+1; 
               else 
                   index = 0;
            } 
        }
        break;
        
   case none:
        {
            TRACE(" to the begining of the list \n");
            index = 0;    
        }
        break;        
                   
   }
  
   for ( int i = V_I4( &Count ); i > 0; i--) 
   {      
       int j = 0;
       do 
       {          
           SysFreeString( new_name );
           new_name = SysAllocString( L"Sheet" );
 
           hr = VarBstrFromI4( count + i + j, 0, 0, &tmp_str);
           
           if ( FAILED( hr ) ) {
                ERR( " VarBSTRFromI4 \n" );
                tmp_str = SysAllocString( L"4" );
           }
          
           VarBstrCat( new_name, tmp_str, &new_name );
          
           SysFreeString(tmp_str);
          
           j++;
          
           VARIANT param1;
           IDispatch *p_disp = NULL;
          
           VariantInit( &param1 );
           V_VT( &param1 )   = VT_BSTR;
           V_BSTR( &param1 ) = SysAllocString( new_name );
          
           hr = get__Default( param1, &p_disp );
                    
           if ( p_disp != NULL ) 
           {
               p_disp->Release(); 
          
               p_disp = NULL;
           }
          
           VariantClear( &param1 );
          
       } while ( !FAILED( hr ) );

       hr = m_oo_sheets.insertNewByName( new_name, index );
       if ( FAILED( hr ) )
       {
           ERR( " m_pd_sheets.insertNewByName \n" ); 
           SysFreeString(new_name);
           TRACE_OUT;
           return ( hr );     
       }
 
       SysFreeString(new_name);
        
   } // for( i = V_I4( &Count ); i > 0; i--) 
   
   index++;
   
   VARIANT param1;
   VariantInit( &param1 );
   V_VT( &param1 ) = VT_I4;
   V_I4( &param1 ) = index;
   
   hr = get__Default( param1, RHS );
   if ( FAILED( hr ) )
   {
       ERR( " get__Default \n" );     
   } else
   {
       hr = reinterpret_cast<Worksheet*>( *RHS )->Activate( 0 );       
       if ( FAILED( hr ) )
       {
           ERR( " Activate() \n" );     
           reinterpret_cast<Worksheet*>( *RHS )->Release();
           *RHS = NULL;
       }
   }
     
   VariantClear( &param1 );
   
   TRACE_OUT;
   return ( hr );            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Copy( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Count( 
            /* [retval][out] */ long *RHS)
{
    TRACE_IN;
    HRESULT hr;

    *RHS = m_oo_sheets.getCount();

    if ( *RHS < 0 )
    {
        ERR( "\n" );
        *RHS = 0;
        hr = E_FAIL;     
    } 

    TRACE_OUT;
    return ( hr );             
}
        
HRESULT STDMETHODCALLTYPE CSheets::Delete( 
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::FillAcrossSheets( 
            /* [in] */ Range	*Range,
            /* [defaultvalue][optional][in] */ XlFillWith Type,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Item( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS)
{
   TRACE_IN;
   
   HRESULT hr = get__Default( Index, RHS );
   
   if ( FAILED( hr ) )
   {
       ERR( " call get__Default \n" );     
   }
   
   TRACE_OUT;
   return ( hr );            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Move( 
            /* [optional][in] */ VARIANT Before,
            /* [optional][in] */ VARIANT After,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS)
{
   TRACE_IN;
   
   HRESULT hr = S_OK;
   
   hr = QueryInterface( IID_IEnumVARIANT, (void**)RHS );
   
   if ( FAILED( hr ) )
   {
        ERR( " FAILED get IID_IEnumVARIANT \n" );    
   }
   
   TRACE_OUT;
   return hr;           
}
        
HRESULT STDMETHODCALLTYPE CSheets::__PrintOut( 
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
        
HRESULT STDMETHODCALLTYPE CSheets::PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::Select( 
            /* [optional][in] */ VARIANT Replace,
            /* [lcid][in] */ long lcid)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_HPageBreaks( 
            /* [retval][out] */ HPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_VPageBreaks( 
            /* [retval][out] */ vPageBreaks **RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get_Visible( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::put_Visible( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS)
{
   TRACE_NOTIMPL;
   return E_NOTIMPL;            
}
        
HRESULT STDMETHODCALLTYPE CSheets::get__Default( 
            /* [in] */ VARIANT Index,
            /* [retval][out] */ IDispatch **RHS)
{
    TRACE_IN;        
            
    HRESULT hr;
    
    CorrectArg( Index, &Index );
     
    // Try to change type to VT_I4 (long)        
    hr = VariantChangeTypeEx(&Index, &Index, 0, 0, VT_I4);        
            
    switch ( V_VT( &Index ) )
    {
    case VT_I4:
         {
             long index = V_I4( &Index );
             
             // we need to do that, bacause
             // Excel start from 1
             // but openoffice start from 0.... 
             index--;
             
             Worksheet* p_worksheet = new Worksheet;
             
             p_worksheet->Put_Application( m_p_application );
             p_worksheet->Put_Parent( this );
             
             OOSheet oo_sheet;
             
             hr = S_OK;
             
             oo_sheet = m_oo_sheets.getByIndex( index ); 
             
             if ( oo_sheet.IsNull() )
             {
                ERR( " m_oo_sheets.getByIndex \n" );
                if ( p_worksheet != NULL )
                    p_worksheet->Release();
                    
                TRACE_OUT;    
                return ( E_FAIL ); 	  
		     }
             
             p_worksheet->InitWrapper( oo_sheet );
             
             hr = p_worksheet->QueryInterface( IID_IDispatch, (void**)RHS );
             
             if ( FAILED( hr ) )
             {
                 ERR( " worksheet.QueryInterface \n" );     
             }
             
             if ( p_worksheet != NULL )
                 p_worksheet->Release();
                      
         }
         break;
         
    case VT_BSTR:
         {            
             Worksheet* p_worksheet = new Worksheet;
             
             p_worksheet->Put_Application( m_p_application );
             p_worksheet->Put_Parent( this );
             
             OOSheet oo_sheet;
             
             hr = S_OK;
             
             oo_sheet = m_oo_sheets.getByName( V_BSTR( &Index ) ); 
             
             if ( oo_sheet.IsNull() )
             {
                ERR( " m_oo_sheets.getByName \n" );
                if ( p_worksheet != NULL )
                    p_worksheet->Release();
                    
                TRACE_OUT;    
                return ( E_FAIL ); 	  
		     }

             p_worksheet->InitWrapper( oo_sheet );
             
             hr = p_worksheet->QueryInterface( IID_IDispatch, (void**)RHS );
             
             if ( FAILED( hr ) )
             {
                 ERR( " worksheet.QueryInterface \n" );     
             } 
             
             p_worksheet->Release();
                       
         }
         break;
         
    default:
         {
             ERR( " Unknown type of Index     V_VT( Index ) = %i \n", V_VT( &Index ) );
             *RHS = NULL;
             hr = E_FAIL;            
         }
         break;       
    }        
                  
    TRACE_OUT;
    return ( hr );            
}
        
HRESULT STDMETHODCALLTYPE CSheets::_PrintOut( 
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
        
HRESULT STDMETHODCALLTYPE CSheets::PrintOut( 
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
            
HRESULT CSheets::Next ( ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched)
{
    TRACE_IN;    
        
    HRESULT hr;
    ULONG l;
    long l1;
    long count = 0;
    ULONG l2;
    IDispatch *dret;
    VARIANT varindex, vNull;

    VariantInit(&vNull);
    V_VT(&vNull) = VT_NULL;

    if ( enum_position < 0 )
    {
        ERR( " enum_position < 0 \n" );
        return ( S_FALSE );
    }
    
    if ( pCeltFetched != NULL )
    {
       *pCeltFetched = 0;
    }
    
    if ( rgVar == NULL )
    {
        ERR( " rgVar == NULL \n" );
        return E_INVALIDARG;
    }

    VariantInit( &varindex );
    
    /*Init Array*/
    for ( l = 0; l < celt; l++)
       VariantInit( &rgVar[l] );

    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
        ERR( " get_Count \n" ); 
        return (E_FAIL);
    }
    
    V_VT( &varindex ) = VT_I4;

    for ( l1 = enum_position, l2 = 0; l1 < count && l2 < celt; l1++, l2++) {
      V_I4( &varindex ) = l1 + 1;    //Because index of sheets start from 1
      
      hr = get_Item( varindex, &dret);
            
      V_VT( &rgVar[l2] )       = VT_DISPATCH;
      V_DISPATCH( &rgVar[l2] ) = static_cast<IDispatch*>( dret );
      
      if ( FAILED( hr ) )
      {
          ERR( " get_Item \n" );
          goto error;
      }
      
    }

    if (pCeltFetched != NULL)
    {
       *pCeltFetched = l2;
    }
    
    enum_position = l1;
    
    TRACE_OUT;     
    return  ((l2 < celt) ? S_FALSE : S_OK);

error:
      
    for ( l = 0; l < celt; l++)
    {
        VariantClear(&rgVar[l]);
    }
   
    VariantClear( &varindex );
   
    TRACE_OUT;
    return ( hr );       
}
        
HRESULT CSheets::Skip ( ULONG celt)
{
    long count = 0;
    HRESULT hr;
    TRACE_IN;

    hr = get_Count( &count );
    if ( FAILED( hr ) )
    {
        ERR( " get_count \n" );     
    }   
    
    enum_position += celt;

    if ( enum_position >= count) 
    {
        enum_position = count - 1;
        TRACE_OUT;
        return S_FALSE;
    }
    
    TRACE_OUT;
    return S_OK;       
}

HRESULT CSheets::Reset( )
{
   TRACE_IN;
   
   enum_position = 0;
   
   TRACE_OUT;
   return S_OK;       
}

HRESULT CSheets::Clone(IEnumVARIANT** ppEnum)
{
   TRACE_IN;
   
   HRESULT hr = S_OK;
   
   hr = QueryInterface( IID_IEnumVARIANT, (void**)ppEnum );
   
   if ( FAILED( hr ) )
   {
        ERR( " FAILED get IID_IEnumVARIANT \n" );    
   }
   
   TRACE_OUT;
   return hr;        
}





HRESULT CSheets::Init()
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID_Sheets, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;        
}

HRESULT CSheets::Put_Application( void* p_application )
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;      
}

HRESULT CSheets::Put_Parent( void* p_parent )
{
   TRACE_IN;  
      
   m_p_parent = p_parent;
   
   TRACE_OUT;
   return S_OK;     
}

HRESULT CSheets::InitWrapper( OOSheets oo_sheets )
{
    m_oo_sheets = oo_sheets;    
    
    return ( S_OK );    
}

long CSheets::FindIndexWorksheetByName( BSTR _name )
{
    TRACE_IN; 
    HRESULT hr;
    long count;
    VARIANT par_tmp;
    BSTR tmp_name;
    
    VariantInit( &par_tmp );
    
    hr = get_Count( &count);
    if ( FAILED( hr ) ) {
        ERR(" get_Count\n");
        return ( -1 );
    }
     
    int i = 1;
    
    while ( i <= count ) 
    {
        IDispatch* p_disp = NULL;  
        
        VariantClear( &par_tmp );
        V_VT( &par_tmp ) = VT_I4;
        V_I4( &par_tmp ) = i;
        
        hr = get__Default( par_tmp, &p_disp);      
        if ( !FAILED( hr ) )
        {
            hr = reinterpret_cast<Worksheet*>( p_disp )->get_Name( &tmp_name ); 
            if ( !FAILED(hr) ) {
                if ( !lstrcmpiW( tmp_name, _name ) ) {
                     
                    SysFreeString( tmp_name );
                    VariantClear( &par_tmp );
                    
                    if ( p_disp != NULL )
                    {
                        p_disp->Release();
                        p_disp = NULL;     
                    }  
                    
                    TRACE_OUT;
                    return (i - 1);
                }
                
                SysFreeString( tmp_name );
                
            } else
            {
                ERR( " get_Name \n" );
                
                if ( p_disp != NULL )
                {
                    p_disp->Release();
                    p_disp = NULL;     
                } 
            }
             
        } else
        { 
            if ( p_disp != NULL )
            {
                p_disp->Release();
                p_disp = NULL;     
            }      
        }    
        
        VariantClear( &par_tmp );
          
        i++;   
    }
    
    ERR( " NOT FIND \n" );
    TRACE_OUT; 
    return ( -1 );     
}

HRESULT CSheets::RemoveWorksheetByName( BSTR _name )
{
    TRACE_IN;
    HRESULT hr;
    
    hr = m_oo_sheets.removeByName( _name );
    if ( FAILED( hr ) )
    {
        ERR( " m_oo_sheets.removeByName \n" );     
    }

    TRACE_OUT;       
    return ( hr );
} 


