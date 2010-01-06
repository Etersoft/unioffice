/*
 * implementation of Font
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

#include "font.h"
#include "application.h"

       // IUnknown
HRESULT STDMETHODCALLTYPE CFont::QueryInterface(const IID& iid, void** ppv)
{
    *ppv = NULL;    
        
    if ( iid == IID_IUnknown ) {
        TRACE("IUnknown \n");
        *ppv = static_cast<IUnknown*>(static_cast<_IFont*>(this));
    }
        
    if ( iid == IID_IDispatch ) {
        TRACE("IDispatch \n");
        *ppv = static_cast<IDispatch*>(static_cast<_IFont*>(this));
    }     
    
    if ( iid == IID__IFont) {
        TRACE("IRange\n");
        *ppv = static_cast<_IFont*>(this);
    } 
    
    if ( iid == DIID_Font) {
        TRACE("Range \n");
        *ppv = static_cast<Font*>(this);
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
        
ULONG STDMETHODCALLTYPE CFont::AddRef( )
{
      TRACE( " ref = %i \n", m_cRef );
      
      return InterlockedIncrement(&m_cRef);   		
}
        
ULONG STDMETHODCALLTYPE CFont::Release( )
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
HRESULT STDMETHODCALLTYPE CFont::GetTypeInfoCount( UINT * pctinfo )
{
    *pctinfo = 1;
    return S_OK;  		
}
        
HRESULT STDMETHODCALLTYPE CFont::GetTypeInfo(
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
        
HRESULT STDMETHODCALLTYPE CFont::GetIDsOfNames(
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
        
HRESULT STDMETHODCALLTYPE CFont::Invoke(
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
                 static_cast<IDispatch*>(static_cast<_IFont*>(this)), 
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
HRESULT STDMETHODCALLTYPE CFont::get_Application( 
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
        
HRESULT STDMETHODCALLTYPE CFont::get_Creator( 
            /* [retval][out] */ XlCreator *RHS)
{
   TRACE_IN;
   
   *RHS = xlCreatorCode;
   
   TRACE_OUT;
   return S_OK;  		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Parent( 
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
   
   hr = (static_cast<IUnknown*>( m_p_parent ))->QueryInterface( IID_IDispatch,(void**)RHS );          
             
   TRACE_OUT;
   return hr;   		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Background( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Background( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Bold( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    float value = 0.0f;
    		
	hr = m_oo_font.getCharWeight( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.getCharWeight \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_BOOL;
    if ( value > 140 )
	    V_BOOL( RHS ) = VARIANT_TRUE;
	else
	    V_BOOL( RHS ) = VARIANT_FALSE;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Bold( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    float value = 0.0f;
    
    CorrectArg(RHS, &RHS);
    
    VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    
    if ( V_BOOL(&RHS) == VARIANT_TRUE )
	    value = 150;
	else
	    value = 100;
		
	hr = m_oo_font.setCharWeight( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.setCharWeight \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Color( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    		
	hr = m_oo_font.getCharColor( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.getCharColor \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = value;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Color( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    
    CorrectArg(RHS, &RHS);
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    if ( FAILED( hr ) )
    {
	    ERR( " VariantChangeTypeEx \n" );   
		TRACE_OUT;
		return ( E_FAIL );	 
    }
    
	value = V_I4( &RHS );
		
	hr = m_oo_font.setCharColor( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.setCharColor \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_ColorIndex( 
            /* [retval][out] */ VARIANT *RHS)
{
 	TRACE_IN;
	HRESULT hr;
    long tmpcolor;
    VARIANT var_color;
    
    VariantInit( &var_color );

    hr = get_Color( &var_color );
    if ( FAILED( hr ) ) 
	{
        ERR(" failed get_Color \n");              
        TRACE_OUT;
        return ( hr );
    }
    
    hr = VariantChangeTypeEx(&var_color, &var_color, 0, 0, VT_I4);
    if ( FAILED( hr ) )
    {
	    ERR( " VariantChangeTypeEx \n" );   
		TRACE_OUT;
		return ( E_FAIL );	 
    }
    
    tmpcolor = V_I4( &var_color );
    
    VariantInit( RHS );
    V_VT( RHS ) = VT_I4;
    
    for ( int i=0; i < 56; i++ )
    {
        if ( color[i] == tmpcolor ) 
		{
            V_I4( RHS ) = i+1;
			TRACE_OUT;
            return ( S_OK );
        }
    }

    ERR(" Color don`t have colorindex put defaut and return S_OK \n ");
    V_I4( RHS ) = 1;
					
    TRACE_OUT;
    return ( S_OK );  		
}

HRESULT STDMETHODCALLTYPE CFont::put_ColorIndex( 
            /* [in] */ VARIANT RHS)
{
 	TRACE_IN;
	HRESULT hr;		
    long tmpcolor;
    VARIANT var_color;
    
    CorrectArg(RHS, &RHS);
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    if ( FAILED( hr ) )
    {
	    ERR( " VariantChangeTypeEx \n" );   
		TRACE_OUT;
		return ( E_FAIL );	 
    }
    
	tmpcolor = V_I4( &RHS );
    
    if ( tmpcolor == xlColorIndexNone ) 
	    tmpcolor = 2;
    if ( tmpcolor == xlColorIndexAutomatic ) 
	    tmpcolor = 1;

    if ( 
	    ( tmpcolor < 1 ) ||
		( tmpcolor > 56 ) 
		) 
	{
        ERR(" Incorrect colorindex \n ");
        TRACE_OUT;
        return ( S_OK );
    } 	
	
	VariantInit( &var_color );
	V_VT( &var_color ) = VT_I4;
	V_I4( &var_color ) = color[ tmpcolor - 1 ];
	
	hr = put_Color( var_color );
	if ( FAILED( hr ) )
	{
	    ERR( " failed put_Color " );   	  
    }   
 
    VariantClear( &var_color );
    TRACE_OUT;
	return ( hr );    
}
        
HRESULT STDMETHODCALLTYPE CFont::get_FontStyle( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    VARIANT tmp;
    WCHAR str[200];
    bool pusto = true;
    
	VariantInit( &tmp );

	V_VT( &tmp ) = VT_BOOL;
	V_BOOL( &tmp ) = VARIANT_FALSE;
    
    hr = get_Bold( &tmp );
    if ( FAILED( hr ) ) {
        ERR(" when get_Bold ");
    }
    
    if ( V_BOOL( &tmp ) == VARIANT_TRUE ) 
	{
        if ( pusto ) 
		    swprintf(str, L"%s", L"bold");
        else 
		    swprintf(str, L"%s %s", str, L"bold");
		    
        pusto = false;
    }

	V_BOOL( &tmp ) = VARIANT_FALSE;
    
    hr = get_Italic( &tmp );
    if ( FAILED( hr ) ) {
        ERR(" when get_Italic ");
    }
 
    if ( V_BOOL( &tmp ) == VARIANT_TRUE ) 
	{
        if ( pusto ) 
		    swprintf(str, L"%s", L"italic");
        else 
			swprintf(str, L"%s %s", str, L"italic");
			
        pusto = false;
    }

    if ( pusto ) 
	    swprintf( str, L"%s", L"normal");

    V_VT( RHS ) = VT_BSTR;
    V_BSTR( RHS ) = SysAllocString(str);

    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_FontStyle( 
            /* [in] */ VARIANT RHS)
{
 	TRACE_IN;
 	HRESULT hr;
 	
    static WCHAR str_bold_en[] = {
        'b','o','l','d',0};
    static WCHAR str_italic_en[] = {
        'i','t','a','l','i','c',0};
    static WCHAR str_bold_ru[] = {
        0x0436,0x0438,0x0440,0x043d,0x044b, 0x0439,0};
    static WCHAR str_italic_ru[] = {
        0x043a,0x0443,0x0440,0x0441,0x0438, 0x0432,0};
    static WCHAR str_bold2_ru[] = {
        0x043f, 0x043e, 0x043b, 0x0443,0x0436,0x0438,0x0440,0x043d,0x044b, 0x0439,0};

    VARIANT tmp;
    int i = 0;
    WCHAR str[100];

    CorrectArg(RHS, &RHS);		
 			
 	if ( V_VT( &RHS ) != VT_BSTR) 
	{
        ERR(" parameter not BSTR ");
        TRACE_OUT;
        return ( E_FAIL) ;
    }
	
	VariantInit( &tmp );
	
	V_VT( &tmp ) = VT_BOOL;
	V_BOOL( &tmp ) = VARIANT_TRUE;
	
    str[0] = 0;
    while (*(V_BSTR(&RHS)+i)) {
        if (*(V_BSTR(&RHS)+i)==L' ') {
            if ((!lstrcmpiW(str, str_bold_en)) ||
                (!lstrcmpiW(str, str_bold_ru)) ||
                (!lstrcmpiW(str, str_bold2_ru))) {
                 put_Bold(tmp);
            }
            if ((!lstrcmpiW(str, str_italic_en)) ||
                (!lstrcmpiW(str, str_italic_ru))) {
                 put_Italic(tmp);
            }
            str[0] = 0;
        } else {
            swprintf(str, L"%s%c",str, *(V_BSTR(&RHS)+i));
        }
        i++;
    }
    
    if ((!lstrcmpiW(str, str_bold_en)) ||
        (!lstrcmpiW(str, str_bold_ru)) ||
        (!lstrcmpiW(str, str_bold2_ru)))  {
         put_Bold(tmp);
    }
    if ((!lstrcmpiW(str, str_italic_en)) ||
        (!lstrcmpiW(str, str_italic_ru))) {
         put_Italic(tmp);
    }	 	 	
 			
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Italic( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::awt::FontSlant value = com::sun::star::awt::NONE;
    		
	hr = m_oo_font.getCharPosture( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.getCharPosture \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_BOOL;
    if ( value == com::sun::star::awt::ITALIC )
	    V_BOOL( RHS ) = VARIANT_TRUE;
	else
	    V_BOOL( RHS ) = VARIANT_FALSE;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Italic( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    com::sun::star::awt::FontSlant value = com::sun::star::awt::NONE;
    
    CorrectArg(RHS, &RHS);
    
    VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    
    if ( V_BOOL(&RHS) == VARIANT_TRUE )
	    value = com::sun::star::awt::ITALIC;
	else
	    value = com::sun::star::awt::NONE;
		
	hr = m_oo_font.setCharPosture( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.setCharPosture \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Name( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    BSTR font_name = SysAllocString( L"" );
    
    hr = m_oo_font.getCharFontName( font_name );
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_font.getCharFontName \n" );
		TRACE_OUT;
		return ( hr );   	 
  	}
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_BSTR;
    V_BSTR( RHS ) = SysAllocString( font_name );
    
    SysFreeString( font_name );
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Name( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    
    CorrectArg(RHS, &RHS);
    
    hr = m_oo_font.setCharFontName( V_BSTR( &RHS ) );
    if ( FAILED( hr ) )
    {
	    ERR( " m_oo_font.setCharFontName \n" );
		TRACE_OUT;
		return ( hr );   	 
  	}
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_OutlineFont( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_STUB;
    V_VT(RHS) = VT_BOOL;
    V_BOOL(RHS) = VARIANT_FALSE;
    return ( S_OK );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_OutlineFont( 
            /* [in] */ VARIANT RHS)
{
    TRACE_STUB;
    return ( S_OK );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Shadow( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    bool value = false;
    		
	hr = m_oo_font.getCharShadowed( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.getCharShadowed \n" );  
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
        
HRESULT STDMETHODCALLTYPE CFont::put_Shadow( 
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
		
	hr = m_oo_font.setCharShadowed( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.setCharShadowed \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Size( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    		
	hr = m_oo_font.getCharHeight( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.getCharHeight \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    V_I4( RHS ) = value;
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Size( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    
    CorrectArg(RHS, &RHS);
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    if ( FAILED( hr ) )
    {
	    ERR( " VariantChangeTypeEx \n" );   
		TRACE_OUT;
		return ( E_FAIL );	 
    }
    
	value = V_I4( &RHS );
		
	hr = m_oo_font.setCharHeight( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.setCharHeight \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr ); 		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Strikethrough( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    		
	hr = m_oo_font.getCharStrikeout( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.getCharStrikeout \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_BOOL;
    
    switch( value ) 
	{
        case underline_style_SINGLE:
            V_BOOL( RHS ) = VARIANT_TRUE;
            break;
        case underline_style_NONE:
            V_BOOL( RHS ) = VARIANT_FALSE;
            break;
        default:
            ERR(" CharStrikeout\n");
            TRACE_OUT;
            return (E_FAIL);
    }
    
    TRACE_OUT;
    return ( hr );		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Strikethrough( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    
    CorrectArg(RHS, &RHS);
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_BOOL);
    if ( FAILED( hr ) )
    {
	    ERR( " VariantChangeTypeEx \n" );   
		TRACE_OUT;
		return ( E_FAIL );	 
    }
    
    switch ( V_BOOL( &RHS ) ) 
	{
        case VARIANT_TRUE:
        	 value = underline_style_SINGLE;
        	 break;
        case VARIANT_FALSE:
        	 value = strikeout_style_NONE;
        	 break;
    default :
        ERR(" parameters \n");
        TRACE_OUT;
        return ( E_FAIL );
    }
    
	hr = m_oo_font.setCharStrikeout( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.setCharStrikeout \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr ); 	   		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Subscript( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Subscript( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_Superscript( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_Superscript( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
HRESULT STDMETHODCALLTYPE CFont::get_Underline( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    		
	hr = m_oo_font.getCharUnderline( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.getCharUnderline \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    VariantClear( RHS );
    V_VT( RHS ) = VT_I4;
    
    switch( value ) 
	{
        case underline_style_SINGLE:
            V_I4(RHS) = xlUnderlineStyleSingle;
            break;
        case underline_style_DOUBLE:
            V_I4(RHS) = xlUnderlineStyleDouble;
            break;
        case underline_style_NONE:
            V_I4(RHS) = xlUnderlineStyleNone;
            break;
        default:
            ERR(" CharUnderline \n");
            TRACE_OUT;
            return (E_FAIL);
    }
    
    TRACE_OUT;
    return ( hr );  		
}
        
HRESULT STDMETHODCALLTYPE CFont::put_Underline( 
            /* [in] */ VARIANT RHS)
{
    TRACE_IN;
    HRESULT hr;
    long value = 0;
    
    CorrectArg(RHS, &RHS);
    
    hr = VariantChangeTypeEx(&RHS, &RHS, 0, 0, VT_I4);
    if ( FAILED( hr ) )
    {
	    ERR( " VariantChangeTypeEx \n" );   
		TRACE_OUT;
		return ( E_FAIL );	 
    }
    
    switch ( V_I4( &RHS ) ) 
	{
        case xlUnderlineStyleDouble:
        case xlUnderlineStyleDoubleAccounting:
        value = underline_style_DOUBLE;
        break;
        case xlUnderlineStyleNone:
        value = underline_style_NONE;
        break;
        case xlUnderlineStyleSingle:
        case xlUnderlineStyleSingleAccounting:
        value = underline_style_SINGLE;
        break;
    default :
        ERR(" parameters \n");
        TRACE_OUT;
        return ( E_FAIL );
    }
    
	hr = m_oo_font.setCharUnderline( value );
	if ( FAILED( hr ) )
	{
	    ERR( " m_oo_font.setCharUnderline \n" );  
		TRACE_OUT;
		return ( hr ); 	 
    }			    
    
    TRACE_OUT;
    return ( hr ); 		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_ThemeColor( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_ThemeColor( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_TintAndShade( 
            /* [retval][out] */ VARIANT *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_TintAndShade( 
            /* [in] */ VARIANT RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}

        /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE CFont::get_ThemeFont( 
            /* [retval][out] */ XlThemeFont *RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
        
        /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE CFont::put_ThemeFont( 
            /* [in] */ XlThemeFont RHS)
{
    TRACE_NOTIMPL;
    return E_NOTIMPL;  		
}
            
HRESULT CFont::Init( )
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
       hr = pITypeLib->GetTypeInfoOfGuid(IID__IFont, &m_pITypeInfo);
       pITypeLib->Release();
       if (FAILED(hr))
       {
          ERR(" GetTypeInfoOfGuid \n ");
          return hr;
       }
       
     }

     return hr;		
}
         
HRESULT CFont::Put_Application( void* p_application)
{
    TRACE_IN;
    
    m_p_application = p_application;
    
    TRACE_OUT;    
    return S_OK;  		
}
        
HRESULT CFont::Put_Parent( void* p_parent)
{
   TRACE_IN;  
      
   m_p_parent = p_parent;       
      
   TRACE_OUT;
   return S_OK; 		
}
        
HRESULT CFont::InitWrapper( OOFont _oo_font )
{
    m_oo_font = _oo_font;     
}            
            
