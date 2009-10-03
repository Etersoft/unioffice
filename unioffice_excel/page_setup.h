/*
 * header file - PageSetup
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

#ifndef __UNIOFFICE_EXCEL_PAGE_SETUP_H__
#define __UNIOFFICE_EXCEL_PAGE_SETUP_H__

#include "unioffice_excel_private.h"


class CPageSetup : public IPageSetup, public PageSetup
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


        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_BlackAndWhite( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_BlackAndWhite( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_BottomMargin( 
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_BottomMargin( 
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CenterFooter( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_CenterFooter( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CenterHeader( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_CenterHeader( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CenterHorizontally( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_CenterHorizontally( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CenterVertically( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_CenterVertically( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE get_ChartSize( 
            /* [retval][out] */ XlObjectSize *RHS);
        
        virtual /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE put_ChartSize( 
            /* [in] */ XlObjectSize RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Draft( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Draft( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FirstPageNumber( 
            /* [retval][out] */ long *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FirstPageNumber( 
            /* [in] */ long RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FitToPagesTall( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FitToPagesTall( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FitToPagesWide( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FitToPagesWide( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FooterMargin( 
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FooterMargin( 
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_HeaderMargin( 
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_HeaderMargin( 
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_LeftFooter( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_LeftFooter( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_LeftHeader( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_LeftHeader( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_LeftMargin( 
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_LeftMargin( 
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Order( 
            /* [retval][out] */ XlOrder *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Order( 
            /* [in] */ XlOrder RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Orientation( 
            /* [retval][out] */ XlPageOrientation *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Orientation( 
            /* [in] */ XlPageOrientation RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PaperSize( 
            /* [retval][out] */ XlPaperSize *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PaperSize( 
            /* [in] */ XlPaperSize RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintArea( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintArea( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintGridlines( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintGridlines( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintHeadings( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintHeadings( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintNotes( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintNotes( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintQuality( 
            /* [optional][in] */ VARIANT Index,
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintQuality( 
            /* [optional][in] */ VARIANT Index,
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintTitleColumns( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintTitleColumns( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintTitleRows( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintTitleRows( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_RightFooter( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_RightFooter( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_RightHeader( 
            /* [retval][out] */ BSTR *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_RightHeader( 
            /* [in] */ BSTR RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_RightMargin( 
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_RightMargin( 
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_TopMargin( 
            /* [retval][out] */ double *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_TopMargin( 
            /* [in] */ double RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Zoom( 
            /* [retval][out] */ VARIANT *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Zoom( 
            /* [in] */ VARIANT RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintComments( 
            /* [retval][out] */ XlPrintLocation *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintComments( 
            /* [in] */ XlPrintLocation RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrintErrors( 
            /* [retval][out] */ XlPrintErrors *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PrintErrors( 
            /* [in] */ XlPrintErrors RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CenterHeaderPicture( 
            /* [retval][out] */ Graphic **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CenterFooterPicture( 
            /* [retval][out] */ Graphic **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_LeftHeaderPicture( 
            /* [retval][out] */ Graphic **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_LeftFooterPicture( 
            /* [retval][out] */ Graphic **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_RightHeaderPicture( 
            /* [retval][out] */ Graphic **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_RightFooterPicture( 
            /* [retval][out] */ Graphic **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_OddAndEvenPagesHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_OddAndEvenPagesHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_DifferentFirstPageHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_DifferentFirstPageHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ScaleWithDocHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ScaleWithDocHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_AlignMarginsHeaderFooter( 
            /* [retval][out] */ VARIANT_BOOL *RHS);
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_AlignMarginsHeaderFooter( 
            /* [in] */ VARIANT_BOOL RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Pages( 
            /* [retval][out] */ Pages **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_EvenPage( 
            /* [retval][out] */ Page **RHS);
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FirstPage( 
            /* [retval][out] */ Page **RHS);

       CPageSetup()
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
               
       virtual ~CPageSetup()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }

       HRESULT Init( ); 
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* ); 



private:

       long         m_cRef; 
       
       ITypeInfo*   m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;
		
};


#endif // __UNIOFFICE_EXCEL_PAGE_SETUP_H__
