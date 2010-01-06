/*
 * header file - Range
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

#ifndef __UNIOFFICE_EXCEL_RANGE_H__
#define __UNIOFFICE_EXCEL_RANGE_H__

#include "unioffice_excel_private.h"
#include "../OOWrappers/oo_range.h"
#include "../OOWrappers/oo_sheet.h"
#include "worksheet.h"

class CRange : public IRange, public Range
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
               
        // IRange       
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
            /* [retval][out] */ Application	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Creator( 
            /* [retval][out] */ XlCreator *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 
            /* [retval][out] */ IDispatch **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Activate( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_AddIndent( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_AddIndent( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Address( 
            /* [optional][in] */ VARIANT RowAbsolute,
            /* [optional][in] */ VARIANT ColumnAbsolute,
            /* [defaultvalue][optional][in] */ XlReferenceStyle ReferenceStyle,
            /* [optional][in] */ VARIANT External,
            /* [optional][in] */ VARIANT RelativeTo,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ BSTR *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_AddressLocal( 
            /* [optional][in] */ VARIANT RowAbsolute,
            /* [optional][in] */ VARIANT ColumnAbsolute,
            /* [defaultvalue][optional][in] */ XlReferenceStyle ReferenceStyle,
            /* [optional][in] */ VARIANT External,
            /* [optional][in] */ VARIANT RelativeTo,
            /* [retval][out] */ BSTR *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE AdvancedFilter( 
            /* [in] */ XlFilterAction Action,
            /* [optional][in] */ VARIANT CriteriaRange,
            /* [optional][in] */ VARIANT CopyToRange,
            /* [optional][in] */ VARIANT Unique,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ApplyNames( 
            /* [optional][in] */ VARIANT Names,
            /* [optional][in] */ VARIANT IgnoreRelativeAbsolute,
            /* [optional][in] */ VARIANT UseRowColumnNames,
            /* [optional][in] */ VARIANT OmitColumn,
            /* [optional][in] */ VARIANT OmitRow,
            /* [defaultvalue][optional][in] */ XlApplyNamesOrder Order,
            /* [optional][in] */ VARIANT AppendLast,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ApplyOutlineStyles( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Areas( 
            /* [retval][out] */ Areas **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE AutoComplete( 
            /* [in] */ BSTR String,
            /* [retval][out] */ BSTR *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE AutoFill( 
            /* [in] */ Range	*Destination,
            /* [defaultvalue][optional][in] */ XlAutoFillType Type,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE AutoFilter( 
            /* [optional][in] */ VARIANT Field,
            /* [optional][in] */ VARIANT Criteria1,
            /* [defaultvalue][optional][in] */ XlAutoFilterOperator Operator,
            /* [optional][in] */ VARIANT Criteria2,
            /* [optional][in] */ VARIANT VisibleDropDown,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE AutoFit( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE AutoFormat( 
            /* [defaultvalue][optional][in] */ XlRangeAutoFormat Format,
            /* [optional][in] */ VARIANT Number,
            /* [optional][in] */ VARIANT Font,
            /* [optional][in] */ VARIANT Alignment,
            /* [optional][in] */ VARIANT Border,
            /* [optional][in] */ VARIANT Pattern,
            /* [optional][in] */ VARIANT Width,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE AutoOutline( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE BorderAround( 
            /* [optional][in] */ VARIANT LineStyle,
            /* [defaultvalue][optional][in] */ XlBorderWeight Weight,
            /* [defaultvalue][optional][in] */ XlColorIndex ColorIndex,
            /* [optional][in] */ VARIANT Color,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Borders( 
            /* [retval][out] */ Borders	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Calculate( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Cells( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Characters( 
            /* [optional][in] */ VARIANT Start,
            /* [optional][in] */ VARIANT Length,
            /* [retval][out] */ Characters **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CheckSpelling( 
            /* [optional][in] */ VARIANT CustomDictionary,
            /* [optional][in] */ VARIANT IgnoreUppercase,
            /* [optional][in] */ VARIANT AlwaysSuggest,
            /* [optional][in] */ VARIANT SpellLang,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Clear( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ClearContents( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ClearFormats( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ClearNotes( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ClearOutline( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Column( 
            /* [retval][out] */ long *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ColumnDifferences( 
            /* [in] */ VARIANT Comparison,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Columns( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ColumnWidth( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ColumnWidth( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Consolidate( 
            /* [optional][in] */ VARIANT Sources,
            /* [optional][in] */ VARIANT Function,
            /* [optional][in] */ VARIANT TopRow,
            /* [optional][in] */ VARIANT LeftColumn,
            /* [optional][in] */ VARIANT CreateLinks,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Copy( 
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CopyFromRecordset( 
            /* [in] */ IUnknown *Data,
            /* [optional][in] */ VARIANT MaxRows,
            /* [optional][in] */ VARIANT MaxColumns,
            /* [retval][out] */ long *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CopyPicture( 
            /* [defaultvalue][optional][in] */ XlPictureAppearance Appearance,
            /* [defaultvalue][optional][in] */ XlCopyPictureFormat Format,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CreateNames( 
            /* [optional][in] */ VARIANT Top,
            /* [optional][in] */ VARIANT Left,
            /* [optional][in] */ VARIANT Bottom,
            /* [optional][in] */ VARIANT Right,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE CreatePublisher( 
            /* [optional][in] */ VARIANT Edition,
            /* [defaultvalue][optional][in] */ XlPictureAppearance Appearance,
            /* [optional][in] */ VARIANT ContainsPICT,
            /* [optional][in] */ VARIANT ContainsBIFF,
            /* [optional][in] */ VARIANT ContainsRTF,
            /* [optional][in] */ VARIANT ContainsVALU,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CurrentArray( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CurrentRegion( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Cut( 
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE DataSeries( 
            /* [optional][in] */ VARIANT Rowcol,
            /* [defaultvalue][optional][in] */ XlDataSeriesType Type,
            /* [defaultvalue][optional][in] */ XlDataSeriesDate Date,
            /* [optional][in] */ VARIANT Step,
            /* [optional][in] */ VARIANT Stop,
            /* [optional][in] */ VARIANT Trend,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get__Default( 
            /* [optional][in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put__Default( 
            /* [optional][in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Delete( 
            /* [optional][in] */ VARIANT Shift,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Dependents( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE DialogBox( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_DirectDependents( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_DirectPrecedents( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE EditionOptions( 
            /* [in] */ XlEditionType Type,
            /* [in] */ XlEditionOptionsOption Option,
            /* [optional][in] */ VARIANT Name,
            /* [optional][in] */ VARIANT Reference,
            /* [defaultvalue][optional][in] */ XlPictureAppearance Appearance,
            /* [defaultvalue][optional][in] */ XlPictureAppearance ChartSize,
            /* [optional][in] */ VARIANT Format,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_End( 
            /* [in] */ XlDirection Direction,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_EntireColumn( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_EntireRow( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE FillDown( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE FillLeft( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE FillRight( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE FillUp( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Find( 
            /* [in] */ VARIANT What,
            /* [optional][in] */ VARIANT After,
            /* [optional][in] */ VARIANT LookIn,
            /* [optional][in] */ VARIANT LookAt,
            /* [optional][in] */ VARIANT SearchOrder,
            /* [defaultvalue][optional][in] */ XlSearchDirection SearchDirection,
            /* [optional][in] */ VARIANT MatchCase,
            /* [optional][in] */ VARIANT MatchByte,
            /* [optional][in] */ VARIANT SearchFormat,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE FindNext( 
            /* [optional][in] */ VARIANT After,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE FindPrevious( 
            /* [optional][in] */ VARIANT After,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Font( 
            /* [retval][out] */ /* external definition not present */ Font **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Formula( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Formula( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FormulaArray( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FormulaArray( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][hidden][propget] */ HRESULT STDMETHODCALLTYPE get_FormulaLabel( 
            /* [retval][out] */ XlFormulaLabel *RHS) ;
        
        virtual /* [helpcontext][hidden][propput] */ HRESULT STDMETHODCALLTYPE put_FormulaLabel( 
            /* [in] */ XlFormulaLabel RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FormulaHidden( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FormulaHidden( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FormulaLocal( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FormulaLocal( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FormulaR1C1( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FormulaR1C1( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FormulaR1C1Local( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_FormulaR1C1Local( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE FunctionWizard( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE GoalSeek( 
            /* [in] */ VARIANT Goal,
            /* [in] */ Range	*ChangingCell,
            /* [retval][out] */ VARIANT_BOOL *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Group( 
            /* [optional][in] */ VARIANT Start,
            /* [optional][in] */ VARIANT End,
            /* [optional][in] */ VARIANT By,
            /* [optional][in] */ VARIANT Periods,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_HasArray( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_HasFormula( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Height( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Hidden( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Hidden( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_HorizontalAlignment( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_HorizontalAlignment( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_IndentLevel( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_IndentLevel( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE InsertIndent( 
            /* [in] */ long InsertAmount) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Insert( 
            /* [optional][in] */ VARIANT Shift,
            /* [optional][in] */ VARIANT CopyOrigin,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Interior( 
            /* [retval][out] */ Interior	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Item( 
            /* [in] */ VARIANT RowIndex,
            /* [optional][in] */ VARIANT ColumnIndex,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Justify( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Left( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ListHeaderRows( 
            /* [retval][out] */ long *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ListNames( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_LocationInTable( 
            /* [retval][out] */ XlLocationInTable *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Locked( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Locked( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Merge( 
            /* [optional][in] */ VARIANT Across) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE UnMerge( void) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_MergeArea( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_MergeCells( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_MergeCells( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Name( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE NavigateArrow( 
            /* [optional][in] */ VARIANT TowardPrecedent,
            /* [optional][in] */ VARIANT ArrowNumber,
            /* [optional][in] */ VARIANT LinkNumber,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Next( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE NoteText( 
            /* [optional][in] */ VARIANT Text,
            /* [optional][in] */ VARIANT Start,
            /* [optional][in] */ VARIANT Length,
            /* [retval][out] */ BSTR *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_NumberFormat( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_NumberFormat( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_NumberFormatLocal( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_NumberFormatLocal( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Offset( 
            /* [optional][in] */ VARIANT RowOffset,
            /* [optional][in] */ VARIANT ColumnOffset,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Orientation( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Orientation( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_OutlineLevel( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_OutlineLevel( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PageBreak( 
            /* [retval][out] */ long *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_PageBreak( 
            /* [in] */ long RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Parse( 
            /* [optional][in] */ VARIANT ParseLine,
            /* [optional][in] */ VARIANT Destination,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE _PasteSpecial( 
            /* [defaultvalue][optional][in] */ XlPasteType Paste,
            /* [defaultvalue][optional][in] */ XlPasteSpecialOperation Operation,
            /* [optional][in] */ VARIANT SkipBlanks,
            /* [optional][in] */ VARIANT Transpose,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PivotField( 
            /* [retval][out] */ PivotField **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PivotItem( 
            /* [retval][out] */ PivotItem **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PivotTable( 
            /* [retval][out] */ PivotTable **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Precedents( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PrefixCharacter( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Previous( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE __PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE PrintPreview( 
            /* [optional][in] */ VARIANT EnableChanges,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_QueryTable( 
            /* [retval][out] */ QueryTable **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Range( 
            /* [in] */ VARIANT Cell1,
            /* [optional][in] */ VARIANT Cell2,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE RemoveSubtotal( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Replace( 
            /* [in] */ VARIANT What,
            /* [in] */ VARIANT Replacement,
            /* [optional][in] */ VARIANT LookAt,
            /* [optional][in] */ VARIANT SearchOrder,
            /* [optional][in] */ VARIANT MatchCase,
            /* [optional][in] */ VARIANT MatchByte,
            /* [optional][in] */ VARIANT SearchFormat,
            /* [optional][in] */ VARIANT ReplaceFormat,
            /* [retval][out] */ VARIANT_BOOL *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Resize( 
            /* [optional][in] */ VARIANT RowSize,
            /* [optional][in] */ VARIANT ColumnSize,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Row( 
            /* [retval][out] */ long *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE RowDifferences( 
            /* [in] */ VARIANT Comparison,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_RowHeight( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_RowHeight( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Rows( 
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Run( 
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
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Select( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Show( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ShowDependents( 
            /* [optional][in] */ VARIANT Remove,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ShowDetail( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ShowDetail( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ShowErrors( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ShowPrecedents( 
            /* [optional][in] */ VARIANT Remove,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ShrinkToFit( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ShrinkToFit( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Sort( 
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
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE SortSpecial( 
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
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_SoundNote( 
            /* [retval][out] */ SoundNote **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE SpecialCells( 
            /* [in] */ XlCellType Type,
            /* [optional][in] */ VARIANT Value,
            /* [retval][out] */ Range	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Style( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Style( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE SubscribeTo( 
            /* [in] */ BSTR Edition,
            /* [defaultvalue][optional][in] */ XlSubscribeToFormat Format,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Subtotal( 
            /* [in] */ long GroupBy,
            /* [in] */ XlConsolidationFunction Function,
            /* [in] */ VARIANT TotalList,
            /* [optional][in] */ VARIANT Replace,
            /* [optional][in] */ VARIANT PageBreaks,
            /* [defaultvalue][optional][in] */ XlSummaryRow SummaryBelowData,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Summary( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Table( 
            /* [optional][in] */ VARIANT RowInput,
            /* [optional][in] */ VARIANT ColumnInput,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Text( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE TextToColumns( 
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
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Top( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Ungroup( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_UseStandardHeight( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_UseStandardHeight( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_UseStandardWidth( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_UseStandardWidth( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Validation( 
            /* [retval][out] */ Validation **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Value( 
            /* [optional][in] */ VARIANT RangeValueDataType,
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Value( 
            /* [optional][in] */ VARIANT RangeValueDataType,
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Value2( 
            /* [lcid][in] */ long lcid,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_Value2( 
            /* [lcid][in] */ long lcid,
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_VerticalAlignment( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_VerticalAlignment( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Width( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Worksheet( 
            /* [retval][out] */ Worksheet	**RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_WrapText( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_WrapText( 
            /* [in] */ VARIANT RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE AddComment( 
            /* [optional][in] */ VARIANT Text,
            /* [retval][out] */ Comment **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Comment( 
            /* [retval][out] */ Comment **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ClearComments( void) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Phonetic( 
            /* [retval][out] */ Phonetic **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_FormatConditions( 
            /* [retval][out] */ FormatConditions **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ReadingOrder( 
            /* [retval][out] */ long *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ReadingOrder( 
            /* [in] */ long RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Hyperlinks( 
            /* [retval][out] */ HyperLinks **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Phonetics( 
            /* [retval][out] */ Phonetics **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE SetPhonetic( void) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ID( 
            /* [retval][out] */ BSTR *RHS) ;
        
        virtual /* [helpcontext][propput] */ HRESULT STDMETHODCALLTYPE put_ID( 
            /* [in] */ BSTR RHS) ;
        
        virtual /* [helpcontext][hidden] */ HRESULT STDMETHODCALLTYPE _PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_PivotCell( 
            /* [retval][out] */ PivotCell **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Dirty( void) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_Errors( 
            /* [retval][out] */ Errors **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_SmartTags( 
            /* [retval][out] */ SmartTags **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE Speak( 
            /* [optional][in] */ VARIANT SpeakDirection,
            /* [optional][in] */ VARIANT SpeakFormulas) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE PasteSpecial( 
            /* [defaultvalue][optional][in] */ XlPasteType Paste,
            /* [defaultvalue][optional][in] */ XlPasteSpecialOperation Operation,
            /* [optional][in] */ VARIANT SkipBlanks,
            /* [optional][in] */ VARIANT Transpose,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_AllowEdit( 
            /* [retval][out] */ VARIANT_BOOL *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ListObject( 
            /* [retval][out] */ ListObject **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_XPath( 
            /* [retval][out] */ XPath **RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_ServerActions( 
            /* [retval][out] */ Actions **RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE RemoveDuplicates( 
            /* [optional][in] */ VARIANT Columns,
            /* [defaultvalue][optional][in] */ XlYesNoGuess Header = 2) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE PrintOut( 
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT Copies,
            /* [optional][in] */ VARIANT Preview,
            /* [optional][in] */ VARIANT ActivePrinter,
            /* [optional][in] */ VARIANT PrintToFile,
            /* [optional][in] */ VARIANT Collate,
            /* [optional][in] */ VARIANT PrToFileName,
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_MDX( 
            /* [retval][out] */ BSTR *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE ExportAsFixedFormat( 
            /* [in] */ XlFixedFormatType Type,
            /* [optional][in] */ VARIANT Filename,
            /* [optional][in] */ VARIANT Quality,
            /* [optional][in] */ VARIANT IncludeDocProperties,
            /* [optional][in] */ VARIANT IgnorePrintAreas,
            /* [optional][in] */ VARIANT From,
            /* [optional][in] */ VARIANT To,
            /* [optional][in] */ VARIANT OpenAfterPublish,
            /* [optional][in] */ VARIANT FixedFormatExtClassPtr) ;
        
        virtual /* [helpcontext][propget] */ HRESULT STDMETHODCALLTYPE get_CountLarge( 
            /* [retval][out] */ VARIANT *RHS) ;
        
        virtual /* [helpcontext] */ HRESULT STDMETHODCALLTYPE CalculateRowMajorOrder( 
            /* [retval][out] */ VARIANT *RHS) ;
               
       CRange()
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

       virtual ~CRange()
       {
            InterlockedDecrement(&g_cComponents);    
              
            m_p_application = NULL;
            
            if ( m_p_parent != NULL )
   			{
   	  		    (static_cast<Worksheet*>( m_p_parent ))->Release( );	
	       	    m_p_parent = NULL;	  	
   		   	}
            
            m_p_parent = NULL;  
                            
            DELETE_OBJECT;             
       }

       HRESULT Init( ); 
       
       HRESULT Put_Application( void* );
       HRESULT Put_Parent( void* ); 
	   
	   HRESULT InitWrapper( OORange );	   
	   	   
private:
	
	   OOSheet      getParentOOSheet();
	    	
       long         m_cRef; 
       
       ITypeInfo*   m_pITypeInfo;               
       
       void*        m_p_application;
       void*        m_p_parent;	
       
       OORange      m_oo_range;
	   
};

#endif // __UNIOFFICE_EXCEL_RANGE_H__
	   			 
