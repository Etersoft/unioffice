

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Sat Sep 12 01:39:21 2009
 */
/* Compiler settings for unioffice_excel.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 

#define _MIDL_USE_GUIDDEF_

#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <initguid.h>
#undef INITGUID
#else
#include <initguid.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif // !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_Office,0x2DF8D04C,0x5BFA,0x101B,0xBD,0xE5,0x00,0xAA,0x00,0x44,0xDE,0x52);


MIDL_DEFINE_GUID(IID, IID__Application,0x000208D5,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_Workbooks,0x000208DB,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID__Workbook,0x000208DA,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IWorkbookEvents,0x00024412,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_Sheets,0x000208D7,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_INames,0x000208B8,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IName,0x000208B9,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IBorders,0x00020855,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IBorder,0x00020854,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IInterior,0x00020870,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IOutline,0x000208AB,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IPageSetup,0x000208B4,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID__IFont,0x0002084D,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IRange,0x00020846,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID__Worksheet,0x000208D8,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, IID_IWorksheets,0x000208B1,0x0001,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_AppEvents,0x00024413,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_WorkbookEvents,0x00024412,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_PivotCaches,0x0002441D,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Names,0x000208B8,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Name,0x000208B9,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Borders,0x00020855,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Border,0x00020854,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Interior,0x00020870,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Outline,0x000208AB,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_PageSetup,0x000208B4,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Font,0x0002084D,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Range,0x00020846,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_DocEvents,0x00024411,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(IID, DIID_Worksheets,0x000208B1,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(CLSID, CLSID_Worksheet,0x00020820,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(CLSID, CLSID_Application,0x00024500,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);


MIDL_DEFINE_GUID(CLSID, CLSID_Workbook,0x00020819,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif


