// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__F3490AED_CBCF_4688_9EA9_E04D5071E5FF__INCLUDED_)
#define AFX_STDAFX_H__F3490AED_CBCF_4688_9EA9_E04D5071E5FF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxole.h>         // MFC OLE classes
#include <afxodlgs.h>       // MFC OLE dialog classes
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT


#ifndef _AFX_NO_DB_SUPPORT
#include <afxdb.h>			// MFC ODBC database classes
#endif // _AFX_NO_DB_SUPPORT

#ifndef _AFX_NO_DAO_SUPPORT
#include <afxdao.h>			// MFC DAO database classes
#endif // _AFX_NO_DAO_SUPPORT

#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#import "C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll" no_namespace \
rename("EOF","adoEOF") rename("EditModeEnum","adoEditModeEnum")\
rename("LockTypeEnum","adoLockTypeEnum") rename("FieldAttributeEnum","adoFieldAttributeEnum")\
rename("DataTypeEnum","adoDataTypeEnum") rename("ParameterDirectionEnum","adoParameterDirectionEnum")\
rename("RecordStatusEnum","adoRecordStatusEnum") 

/* chenbl 2013.04.29 没必要这样定义影响自动编译,PDarx中的引用也去掉
#ifdef _ARX2004
   #define CSelEngVolDll  CSelEngVolDll##2004
#endif
   */


//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__F3490AED_CBCF_4688_9EA9_E04D5071E5FF__INCLUDED_)
extern HINSTANCE ghDllInstance;