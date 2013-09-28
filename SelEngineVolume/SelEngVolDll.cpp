// SelEngVolDll.cpp: implementation of the CSelEngVolDll class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "SelEngVolDll.h"
#include "DlgSelEngVol.h"
#include "InstallLang.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CSelEngVolDll::CSelEngVolDll()
{
	m_SelPrjID = NULL;//工程的ID
	m_SelPrjName = NULL;//工程名称
	m_SelVlmName = NULL;//卷册名称
	m_SelVlmCODE = NULL;//卷册代号

	m_sortPathName=_T("");
	m_AllPrjDBPathName=_T("");
	m_workprjPathName=_T("");
	m_SelSpecID=0;
	m_SelDsgnID=0;
	m_SelPrjID=_T("");
	m_SelHyID=0;
	m_SelVlmID=0;
}

CSelEngVolDll::~CSelEngVolDll()
{
  
    	
}


//FUNCTION:
//		ShowEngVolDlg()
//Description:
//		选择行业、专业、设计阶段、工程、卷册，并返回它们的ID
//Table Accessed:
//   数据库AllPrjdb.mdb中的：
//   数据库Sort.mdb中的：
//   数据库workprj中的
//Others:
//    Copyright (C),2000-2004,UESOFT Co.,Ltd.

//    Author By 彭范庚、周诗语
//    Date:  2004/07/18


BOOL CSelEngVolDll::ShowEngVolDlg(LPTSTR cWorkPrjPathName,LPTSTR cAllProjDBPathName,LPTSTR cSortPathName )
{
	Language l;
//	l.Create(Language::INSTALL);
	CDlgSelEngVol  selDlg;
 	selDlg.m_sortPathName= CString(cSortPathName);
 	m_sortPathName=CString(cSortPathName);//数据库sort的全路径名
 	selDlg.m_AllPrjDBPathName= CString(cAllProjDBPathName);
 	m_AllPrjDBPathName=CString(cAllProjDBPathName);//数据库AllPrjDB的全路径名
 	selDlg.m_workprjPathName= CString(cWorkPrjPathName);
 	m_workprjPathName= CString(cWorkPrjPathName);//数据库workprj的全路径名
	if(selDlg.DoModal()==IDOK)
	{
 		m_SelSpecID=selDlg.m_SelSpecID;//专业的ID
 		m_SelDsgnID=selDlg.m_SelDsgnID;//设计阶段的ID
 
		m_strSelPrjID=selDlg.m_SelPrjID.GetBuffer(255);//pfg2007.08.20 工程的ID
 		m_SelPrjID=(LPTSTR)m_strSelPrjID.GetBuffer(255);
		m_strSelVlmCODE=selDlg.m_SelVlmCODE;
		m_SelVlmCODE=(LPSTR)m_strSelVlmCODE.GetBuffer(255);
		
		m_SelHyID=selDlg.m_SelHyID;//行业的ID
		m_SelVlmID=selDlg.m_SelVlmID;//卷的ID
		m_strSelPrjName=selDlg.m_SelPrjName;//工程名称
		m_SelPrjName = m_strSelPrjName.GetBuffer(255);
		m_strSelVlmName = selDlg.m_SelVlmName;
		m_SelVlmName=m_strSelVlmName.GetBuffer(255);//卷册名称
		l.Create(Language::UNINSTALL);
		return true;
	}
// 	l.Create(Language::UNINSTALL);

	return false;
}
