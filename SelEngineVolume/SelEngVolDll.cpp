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
	m_SelPrjID = NULL;//���̵�ID
	m_SelPrjName = NULL;//��������
	m_SelVlmName = NULL;//�������
	m_SelVlmCODE = NULL;//������

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
//		ѡ����ҵ��רҵ����ƽ׶Ρ����̡���ᣬ���������ǵ�ID
//Table Accessed:
//   ���ݿ�AllPrjdb.mdb�еģ�
//   ���ݿ�Sort.mdb�еģ�
//   ���ݿ�workprj�е�
//Others:
//    Copyright (C),2000-2004,UESOFT Co.,Ltd.

//    Author By ��������ʫ��
//    Date:  2004/07/18


BOOL CSelEngVolDll::ShowEngVolDlg(LPTSTR cWorkPrjPathName,LPTSTR cAllProjDBPathName,LPTSTR cSortPathName )
{
	Language l;
//	l.Create(Language::INSTALL);
	CDlgSelEngVol  selDlg;
 	selDlg.m_sortPathName= CString(cSortPathName);
 	m_sortPathName=CString(cSortPathName);//���ݿ�sort��ȫ·����
 	selDlg.m_AllPrjDBPathName= CString(cAllProjDBPathName);
 	m_AllPrjDBPathName=CString(cAllProjDBPathName);//���ݿ�AllPrjDB��ȫ·����
 	selDlg.m_workprjPathName= CString(cWorkPrjPathName);
 	m_workprjPathName= CString(cWorkPrjPathName);//���ݿ�workprj��ȫ·����
	if(selDlg.DoModal()==IDOK)
	{
 		m_SelSpecID=selDlg.m_SelSpecID;//רҵ��ID
 		m_SelDsgnID=selDlg.m_SelDsgnID;//��ƽ׶ε�ID
 
		m_strSelPrjID=selDlg.m_SelPrjID.GetBuffer(255);//pfg2007.08.20 ���̵�ID
 		m_SelPrjID=(LPTSTR)m_strSelPrjID.GetBuffer(255);
		m_strSelVlmCODE=selDlg.m_SelVlmCODE;
		m_SelVlmCODE=(LPSTR)m_strSelVlmCODE.GetBuffer(255);
		
		m_SelHyID=selDlg.m_SelHyID;//��ҵ��ID
		m_SelVlmID=selDlg.m_SelVlmID;//���ID
		m_strSelPrjName=selDlg.m_SelPrjName;//��������
		m_SelPrjName = m_strSelPrjName.GetBuffer(255);
		m_strSelVlmName = selDlg.m_SelVlmName;
		m_SelVlmName=m_strSelVlmName.GetBuffer(255);//�������
		l.Create(Language::UNINSTALL);
		return true;
	}
// 	l.Create(Language::UNINSTALL);

	return false;
}
