// SelEngVolDll.h: interface for the CSelEngVolDll class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_SELENGVOLDLL_H__196ECA30_613B_49FD_822A_E0FB9DB7AEDA__INCLUDED_)
#define AFX_SELENGVOLDLL_H__196ECA30_613B_49FD_822A_E0FB9DB7AEDA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class AFX_EXT_CLASS CSelEngVolDll  
{
public:
	char* m_SelPrjName;//��������
	char* m_SelVlmName;//�������
	char* m_SelVlmCODE;//������
	long m_SelVlmID;//���ID
	long m_SelHyID;//��ҵ��ID
	char* m_SelPrjID;//���̵�ID
	long m_SelDsgnID;//��ƽ׶ε�ID
	long m_SelSpecID;//רҵ��ID
	BOOL ShowEngVolDlg(LPTSTR cWorkPrjPathName,LPTSTR cAllProjDBPathName,LPTSTR cSortPathName );
	CSelEngVolDll();
	virtual ~CSelEngVolDll();
private:
	CString m_strSelPrjID;//pfg2007.08.20
	CString m_strSelVlmCODE;
	CString m_strSelPrjName;
	CString m_strSelVlmName;
	
	CString m_workprjPathName;//���ݿ�workprj��ȫ·����
	CString m_AllPrjDBPathName;//���ݿ�AllPrjDB��ȫ·����
	CString m_sortPathName;//���ݿ�sort��ȫ·����	

};

#endif // !defined(AFX_SELENGVOLDLL_H__196ECA30_613B_49FD_822A_E0FB9DB7AEDA__INCLUDED_)
