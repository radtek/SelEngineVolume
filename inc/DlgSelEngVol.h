//{{AFX_INCLUDES()
#include "Uedatagrid.h"
//}}AFX_INCLUDES
#if !defined(AFX_DLGSELENGVOL_H__3AA03FB2_CC37_4567_B46B_6B8CA78F1948__INCLUDED_)
#define AFX_DLGSELENGVOL_H__3AA03FB2_CC37_4567_B46B_6B8CA78F1948__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DlgSelEngVol.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDlgSelEngVol dialog
#include "resource.h"
#include "DBComboBox.h"

//#include "..\ZSY\PFG COPY\DBComboBox.h"	// Added by ClassView
class __declspec(dllexport) CDlgSelEngVol : public CDialog
{
// Construction 
public:
	bool tdfExists(_ConnectionPtr pConn,CString tbn);
	void ReleaseMemoryEngVol();
	BOOL m_bEngDelete;
	BOOL m_bVlmDelete;
	long m_SelVlmID;
	CString m_workprjPathName;
	CString m_AllPrjDBPathName;
	CString m_sortPathName;
//
	CString m_SelPrjName;//工程名称  zsy add
	CString m_SelVlmName;//卷册名称
	CString m_SelVlmCODE;//卷册代号
//
	CString m_currEngID;
	short   m_lastEngRow;
	CString m_currVlmID;
	short   m_lastVlmRow;

	void changeDsgnSpeGoRY();
	_ConnectionPtr condbPRJ;
	BOOL UnInitHook();
	BOOL InitHook();
	HHOOK m_hHook;
	static LRESULT CALLBACK CBTProc(int nCode,WPARAM wParam, LPARAM lParam);
	void SaveDBGridColumnCaptionAndWidth(CDataGrid& MyDBGrid,long ColIndex,CString tbn);
	bool m_bIsVlmNew;
	bool m_bIsNew;
	_RecordsetPtr m_DataCategory;
	_RecordsetPtr m_DataSpc;
	_RecordsetPtr m_DataDsgn;
	_RecordsetPtr m_DataCurrWork;

	long GetMaxVolumeID();
	CString inline ltos(long v);
	_RecordsetPtr m_DataVlm;
	int vtoi(_variant_t& v);
	_ConnectionPtr conSORT;
	void SetDBGridColumnCaptionAndWidth(CDataGrid& MyDBGrid, CString tbn);
	CString vtos(_variant_t &v);
	CString Trim(LPCTSTR pcs);
	long m_SelSpecID;
	long m_SelDsgnID;
	CString m_SelPrjID;
	long m_SelHyID;
	void SetWindowCenter( HWND hWnd );
	void InitPrjDb();
	void InitDBVlm();
	void InitENG();
	BOOL ConnectDatabase();
	_RecordsetPtr m_DataEng;
	_ConnectionPtr conPRJDB;
	CDlgSelEngVol(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDlgSelEngVol)
	enum { IDD = IDD_SELPDSV };
	CButton	m_okSelEng;
	CButton	m_cancalSelEng;
	UeDataGrid	m_DBGeng;
	UeDataGrid	m_DBGvlm;
	//CString	m_csDBCategory;
	//CString	m_csDBDsgn;
	//CString	m_csDBSpec;
	CDBComboBox m_DBComboSpec;
	CDBComboBox m_DBComboDSGN;
	CDBComboBox m_DBComboCategory;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDlgSelEngVol)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	BOOL CreateTableSaveCurrent();
	bool m_stateEditVlm;
	bool m_stateEditEng;

	// Generated message map functions
	//{{AFX_MSG(CDlgSelEngVol)
	virtual BOOL OnInitDialog();
	afx_msg void OnRowColChangeDbgeng(VARIANT FAR* LastRow, short LastCol);
	afx_msg void OnRowColChangeDbgvim(VARIANT FAR* LastRow, short LastCol);
	afx_msg void OnOnAddNewDbgeng();
	afx_msg void OnOnAddNewDbgvim();
	afx_msg void OnAfterUpdateDbgeng();
	afx_msg void OnAfterUpdateDbgvim();
	afx_msg void OnColResizeDbgeng(short ColIndex, short FAR* Cancel);
	afx_msg void OnColResizeDbgvim(short ColIndex, short FAR* Cancel);
	afx_msg void OnBeforeUpdateDbgvim(short FAR* Cancel);
	afx_msg void OnBeforeUpdateDbgeng(short FAR* Cancel);
	afx_msg void OnBeforeDeleteDbgeng(short FAR* Cancel);
	afx_msg void OnBeforeDeleteDbgvim(short FAR* Cancel);
	afx_msg void OnSelchangeComboHY();
	afx_msg void OnSelchangeComboDsgn();
	afx_msg void OnSelchangeComboSpec();
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnColEditDbgeng(short ColIndex);
	afx_msg void OnColEditDbgvim(short ColIndex);
	afx_msg void OnKeyDownDbgeng(short FAR* KeyCode, short Shift);
	afx_msg void OnKeyUpDbgeng(short FAR* KeyCode, short Shift);
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGSELENGVOL_H__3AA03FB2_CC37_4567_B46B_6B8CA78F1948__INCLUDED_)
