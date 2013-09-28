#if !defined(AFX_DBCOMBOBOX_H__591B0B3E_0EEB_48C2_AC34_C4D5D986CB72__INCLUDED_)
#define AFX_DBCOMBOBOX_H__591B0B3E_0EEB_48C2_AC34_C4D5D986CB72__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DBComboBox.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDBComboBox window

class __declspec(dllexport)CDBComboBox : public CComboBox
{
// Construction
public:
	CDBComboBox();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDBComboBox)
	//}}AFX_VIRTUAL

// Implementation
public:
	void RefLst();
	void LoadList();
	CString m_ListField;
	CString m_BoundField;
	CString m_Field;
	_RecordsetPtr m_RowRs;
	_RecordsetPtr m_Rs;
	void OnSelchange();
	virtual ~CDBComboBox();

	// Generated message map functions
protected:
	//{{AFX_MSG(CDBComboBox)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DBCOMBOBOX_H__591B0B3E_0EEB_48C2_AC34_C4D5D986CB72__INCLUDED_)
