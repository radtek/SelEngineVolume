#if !defined(AFX_UEDATAGRID_H__2AE55434_084E_4780_B854_CC6C0CF97E14__INCLUDED_)
#define AFX_UEDATAGRID_H__2AE55434_084E_4780_B854_CC6C0CF97E14__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// UeDataGrid.h : header file
//
#include "DataGrid.h"
#include <Afxtempl.h>
/////////////////////////////////////////////////////////////////////////////
// UeDataGrid window

class __declspec(dllexport) UeDataGrid : public CDataGrid
{
	class HookInfo
	{
	public:
		class Hook
		{
		public:
			Hook(){
				m_dwThreadID = 0;
				m_hHook = NULL;
			}
			virtual ~Hook(){} 

		BOOL HasWindow(HWND hWnd)
		{
			for(int i = 0; i < m_theWindows.GetSize(); i++)
			{
				if(hWnd == m_theWindows.GetAt(i)->GetSafeHwnd())
					return TRUE;
			}
			return FALSE;
		}

		public:
			CArray<CWnd*, CWnd*> m_theWindows;
			DWORD m_dwThreadID;
			HHOOK m_hHook;
		};
	public:
		static void DeleteHook(DWORD dwThreadID, CWnd* pWnd);
		static BOOL AddHook(DWORD dwThreadID, CWnd* pWnd);

		static CArray<Hook* , Hook*> m_theHooks;

		static LRESULT CALLBACK MouseProc( int nCode, WPARAM wParam, LPARAM lParam);
		static Hook* GetHook(DWORD dwThreadID);
	public:
		HookInfo();
		virtual ~HookInfo();

	};
// Construction
public:
	UeDataGrid();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(UeDataGrid)
	//}}AFX_VIRTUAL

private:

private:
	UINT m_nPrevVerticalPos;
	UINT m_nPrevHorizonPos;

	HWND m_hVerticalBar;
	HWND m_hHorizonBar;
// Implementation
public:
	CString GetSelectData(INT nColumn);
	virtual ~UeDataGrid();

	// Generated message map functions
protected:
	//{{AFX_MSG(UeDataGrid)
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnHScrollClipboard(CWnd* pClipAppWnd, UINT nSBCode, UINT nPos);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_UEDATAGRID_H__2AE55434_084E_4780_B854_CC6C0CF97E14__INCLUDED_)
