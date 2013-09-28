// UeDataGrid.cpp : implementation file
//

#include "stdafx.h"
//#include "phs.h"
#include "UeDataGrid.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// UeDataGrid

UeDataGrid::UeDataGrid()
{
	m_nPrevVerticalPos = 0;
	m_nPrevHorizonPos = 0;

	HookInfo::AddHook(::GetCurrentThreadId(), this);
}

UeDataGrid::~UeDataGrid()
{
	HookInfo::DeleteHook(::GetCurrentThreadId(), this);
}


BEGIN_MESSAGE_MAP(UeDataGrid, CDataGrid)
	//{{AFX_MSG_MAP(UeDataGrid)
	ON_WM_VSCROLL()
	ON_WM_HSCROLLCLIPBOARD()
	ON_WM_HSCROLL()
	ON_WM_MOUSEWHEEL()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// UeDataGrid message handlers

void UeDataGrid::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	if(this->m_hVerticalBar == NULL)
		this->m_hVerticalBar = pScrollBar->GetSafeHwnd();

	if(nSBCode == SB_THUMBPOSITION || nSBCode == SB_THUMBTRACK)
	{
		if(this->m_nPrevVerticalPos != nPos)
		{
			this->Scroll(0, nPos - this->m_nPrevVerticalPos);
			this->m_nPrevVerticalPos = nPos;
			return;
		}
	}
	else
	{
		CDataGrid::OnVScroll(nSBCode, nPos, pScrollBar);
	}
}

void UeDataGrid::OnHScrollClipboard(CWnd* pClipAppWnd, UINT nSBCode, UINT nPos) 
{
	// TODO: Add your message handler code here and/or call default
	
	UeDataGrid::OnHScrollClipboard(pClipAppWnd, nSBCode, nPos);
}

CString UeDataGrid::GetSelectData(INT nColumn)
{
//	_RecordsetPtr rs = this->GetDataSource();
//	if(rs == NULL)
		return CString();
//	return vtos(rs->GetCollect((long)nColumn));
}

void UeDataGrid::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{	
	if(this->m_hHorizonBar == NULL)
		this->m_hHorizonBar = pScrollBar->GetSafeHwnd();

	if(nSBCode == SB_THUMBPOSITION || nSBCode == SB_THUMBTRACK)
	{
		if(this->m_nPrevHorizonPos != nPos)
		{
			this->Scroll(nPos - this->m_nPrevHorizonPos, 0);
			this->m_nPrevHorizonPos = nPos;
			return;
		}
	}
	else
	{
		CDataGrid::OnHScroll(nSBCode, nPos, pScrollBar);
	}
}

BOOL UeDataGrid::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	if(zDelta > 0)
	{
		this->Scroll(0, -3);
	}
	else if(zDelta < 0)
	{
		this->Scroll(0, 3);
	}
	
	return CWnd::OnMouseWheel(nFlags, zDelta, pt);
}

UeDataGrid::HookInfo::HookInfo()
{

}

UeDataGrid::HookInfo::~HookInfo()
{
	Hook* pHook = NULL;
	for(int i = 0; i < m_theHooks.GetSize(); i++)
	{
		pHook = m_theHooks.GetAt(i);
		::UnhookWindowsHookEx(pHook->m_hHook);
		delete pHook;
	}
	m_theHooks.RemoveAll();
}

CArray<UeDataGrid::HookInfo::Hook*, UeDataGrid::HookInfo::Hook*> UeDataGrid::HookInfo::m_theHooks;

BOOL UeDataGrid::HookInfo::AddHook(DWORD dwThreadID, CWnd* pWnd)
{
	if(pWnd == NULL)
		return FALSE;

	Hook* pHook = NULL;

	for(int i = 0; i < m_theHooks.GetSize(); i++)
	{
		pHook = m_theHooks.GetAt(i);
		if(pHook->m_dwThreadID == dwThreadID)
		{
			pHook->m_theWindows.Add(pWnd);
			return TRUE;
		}
	}
	
	HHOOK hHook = ::SetWindowsHookEx(WH_GETMESSAGE, UeDataGrid::HookInfo::MouseProc, NULL, dwThreadID);
	if(hHook == NULL)
		return FALSE;
	pHook = new Hook;
	if(pHook == NULL)
	{
		::UnhookWindowsHookEx(hHook);
		return FALSE;
	}
	pHook->m_dwThreadID = dwThreadID;
	pHook->m_hHook = hHook;
	pHook->m_theWindows.Add(pWnd);
	m_theHooks.Add(pHook);

	return TRUE;
}

void UeDataGrid::HookInfo::DeleteHook(DWORD dwThreadID, CWnd* pWnd)
{
	Hook* pHook = NULL;
	
	for(int i = 0; i < m_theHooks.GetSize(); i++)
	{
		pHook = m_theHooks.GetAt(i);
		if(pHook->m_dwThreadID == dwThreadID)
		{
			for(int j = 0; j < pHook->m_theWindows.GetSize(); j++)
			{
				if(pWnd == pHook->m_theWindows.GetAt(j))
				{
					pHook->m_theWindows.RemoveAt(j);
				}
			}
			if(pHook->m_theWindows.GetSize() == 0)
			{
				::UnhookWindowsHookEx(pHook->m_hHook);
				delete pHook;
				m_theHooks.RemoveAt(i);
			}
			return ;
		}
	}	
}

UeDataGrid::HookInfo::Hook* UeDataGrid::HookInfo::GetHook(DWORD dwThreadID)
{
	Hook* pHook = NULL;
	for(int i = 0; i < m_theHooks.GetSize(); i++)
	{
		pHook = m_theHooks.GetAt(i);
		if(pHook->m_dwThreadID == dwThreadID)
		{
			return pHook;
		}
	}
	return NULL;
}

LRESULT CALLBACK UeDataGrid::HookInfo::MouseProc( int nCode, WPARAM wParam, LPARAM lParam)
{
	DWORD dwThreadID = ::GetCurrentThreadId();
	
	Hook* pHook = GetHook(dwThreadID);
	if(pHook == NULL)
		return 0;

	if(nCode < 0)
		return ::CallNextHookEx(pHook->m_hHook, nCode, wParam, lParam);

	MSG* pMsg = (MSG*)lParam;

	if(pMsg->message == WM_MOUSEWHEEL && pHook->HasWindow(::GetParent(pMsg->hwnd)))
	{
		::SendMessage(::GetParent(pMsg->hwnd), pMsg->message, pMsg->wParam, pMsg->lParam);
	}

	return ::CallNextHookEx(pHook->m_hHook, nCode, wParam, lParam);
}