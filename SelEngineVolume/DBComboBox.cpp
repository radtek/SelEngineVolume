// DBComboBox.cpp : implementation file
//

#include "stdafx.h"
//#include "	\ add additional includes here"
#include "DBComboBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDBComboBox

CDBComboBox::CDBComboBox()
{
}

CDBComboBox::~CDBComboBox()
{
}


BEGIN_MESSAGE_MAP(CDBComboBox, CComboBox)
	//{{AFX_MSG_MAP(CDBComboBox)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDBComboBox message handlers

void CDBComboBox::OnSelchange()
{

	if(m_RowRs->State!=adStateOpen||m_RowRs->GetRecordCount()==0)
		return;
//7/24	if(m_Rs->State!=adStateOpen||m_Rs->GetRecordCount()==0)
//7/24		return;
	if(m_BoundField==""||m_Field=="")
		return;
	int icursel;
	if((icursel=this->GetCurSel())<0)
		return;
	m_RowRs->MoveFirst();
	for(int i=0;!m_RowRs->adoEOF;i++,m_RowRs->MoveNext())
	{
		if(i==icursel)
		{
		/*	_variant_t val;
			val=m_RowRs->GetCollect(_variant_t(m_BoundField));
			m_Rs->PutCollect(_T(_variant_t(m_Field)),_variant_t(val));
			val=m_RowRs->GetCollect(_variant_t(m_ListField));
			m_Rs->PutCollect(_T(_variant_t(m_ListField)),_variant_t(val));
			m_Rs->Update();*/
			break;

		}
	}



}

void CDBComboBox::LoadList()
{
	_variant_t temp;
	try
	{
		if(m_RowRs->GetRecordCount()==0||m_ListField==_T(""))
		return;
		this->ResetContent();//清空ComboBox控件中的内容
		m_RowRs->MoveFirst();
		CString str;
		while(!m_RowRs->adoEOF)
		{
			temp=m_RowRs->GetCollect(_variant_t(m_ListField));
			if(temp.vt==VT_EMPTY)
			{//1
				str="";
			}//1
			else
			{//1
				str=temp.bstrVal;
				this->AddString(str);

			}//1
			m_RowRs->MoveNext();
		}

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}

}

void CDBComboBox::RefLst()
{
//	m_fis=true;
	try
	{
		
	//	if(m_RowRs==NULL || !m_RowRs->IsOpen())
//			return;
	//	if(m_Rs==NULL || !m_Rs->IsOpen())
	//		return;
		if(m_BoundField==_T("") || m_Field==_T(""))
			return;
		int i=0;
		if(m_RowRs->GetRecordCount()==0)
		{
			this->SetCurSel(-1);
			return;
		}
		if(m_RowRs->BOF)
		{
			m_RowRs->MoveFirst();
		}
		else if(m_RowRs->adoEOF)
		{
			m_RowRs->MoveLast();
		}
		_variant_t key;
		key=m_RowRs->GetCollect(_variant_t(m_ListField));
		CString str;
		if(key.vt==VT_EMPTY)
		{str="";}
		else
		{str=key.bstrVal;}
		i=this->FindStringExact(-1,str);
		this->SetCurSel(i);
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}
//	m_fis=false;

}
