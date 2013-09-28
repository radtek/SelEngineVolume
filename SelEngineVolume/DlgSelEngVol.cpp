// DlgSelEngVol.cpp : implementation file
//

#include "stdafx.h"
//#include "	\ add additional includes here"
#include "DlgSelEngVol.h"
#include "Columns.h"
#include "Column.h"
#include "DBComboBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDlgSelEngVol dialog


CDlgSelEngVol::CDlgSelEngVol(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgSelEngVol::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDlgSelEngVol)
//	m_csDBCategory = -1;
//	m_csDBDsgn = -1;
//	m_csDBSpec = -1;
	//}}AFX_DATA_INIT
	m_bIsNew  = false;
	m_bIsVlmNew = false;
	m_bEngDelete = false;
	m_bVlmDelete = false;
	m_stateEditEng=false;//7/25
	m_stateEditVlm=false;//7/25

}


void CDlgSelEngVol::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDlgSelEngVol)
	DDX_Control(pDX, IDOK, m_okSelEng);
	DDX_Control(pDX, IDCANCEL, m_cancalSelEng);
	DDX_Control(pDX, IDC_DBGENG, m_DBGeng);
	DDX_Control(pDX, IDC_DBGVIM, m_DBGvlm);
//	DDX_CBStringExact(pDX, IDC_COMBO_CATEGORY, m_csDBCategory);
//	DDX_CBStringExact(pDX, IDC_COMBO_DSGN, m_csDBDsgn);
//	DDX_CBStringExact(pDX, IDC_COMBO_SPEC, m_csDBSpec);
	DDX_Control(pDX, IDC_COMBO_SPEC, m_DBComboSpec);
	DDX_Control(pDX, IDC_COMBO_DSGN, m_DBComboDSGN);
	DDX_Control(pDX, IDC_COMBO_CATEGORY, m_DBComboCategory);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDlgSelEngVol, CDialog)
	//{{AFX_MSG_MAP(CDlgSelEngVol)
	ON_CBN_SELCHANGE(IDC_COMBO_CATEGORY, OnSelchangeComboHY)
	ON_CBN_SELCHANGE(IDC_COMBO_DSGN, OnSelchangeComboDsgn)
	ON_CBN_SELCHANGE(IDC_COMBO_SPEC, OnSelchangeComboSpec)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDlgSelEngVol message handlers

BOOL CDlgSelEngVol::ConnectDatabase()
{
	conPRJDB.CreateInstance(__uuidof(Connection));
//	CString conPRJDBSql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\prjdb4.0\\AllPrjDB.mdb;PerSist Security Info=False";
// 	CString conPRJDBSql = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source='"+m_AllPrjDBPathName+"';PerSist Security Info=False";
	CString conPRJDBSql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='"+m_AllPrjDBPathName+"';PerSist Security Info=False";

	conSORT.CreateInstance(__uuidof(Connection));
//	CString conSORTSql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\prjdb4.0\\sort.mdb;PerSist Security Info=False";
// 	CString conSORTSql = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source='"+m_sortPathName+"';PerSist Security Info=False";
	CString conSORTSql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='"+m_sortPathName+"';PerSist Security Info=False";

	condbPRJ.CreateInstance(__uuidof(Connection));
//	CString condbprjSql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\prj\\workprj.mdb;PerSist Security Info=False";
// 	CString condbprjSql = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source='"+m_workprjPathName+"';PerSist Security Info=False";
	CString condbprjSql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='"+m_workprjPathName+"';PerSist Security Info=False";

	try
	{
	   conPRJDB->Open(_bstr_t(conPRJDBSql),"","",-1);

	   conSORT->Open(_bstr_t(conSORTSql),"","",-1);

	   condbPRJ->Open(_bstr_t(condbprjSql),"","",-1);

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
		return FALSE;
	}
	return TRUE;
}

BOOL CDlgSelEngVol::OnInitDialog() 
{
	CDialog::OnInitDialog();
	this->SetWindowCenter(this->m_hWnd);
	if( !this->ConnectDatabase())  //�������ݿ⡣
	{
		return false;
	}
	//7/23
	CString strSQL;
	_bstr_t sql;
	_variant_t key;
	_RecordsetPtr SaveCurrentSet;//���浱ǰѡ��ļ�¼��
	//7/28
	try
	{
		if(!tdfExists(condbPRJ,"savecurrent"))//�жϱ�savecurrent�����ݿ����Ƿ����
		{
			if(CreateTableSaveCurrent()==false)//�����ڽ�����savecurent
			{
				return false;
			}
		}
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return false;
	}
	//7/28

//	bool SaveCurrent=false;
	try
	{
		SaveCurrentSet.CreateInstance(_uuidof(Recordset));
		SaveCurrentSet->Open(_bstr_t(_T("select * from SaveCurrent")),(IDispatch *)condbPRJ,adOpenStatic,\
						adLockOptimistic,adCmdText);
		if(SaveCurrentSet->GetRecordCount()==0)
		{
			//�����浱ǰֵ�ı�SaveCurrentSetû�м�¼ʱ,����Ĭ��ֵ
			m_SelPrjID="F4412";
			m_SelSpecID=3;
			m_SelDsgnID=4;
			m_SelHyID=0;
			m_SelVlmID=8;
			strSQL="insert into SaveCurrent  (selspecid,seldsgnid,selprjid,selhyid,selvlmid)";
			strSQL=strSQL+" values(3,4,'F4412',0,8)";
			sql=strSQL;
			condbPRJ->Execute(sql,NULL,adCmdText);
		}
		else if(SaveCurrentSet->GetRecordCount()>0)//7/24
		{
				//�ָ�Ϊ���һ�ε�״̬
				key=SaveCurrentSet->GetCollect("selprjid");
				if(key.vt==VT_EMPTY)
				{
					m_SelPrjID="F4412";
				}
				else
				{
					m_SelPrjID=key.bstrVal;
				}
				key=SaveCurrentSet->GetCollect("seldsgnid");
				m_SelDsgnID=(key.vt==VT_EMPTY)?4:key.lVal;
				key=SaveCurrentSet->GetCollect("selspecid");
				m_SelSpecID=(key.vt==VT_EMPTY)?3:key.lVal;
				key=SaveCurrentSet->GetCollect("selhyid");
				m_SelHyID=(key.vt==VT_EMPTY)?0:key.lVal;
				key=SaveCurrentSet->GetCollect("selvlmid");//7/24
				m_SelVlmID=(key.vt==VT_EMPTY)?8:key.lVal;//7/24
		}
	}
	catch(_com_error e)
	{
    	AfxMessageBox(e.ErrorMessage());
		return false;
	//	SaveCurrent=true;
	}
	
	//7/23
//7/24
	if(SaveCurrentSet->State==adStateOpen)
	{
		SaveCurrentSet->Close();
		SaveCurrentSet.Release();
	}
//7/24
	this->InitENG();
	this->InitPrjDb();
	this->InitDBVlm();

	m_DBGeng.SetAllowRowSizing( FALSE );
	m_DBGvlm.SetAllowRowSizing( FALSE );

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
//��ʼ���̶�Ӧ�Ŀؼ���
void CDlgSelEngVol::InitENG()
{
	try
	{
	m_DataEng.CreateInstance(__uuidof(Recordset));
	m_DataEng->CursorLocation = adUseClient;

	m_DataEng->Open(_bstr_t(_T("select * from Engin")),(IDispatch *)conPRJDB,adOpenStatic,\
					adLockOptimistic,adCmdText);
	CString cTmp = _T("EnginID=\'")+Trim(m_SelPrjID)+_T("\'");
	m_DataEng->Find(_bstr_t(cTmp), 0, adSearchForward);
	if( m_DataEng->adoEOF )
	{ 
		m_DataEng->MoveLast();
	}
	if( m_DataEng->BOF )
	{
		m_DataEng->MoveFirst();
	}
	m_SelPrjID = vtos(m_DataEng->GetCollect("EnginID"));
	m_DBGeng.SetRefDataSource( m_DataEng->GetDataSource());
	
	m_DBGeng.SetAllowAddNew(TRUE);
	m_DBGeng.SetAllowUpdate(TRUE);
	m_DBGeng.SetAllowDelete(TRUE);
	m_DBGeng.SetForeColor( 0xff0000 );
	m_DBGeng.SetBackColor( 0xffc0c0 );
	this->SetDBGridColumnCaptionAndWidth(m_DBGeng,"teng");


	}
	catch(_com_error * e)
	{
		AfxMessageBox(e->Description());
	}

}
//��ʼ����Ӧ�Ŀؼ���
void CDlgSelEngVol::InitDBVlm()
{
	m_DataVlm.CreateInstance(__uuidof(Recordset));
	m_DataVlm->CursorLocation = adUseClient;

	CString strSql = _T("SELECT * FROM Volume WHERE EnginID ");
	strSql += (m_SelPrjID ==_T(""))? _T("IS NULL"): (_T("= \'") + m_SelPrjID + _T("\'"));
	strSql += _T(" AND SJHYID=") + ltos(m_SelHyID);
	strSql += _T(" AND SJJDID=") + ltos(m_SelDsgnID);
	strSql += _T(" AND ZYID=") + ltos(m_SelSpecID);
	strSql += _T(" ORDER BY jcdm");
    try
	{
		m_DataVlm->Open(_bstr_t(strSql),(IDispatch*)conPRJDB,\
			adOpenStatic,adLockOptimistic,adCmdText);
		m_DataVlm->Find(_bstr_t(CString(_T("VolumeID=") )+ ltos(m_SelVlmID)),0,adSearchForward);

		if(m_DataVlm->adoEOF) 
		{
			if(m_DataVlm->BOF)      //�����¼����û��¼��Ĭ������һ����¼��
			{
				long maxvid=GetMaxVolumeID()+1;
				m_DataVlm->AddNew();
				m_DataVlm->PutCollect(_T("VolumeID"),_variant_t(maxvid));
				m_DataVlm->PutCollect(_T("EnginID"),_variant_t(m_SelPrjID));
//				m_DataVlm->PutCollect(_T("jcdm"),_variant_t(m_SelJcdm));
				m_DataVlm->PutCollect(_T("SJHYID"),_variant_t(m_SelHyID));
				m_DataVlm->PutCollect(_T("SJJDID"),_variant_t(m_SelDsgnID));
				m_DataVlm->PutCollect(_T("ZYID"),_variant_t(m_SelSpecID));
				m_DataVlm->Update();
			}
			else
				m_DataVlm->MoveLast();
		}
		m_SelHyID = vtoi(m_DataVlm->GetCollect(_T("SJHYID")));
		m_SelDsgnID = vtoi(m_DataVlm->GetCollect(_T("SJJDID")));
		m_SelSpecID = vtoi(m_DataVlm->GetCollect("ZYID"));
		
		//������¼����ؼ���Ӧ��
		m_DBGvlm.SetRefDataSource( m_DataVlm->GetDataSource());

		m_DBGvlm.SetAllowAddNew(TRUE);  //�������Ӽ�¼��
		m_DBGvlm.SetAllowUpdate(TRUE);  //
		m_DBGvlm.SetAllowDelete(TRUE);
		m_DBGvlm.SetForeColor(0xff0000);  //������ɫ��
		m_DBGvlm.SetBackColor(0xffc0c0);

		m_DBGvlm.GetColumns().GetItem(_variant_t("VolumeID")).SetVisible(FALSE);//���ø�����Ϊ���ɼ���
		m_DBGvlm.GetColumns().GetItem(_variant_t("EnginID")).SetVisible(FALSE);  
		m_DBGvlm.GetColumns().GetItem(_variant_t("SJJDID")).SetVisible(FALSE);
		m_DBGvlm.GetColumns().GetItem(_variant_t("SJHYID")).SetVisible(FALSE);  //���ø�����Ϊ���ɼ���
		m_DBGvlm.GetColumns().GetItem(_variant_t("ZYID")).SetVisible(FALSE);
		
		this->SetDBGridColumnCaptionAndWidth(m_DBGvlm,_T("tvlm"));//�����ֶεĿ�Ⱥͱ��⡣

		m_DBGvlm.GetColumns().GetItem(_variant_t("VolumeID")).SetVisible(FALSE);//���ø�����Ϊ���ɼ���

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
	}
}

//��ʼ����ҵ ��ƽ׶� רҵ
void CDlgSelEngVol::InitPrjDb()
{
	_variant_t temp;
	m_DataSpc.CreateInstance(__uuidof(Recordset));//רҵ
	m_DataDsgn.CreateInstance(__uuidof(Recordset));//��ƽ׶�
//7/24	m_DataCurrWork.CreateInstance(__uuidof(Recordset));//7/16
	m_DataCategory.CreateInstance(__uuidof(Recordset));	//��ҵ7/16
	
	CString strSQL;
	strSQL="select * from DrawSize";

//7/24	CString strSQLCurrentWork;//7/16
//7/24	strSQLCurrentWork="select * from CurrentWork";//7/16
	try
	{
		//��ʼ����ҵ
		if(m_DataCategory->State==adStateOpen)
		{
			m_DataCategory->Close();
		}
		m_DataCategory->Open(strSQL.GetBuffer(strSQL.GetLength()+1),
		_variant_t((IDispatch*)conSORT,true),adOpenStatic,
		adLockOptimistic,adCmdText);

//7/24		m_DataCurrWork->Open(strSQLCurrentWork.GetBuffer(strSQL.GetLength()+1),
//7/24		_variant_t((IDispatch*)condbPRJ,true),adOpenStatic,
//7/24		adLockOptimistic,adCmdText);

//7/24		m_DBComboCategory.m_Rs=m_DataCurrWork;
		m_DBComboCategory.m_RowRs=m_DataCategory;
		m_DBComboCategory.m_Field=_T("SJHYINDEX");
		m_DBComboCategory.m_BoundField=_T("SJHYID");
		m_DBComboCategory.m_ListField=_T("SJHY");
		m_DBComboCategory.LoadList();
		
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}
	if(m_DataCategory->GetRecordCount()==0)
		return;
	//7/21
	m_DataCategory->MoveFirst();
//	m_SelHyID = vtoi(m_DataCategory->GetCollect("SJHYID"));  //��
	CString str_sjhyid;
	str_sjhyid.Format("%ld",m_SelHyID);
	CString cTmp="sjhyid="+str_sjhyid+"";
	m_DataCategory->Find(_bstr_t(cTmp), 0, adSearchForward);


	//��ʼ����ƽ׶�
	strSQL="select * from DesignStage where sjhyid="+str_sjhyid+"";//7/18
	try
	{
		if(m_DataDsgn->State==adStateOpen)
		{
			m_DataDsgn->Close();
		}
		m_DataDsgn->Open(strSQL.GetBuffer(strSQL.GetLength()+1),
		_variant_t((IDispatch*)conSORT,true),adOpenStatic,
		adLockOptimistic,adCmdText);
//7/24		m_DBComboDSGN.m_Rs=m_DataCurrWork;
		m_DBComboDSGN.m_RowRs=m_DataDsgn;
		m_DBComboDSGN.m_Field=_T("sjjddm");
		m_DBComboDSGN.m_BoundField=_T("sjjddm");
		m_DBComboDSGN.m_ListField=_T("sjjdmc");
		m_DBComboDSGN.LoadList();

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}
//��ʼ��רҵ
	strSQL="select * from Speciality where sjhyid="+str_sjhyid+"";//7/18
	try
	{
		if(m_DataSpc->State==adStateOpen)
		{
			m_DataSpc->Close();
		}
		m_DataSpc->Open(strSQL.GetBuffer(strSQL.GetLength()+1),
		_variant_t((IDispatch*)conSORT,true),adOpenStatic,
		adLockOptimistic,adCmdText);
//7/24		m_DBComboSpec.m_Rs=m_DataCurrWork;
		m_DBComboSpec.m_RowRs=m_DataSpc;
		m_DBComboSpec.m_Field=_T("zydm");
		m_DBComboSpec.m_BoundField=_T("zydm");
		m_DBComboSpec.m_ListField=_T("zymc");
		m_DBComboSpec.LoadList();

		
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}
	//7/19���õ�ǰֵ����λ��¼��
	m_SelDsgnID=(m_SelDsgnID<=0)?4:m_SelDsgnID;
	//7/21
	int i,j;
	if((i=m_DataSpc->GetRecordCount())==0||(j=m_DataDsgn->GetRecordCount())==0)
		return;
	//721
	m_DataDsgn->MoveFirst();
	CString str_sjjdID;
	str_sjjdID.Format("%ld",m_SelDsgnID);
	cTmp="sjjdid="+str_sjjdID+"";
	m_DataDsgn->Find(_bstr_t(cTmp), 0, adSearchForward);
	m_SelSpecID=(m_SelSpecID<=0)?3:m_SelSpecID;
	m_DataSpc->MoveFirst();
	CString str_zyid;
	str_zyid.Format("%ld",m_SelSpecID);
	cTmp="zyid="+str_zyid+"";
	m_DataSpc->Find(_bstr_t(cTmp),0,adSearchForward);
	m_DBComboCategory.RefLst();
	m_DBComboDSGN.RefLst();
	m_DBComboSpec.RefLst();
}
//���ô�����ʾλ�ã�Ϊ��Ļ���м䣩
void CDlgSelEngVol::SetWindowCenter(HWND hWnd)
{
	long SW,SH;
	if( hWnd == NULL )
	{
		return;
	}
	SW = ::GetSystemMetrics( SM_CXSCREEN );
	SH = ::GetSystemMetrics( SM_CYSCREEN );
	CRect rc;
	::GetWindowRect( hWnd, &rc);
	long x,y;
	x = (SW - rc.Width())/2;
	y = (SH - rc.Height())/2;
	::SetWindowPos( hWnd, NULL, x, y, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);

}

CString CDlgSelEngVol::Trim(LPCTSTR pcs)
{
	CString s = pcs;
	s.TrimLeft();
	s.TrimRight();
	return s;
}

CString CDlgSelEngVol::vtos(_variant_t &v)
{
	CString ret;
	switch(v.vt)
	{
	case VT_NULL:
	case VT_EMPTY:
		ret="";
		break;
	case VT_BSTR:
		ret=V_BSTR(&v);
		break;
	case VT_R4:
		ret.Format("%g", (double)V_R4(&v));
		break;
	case VT_R8:
		ret.Format("%g",V_R8(&v));
		break;
	case VT_I2:
		ret.Format("%d",(int)V_I2(&v));
		break;
	case VT_I4:
		ret.Format("%d",(int)V_I4(&v));
		break;
	case VT_BOOL:
		ret=(V_BOOL(&v) ? "True" : "False");
		break;
	}
	ret.TrimLeft(); ret.TrimRight();
	return ret;
}


// ����datagrid�ؼ����п�ͱ��⡣
void CDlgSelEngVol::SetDBGridColumnCaptionAndWidth(CDataGrid &MyDBGrid, CString tbn)
{
	_RecordsetPtr rs;
	rs.CreateInstance(__uuidof(Recordset));
	rs->CursorLocation = adUseClient;
	int i, iWC;
	try
	{
		_variant_t tmpvar, ix;
		CString sTmp,str;

		tbn.TrimLeft();
		tbn.TrimRight();
		for(i=0; i<MyDBGrid.GetColumns().GetCount(); i++)  //
		{
			ix.ChangeType(VT_I4);
			ix.intVal=i;
			sTmp = MyDBGrid.GetColumns().GetItem(ix).GetDataField();   //���Ӣ�ı��⡣

			if( sTmp == _T("VolumeID") )                       //���ظ��С�
			{
				MyDBGrid.GetColumns().GetItem(ix).SetWidth( 0 );
				continue;
			}
			if( sTmp == _T("EnginID") )
			{
				sTmp = _T("gcdm");
			}
		    CString strSQL=_T("SELECT * FROM "+tbn+" WHERE FieldName='"+sTmp+"' ");
		    rs = conSORT->Execute(_bstr_t(strSQL),NULL,adCmdText);

			if(!rs->adoEOF && !rs->BOF)
			{   
				str = vtos( rs->GetCollect(_T("LocalCaption")));
				MyDBGrid.GetColumns().GetItem(ix).SetCaption(str);  //����ÿ�еı���Ϊ���ģ������ݿ�ȡ��

				str =vtos(rs->GetCollect(_T("width")));
						
				iWC=atoi(str);
				//7/27
				RECT rectEng;
				float feng;
				m_DBGeng.GetWindowRect(&rectEng);
				feng=rectEng.right-rectEng.left;
				iWC=iWC>(feng/3)?(feng/3):iWC;
				MyDBGrid.GetColumns().GetItem(ix).SetWidth((float)iWC);  //ÿ�еĿ�ȡ�
					
				//7/27
//7/27				iWC=iWC>0 ? iWC : 200;
//7/27				MyDBGrid.GetColumns().GetItem(ix).SetWidth((float)iWC/30);  //ÿ�еĿ�ȡ�
						
			}
		}
		
		
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.Description());
	}
    rs.Release();
}

int CDlgSelEngVol::vtoi(_variant_t& v)
{	
	CString tmps;
	int ret=0;
	switch(v.vt)
	{
	case VT_NULL:
		ret= 0;
		break;
	case VT_I2:
		ret= (int)V_I2(&v);
		break;
	case VT_I4:
		ret= V_I4(&v);
		break;
	case VT_R4:
		ret= (int)V_R4(&v);
		break;
	case VT_R8:
		ret =(int)V_R8(&v);
		break;
	case VT_BSTR:
		tmps=(char*)_bstr_t(&v);
		ret=atoi((LPCSTR)tmps);
		break;
	}
	return ret;
}
CString inline CDlgSelEngVol::ltos(long v)
{
	CString s;
	s.Format("%d",v);
	return s;
}

long CDlgSelEngVol::GetMaxVolumeID()
{
	long ret=0;
	try
	{
		_RecordsetPtr rs;
		rs = conPRJDB->Execute(_bstr_t(_T("Select MAX(VolumeID) FROM [Volume]")),NULL,adCmdText);
		ret=vtoi(rs->GetCollect(_variant_t((long)0)));
		rs->Close();
		rs=NULL;
	}
	catch(...)
	{
	}
	return ret;

}

BEGIN_EVENTSINK_MAP(CDlgSelEngVol, CDialog)
    //{{AFX_EVENTSINK_MAP(CDlgSelEngVol)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, 218 /* RowColChange */, OnRowColChangeDbgeng, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGVIM, 218 /* RowColChange */, OnRowColChangeDbgvim, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, 217 /* OnAddNew */, OnOnAddNewDbgeng, VTS_NONE)
	ON_EVENT(CDlgSelEngVol, IDC_DBGVIM, 217 /* OnAddNew */, OnOnAddNewDbgvim, VTS_NONE)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, 204 /* AfterUpdate */, OnAfterUpdateDbgeng, VTS_NONE)
	ON_EVENT(CDlgSelEngVol, IDC_DBGVIM, 204 /* AfterUpdate */, OnAfterUpdateDbgvim, VTS_NONE)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, 213 /* ColResize */, OnColResizeDbgeng, VTS_I2 VTS_PI2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGVIM, 213 /* ColResize */, OnColResizeDbgvim, VTS_I2 VTS_PI2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGVIM, 209 /* BeforeUpdate */, OnBeforeUpdateDbgvim, VTS_PI2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, 209 /* BeforeUpdate */, OnBeforeUpdateDbgeng, VTS_PI2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, 207 /* BeforeDelete */, OnBeforeDeleteDbgeng, VTS_PI2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGVIM, 207 /* BeforeDelete */, OnBeforeDeleteDbgvim, VTS_PI2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, 212 /* ColEdit */, OnColEditDbgeng, VTS_I2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGVIM, 212 /* ColEdit */, OnColEditDbgvim, VTS_I2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, -602 /* KeyDown */, OnKeyDownDbgeng, VTS_PI2 VTS_I2)
	ON_EVENT(CDlgSelEngVol, IDC_DBGENG, -604 /* KeyUp */, OnKeyUpDbgeng, VTS_PI2 VTS_I2)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

////ѡ��ͬ���п��У�ͬʱ���¾����Ϣ��
void CDlgSelEngVol::OnRowColChangeDbgeng(VARIANT FAR* LastRow, short LastCol)   //�ı䵱ǰ�л��С�
{//8/13
	if(m_DBGeng.GetAddNewMode()==2||m_DBGeng.GetAddNewMode()==1)
	{
		m_DBGvlm.SetAllowAddNew(false);
	}
	else
	{
		m_DBGvlm.SetAllowAddNew(true);
	}
	//8/13
//	if(m_bIsNew&&m_DataEng->GetEditMode()!=adEditAdd)//7/26 8/14
//	{
//		m_bIsNew=false;//7/26	
//	}
	try
	{
		try
		{
			if( m_lastEngRow != m_DBGeng.GetRow() )          //7/24{{��ס���е�һ��ѡ��ʱ�Ĺ��̴��ţ���Ϊ�޸�ʱ��������
			{
				m_currEngID = vtos(m_DataEng->GetCollect("EnginID"));
				m_lastEngRow = m_DBGeng.GetRow();
			}                                      //}}

			if(m_DBGvlm.GetAddNewMode()==2)       //ѡ�����һ��ͬʱ�޸��ˡ��������¼�¼��
			{								
				m_DataVlm->Update();			
				if(!m_DataVlm->adoEOF) m_DataVlm->MoveLast();
			}
	//		else if(m_DBGvlm.GetAddNewMode()==0)        //û�����ӡ�//8/16ԭ����
			else if(!(m_DataVlm->BOF||m_DataVlm->adoEOF)&&m_DBGvlm.GetAddNewMode()==0)
			{
					m_DataVlm->Update();	
			}
//			else if(m_DBGvlm.GetAddNewMode()==1)    //8/12    //ѡ�����һ�е�û���޸ġ�
//			{
//				if(m_DBGeng.GetAddNewMode()!=1)//8/12
//				m_DataVlm->Update();//8/12
//			}
		}
		catch (...)
		{
			m_DataVlm->CancelUpdate();   //����ʱ����
		}

		m_SelPrjID = vtos(m_DataEng->GetCollect(_T("EnginID")));
		m_SelDsgnID = vtoi(m_DataDsgn->GetCollect(_T("SJJDID")));
		m_SelSpecID = vtoi(m_DataSpc->GetCollect(("ZYID")));
		m_SelHyID = vtoi(m_DataCategory->GetCollect(_T("SJHYID")));
		
		m_DBGvlm.SetRefDataSource(NULL);

		CString tmpSQL=CString(_T("SELECT *   FROM [Volume] WHERE EnginID ")) +(m_SelPrjID ==_T("") ? CString(_T(" IS NULL ")) : CString(_T("= '"+m_SelPrjID+"' "))) ;
		tmpSQL +=_T(" AND SJHYID=") + ltos(m_SelHyID);
		tmpSQL +=_T(" AND SJJDID=") + ltos(m_SelDsgnID);
		tmpSQL +=_T(" AND ZYID=") + ltos(m_SelSpecID);
		tmpSQL +=_T(" ORDER BY jcdm ");
		//����Ǵ򿪵ģ������ȹر�,��������ر�ǰ���������ڱ༭�ļ�¼����update,�������
		if(m_DataVlm->State==adStateOpen)
			m_DataVlm->Close();
		m_DataVlm->Open(_bstr_t(tmpSQL),(IDispatch*)conPRJDB,adOpenStatic,adLockOptimistic,adCmdText);

		if(m_DataVlm->adoEOF)
		{
			if(!m_DataVlm->BOF)
				m_DataVlm->MoveLast();			
		}
		
		m_DBGvlm.SetRefDataSource(m_DataVlm->GetDataSource());

		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("EnginID"))).SetVisible(FALSE);
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("VolumeID"))).SetVisible(FALSE);
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("SJJDID"))).SetVisible(FALSE);
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("SJHYID"))).SetVisible(FALSE);
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("ZYID"))).SetVisible(FALSE);

		this->SetDBGridColumnCaptionAndWidth (m_DBGvlm, _T("tvlm"));

	}
	catch(_com_error e)
	{
#ifdef _DEBUG	
		::MessageBox(NULL,e.Description(), "��ʾ",MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST);
#endif
	}
}

//ѡ��ͬ���л��У�ͬʱ����/������Ϣ��
void CDlgSelEngVol::OnRowColChangeDbgvim(VARIANT FAR* LastRow, short LastCol) 
{	
//	if(m_bIsVlmNew&&m_DataVlm->GetEditMode()!=adEditAdd)//7/26�����Ƿ��ڱ༭״̬8/14
//	{
//		m_bIsVlmNew=false;
//	}
	try
	{
		try
		{
			if( m_lastVlmRow != m_DBGvlm.GetRow() )          //7/24{{��ס���е�һ��ѡ��ʱ�Ĺ��̴��ţ���Ϊ�޸�ʱ��������
			{
				m_currVlmID = vtos(m_DataVlm->GetCollect("jcdm"));
				m_lastVlmRow = m_DBGvlm.GetRow();
			}                                      //}}

			if(m_DBGeng.GetAddNewMode()==2)       ////ѡ�����һ��ͬʱ�޸��ˡ��������¼�¼��
			{
				m_DataEng->Update();
				if(!m_DataEng->adoEOF)
					m_DataEng->MoveLast();
			}
			else if(m_DBGeng.GetAddNewMode()==0)               //�������һ�С�
			{
				m_DataEng->Update();
			}
//			else if(m_DBGeng.GetAddNewMode()==1)                ////ѡ�����һ�е�û���޸ġ�
//			{
//				if(m_DBGeng.GetAddNewMode()!=1)
//					m_DataEng->Update();
//			}
		}
		catch(...)
		{
			m_DataEng->CancelUpdate();       //ȡ�����¡�
		}
	}
	catch(_com_error e)
	{
	    ::MessageBox(NULL,e.Description(), "��ʾ",MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST);
	}
}

//�����¹��̼�¼��
void CDlgSelEngVol::OnOnAddNewDbgeng() 
{
	m_bIsNew = true;
	m_lastEngRow = m_DBGeng.GetRow();//7/26
	//8/13
	m_DBGvlm.SetAllowAddNew(false);
	//8/13

}
//��������¼�¼��
void CDlgSelEngVol::OnOnAddNewDbgvim() 
{
	m_bIsVlmNew = true;	
}

//�޸�֮�󣬸ı�״̬��
void CDlgSelEngVol::OnAfterUpdateDbgeng() 
{
	m_bIsNew = false;
    m_bEngDelete = false;                       //ɾ���ɹ���
	m_stateEditEng=false;//7/25�����Ƿ��ڱ༭״̬
	//	m_bVlmDelete = false;
//	m_okSelEng.EnableWindow(true);//7/25
//	m_cancalSelEng.EnableWindow(true);//7/25
}

void CDlgSelEngVol::OnAfterUpdateDbgvim() 
{
	m_bIsVlmNew = false;
	m_bVlmDelete = false;
	m_stateEditVlm=false;//7/25����Ƿ��ڱ༭״̬
	
}


void CDlgSelEngVol::OnColResizeDbgeng(short ColIndex, short FAR* Cancel) 
{
	SaveDBGridColumnCaptionAndWidth(m_DBGeng, ColIndex, "teng");
}

void CDlgSelEngVol::OnColResizeDbgvim(short ColIndex, short FAR* Cancel) 
{
	SaveDBGridColumnCaptionAndWidth(m_DBGvlm, ColIndex, "tvlm");

}

//���浱ǰ�п�
void CDlgSelEngVol::SaveDBGridColumnCaptionAndWidth(CDataGrid &MyDBGrid, long ColIndex, CString tbn)
{  
	try
	{
		_RecordsetPtr rs;
		rs.CreateInstance(__uuidof(Recordset));		
		
		CString SQLx, sTmp;
		float sngWidth;
		_variant_t ix;

		ix.ChangeType(VT_I4);
		ix.intVal =ColIndex;

		sTmp=MyDBGrid.GetColumns().GetItem(ix).GetDataField();
		if(sTmp == _T("EnginID") ) //����һ�¡�
			sTmp = "gcdm";
	
//		sngWidth=MyDBGrid.GetColumns().GetItem(ix).GetWidth()*30;
		//7/27
		sngWidth=MyDBGrid.GetColumns().GetItem(ix).GetWidth();
//		RECT rectEng;
//		float feng;
//		m_DBGeng.GetWindowRect(&rectEng);
//		feng=rectEng.right-rectEng.left;
		//7/27
		
		CString wh;
		wh.Format("%f",sngWidth);
		SQLx = _T("UPDATE "+tbn+" SET width = "+wh+"  WHERE FieldName = '"+sTmp+"' ");

		rs = conSORT->Execute(_bstr_t(SQLx), NULL, adCmdText);

		rs.Release();
	}
	catch(CString e)
	{
		AfxMessageBox(e);
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
	}


}
//�޸�֮ǰ�����¼�¼�������ֶθ�ֵ��
void CDlgSelEngVol::OnBeforeUpdateDbgvim(short FAR* Cancel) 
{	

	try
	{
		if( m_bVlmDelete )
			return;

		CString strID = vtos(m_DataVlm->GetCollect("jcdm"));  //��ǰ�ľ����š�
		CString strVlmName = vtos(m_DataVlm->GetCollect("jcmc"));
		CString strMessage = "";

		if(strID.IsEmpty() )        strMessage = "������";
		else if(strVlmName.IsEmpty() )  strMessage = "�������";

		if(!strMessage.IsEmpty() )
		{
			if( m_bIsVlmNew && strID.IsEmpty() && strVlmName.IsEmpty() )       //���û�ȡ������ʱ��
			{
				CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp(); 
				pApp->InitHook();										//����DataGrid�ؼ������ġ�ȡ���ɹ����Ի���

				*Cancel=1;                                  //ȡ�����ӡ�
				m_bIsVlmNew=false;//7/26
				m_DataVlm->CancelUpdate();
				return;
			}
				
			::MessageBox(NULL," "+strMessage+" ����Ϊ�գ�", "��ʾ",MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST);
		
			CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp(); 
			pApp->InitHook();										//����DataGrid�ؼ������ġ�ȡ���ɹ����Ի���

			*Cancel=1;                                  //ȡ���޸ġ�
			return;
		}
		if( m_currVlmID == strID )          //7/24
		{
			return;
		}

		CString tHyID = ltos(m_SelHyID),
		tDsgnID = ltos(m_SelDsgnID), 
		tSpecID = ltos(m_SelSpecID);

		CString strSQL = _T("SELECT VolumeID FROM [Volume] WHERE EnginID='"+m_SelPrjID+"' ") +
		          _T("AND jcdm='"+strID+"' ") +				
			      _T("AND SJHYID="+tHyID+" ") +
				  _T("AND SJJDID="+tDsgnID+" ") +
				  _T("AND ZYID="+tSpecID+" "); 
		_RecordsetPtr rs;
		rs.CreateInstance(__uuidof(Recordset));
		rs->Open(_bstr_t(strSQL), (IDispatch* )conPRJDB, adOpenStatic, 
			adLockOptimistic, adCmdText);
		long k;
//		if( k = rs->GetRecordCount() > 0)//8/14 ԭ����
		//8/14
		k=rs->GetRecordCount();
		if((k>1&&(!(m_DBGvlm.GetAddNewMode()==1||m_DBGvlm.GetAddNewMode()==2)))||k>0&&(m_DBGvlm.GetAddNewMode()==1||m_DBGvlm.GetAddNewMode()==2))

		//8/14
		{ 
		   ::MessageBox(NULL, "�� "+strID+" ������Ѿ����ڣ����������룡", "��ʾ", MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST );
		   CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp();
		   pApp->InitHook();
		   *Cancel=1;	
  		   m_DataVlm->CancelUpdate();       //ȡ�����¡�

		   return;
		}

		if(!m_bIsVlmNew )                       //�����ӡ�
		{
    		return;
		}
		if(m_SelPrjID.IsEmpty() )
		{
			::MessageBox(NULL, "��ѡ��һ�����̼�¼��", "��ʾ", MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST );
			CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp(); 
			pApp->InitHook();										//����DataGrid�ؼ������ġ�ȡ���ɹ����Ի���
			m_DataVlm->CancelUpdate();
			*Cancel=1;                                  //ȡ���޸ġ�
			return;

		}

		long maxvid=GetMaxVolumeID()+1;
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("EnginID"))).SetText(m_SelPrjID);
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("VolumeID"))).SetText(ltos(maxvid));
		
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("SJJDID"))).SetText(ltos(m_SelDsgnID));
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("ZYID"))).SetText(ltos(m_SelSpecID));
		m_DBGvlm.GetColumns().GetItem(_variant_t(_T("SJHYID"))).SetText(ltos(m_SelHyID));	
		
	  }
	catch(_com_error& e)
	{
		::MessageBox(NULL, e.Description(), "��ʾ", MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST );
	}
	
}

void CDlgSelEngVol::OnBeforeUpdateDbgeng(short FAR* Cancel) 
{
	try
	{
		if( m_bEngDelete ) //�����¼�����
			return;

		CString strID, strGcmc, strUnitNum;         // {{7��24
					 
     	strID = vtos(m_DataEng->GetCollect(_T("EnginID")));
		strGcmc = vtos(m_DataEng->GetCollect("gcmc"));
		strUnitNum = vtos(m_DataEng->GetCollect("UnitNum"));
		
		CString strMessage = _T("");                         
		if(strID.IsEmpty()) strMessage = "���̴���";
		else if(strGcmc.IsEmpty() )  strMessage = "��������";
	//	else if(strUnitNum.IsEmpty() ) strMessage = "����̨��";//7/27

    	if( !strMessage.IsEmpty() )                            
		{
			if(m_bIsNew && strID.IsEmpty() && strGcmc.IsEmpty() && strUnitNum.IsEmpty() ) //
			{
				CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp(); 
				pApp->InitHook();										//����DataGrid�ؼ������ġ�ȡ���ɹ����Ի���

				*Cancel=1;                                  //ȡ�����ӡ�
				m_bIsNew=false;	//7/26
				m_DataEng->CancelUpdate();
				return;

			}
			
			::MessageBox(NULL, " "+strMessage+"����Ϊ��", "��ʾ", MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST );
			CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp();
			pApp->InitHook();
			*Cancel = 1;
			return;                                                         //7/24 }}

		}
		_RecordsetPtr pRs;
		pRs.CreateInstance(__uuidof(Recordset));

		pRs->Open(_bstr_t(_T("select * from Engin where EnginID='"+strID+"' ")),(IDispatch *)conPRJDB,adOpenStatic,\
						adLockOptimistic,adCmdText);
		
		if( m_currEngID == strID )          //7/24
		{
			return;
		}

	//	if( pRs->GetRecordCount() > 0 ) //�ظ���¼��8/18ԭ����
		//8/18
		int jj;
		jj=pRs->GetRecordCount();
		if((jj>1&&(!(m_DBGeng.GetAddNewMode()==1||m_DBGeng.GetAddNewMode()==2)))||jj>0&&(m_DBGeng.GetAddNewMode()==1||m_DBGeng.GetAddNewMode()==2))
		//8/18
		{
		   ::MessageBox(NULL, "�� "+strID+" �������Ѿ����ڣ����������룡", "��ʾ", MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST );
		   CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp();
		   pApp->InitHook();
		   *Cancel=1;	
		   pRs->Close();
		   pRs.Release();
		   return;
		}
		pRs->Close();
		pRs.Release();
	}
	catch(_com_error& e)
	{
		::MessageBox(NULL, e.Description(), "��ʾ", MB_OK|MB_ICONEXCLAMATION|MB_TASKMODAL|MB_TOPMOST );
	}

}


void CDlgSelEngVol::OnBeforeDeleteDbgeng(short FAR* Cancel) 
{
	try
	{
		CString sTmp, strSQL;
		sTmp.Format("����׼��ɾ���� %s ��  ���̼���������ص�ԭʼ���ݺͽ�����ݼ�¼��\
			     \n����������ǡ������޷�����ɾ��������ȷʵҪִ��ɾ��������",
				 vtos(m_DataEng->GetCollect("gcmc")) );
		if( ::MessageBox(NULL, sTmp, "ɾ������", MB_YESNO|MB_ICONEXCLAMATION ) == IDNO )
		{
            CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp();
			pApp->InitHook();
			*Cancel = 1;
			return;
		}
		CString tHyID = ltos(m_SelHyID),
			    tDsgnID = ltos(m_SelDsgnID), 
				tSpecID = ltos(m_SelSpecID);

		strSQL = _T("DELETE FROM Volume WHERE EnginID='"+m_SelPrjID+"' ") +
			      _T("AND SJHYID="+tHyID+" ") +
				  _T("AND SJJDID="+tDsgnID+" ") +
				  _T("AND ZYID="+tSpecID+" ");  
		conPRJDB->Execute(_bstr_t(strSQL), NULL, adCmdText);   //ɾ������е������Ϣ��
		m_DBGvlm.SetRefDataSource(NULL);
		m_bEngDelete = true;
	}
	catch(_com_error& e)
	{
		AfxMessageBox(e.Description());
	}
	
}

void CDlgSelEngVol::OnBeforeDeleteDbgvim(short FAR* Cancel) 
{
	try
	{	
		CString sTmp,strSQL;
		sTmp.Format("����׼��ɾ���� %s ��  ��ἰ��������ص�ԭʼ���ݺͽ�����ݼ�¼��\
			     \n����������ǡ������޷�����ɾ��������ȷʵҪִ��ɾ��������",
					vtos(m_DataVlm->GetCollect("jcmc")) );
		if( ::MessageBox(NULL, sTmp, "ɾ�����",MB_YESNO|MB_ICONEXCLAMATION) == IDNO)
		{
			CDlgSelEngVol * pApp = (CDlgSelEngVol *)::AfxGetApp();
			pApp->InitHook();
			*Cancel = 1;
			return;
		}
		else
		{
			CString strSQL;
			LPCTSTR szTab[]={(_T("Z1")),(_T("ZB")),(_T("ZA")),(_T("ZC")),(_T("ZD")),
				(_T("ZG")),(_T("CL")),(_T("ML")),(_T("Z8"))};
			int count = sizeof(szTab)/sizeof(LPCTSTR);
	    	int	vlmID=vtoi(m_DataVlm->GetCollect(_T("VolumeID")));

			for(int i=0; i<count; i++)      //ɾ����������������صļ�¼��
			{
				strSQL.Format((_T("DELETE FROM [%s] WHERE VolumeID=%d")), szTab[i], vlmID);
				try
				{
					conPRJDB->Execute(_bstr_t(strSQL), NULL, adCmdText);
				}
				catch(_com_error e)
				{
#ifdef _DEBUG
					AfxMessageBox(e.Description());
#endif
				}
			}
				m_bVlmDelete = true;       //�����¼״̬��

		}
	}
	catch(_com_error & e)
	{
		AfxMessageBox(e.Description());
	}
	
}

//����ʾ�Ի�������
LRESULT CALLBACK CDlgSelEngVol::CBTProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	CDlgSelEngVol *ptisApp = (CDlgSelEngVol *)::AfxGetApp();
	if( nCode == HCBT_CREATEWND )
	{
		ptisApp->UnInitHook();
		return 1;
	}
	return 	CallNextHookEx(ptisApp->m_hHook, nCode, wParam, lParam);
}


BOOL CDlgSelEngVol::InitHook()
{
	m_hHook = SetWindowsHookEx(WH_CBT, CBTProc, NULL, ::AfxGetThread()->m_nThreadID );
	
	if(m_hHook == NULL)
		return FALSE;

	return TRUE;
}

BOOL CDlgSelEngVol::UnInitHook()
{
	if(m_hHook == NULL)
	{
		return TRUE;
	}
	bool bstr =	UnhookWindowsHookEx(m_hHook);

	m_hHook = NULL;
	return bstr;

}

void CDlgSelEngVol::OnSelchangeComboHY() 
{

	_variant_t strData;
	CString strSQL;
	m_DBComboCategory.OnSelchange();
	strData=m_DataCategory->GetCollect("sjhyid");
	m_SelHyID=(strData.vt==VT_EMPTY)?0:strData.lVal;
	CString str_sjhyid;
	str_sjhyid.Format("%ld",m_SelHyID);
	CString cTmp="sjhyid="+str_sjhyid+"";


	strSQL="select * from DesignStage where sjhyid="+str_sjhyid+"";//7/18
	try
	{
		if(m_DataDsgn->State==adStateOpen)
		{
			m_DataDsgn->Close();
		}
		m_DataDsgn->Open(strSQL.GetBuffer(strSQL.GetLength()+1),
		_variant_t((IDispatch*)conSORT,true),adOpenStatic,
		adLockOptimistic,adCmdText);
		m_DBComboDSGN.LoadList();

	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}

	strSQL="select * from Speciality where sjhyid="+str_sjhyid+"";//7/18
	try
	{
		if(m_DataSpc->State==adStateOpen)
		{
			m_DataSpc->Close();
		}
		m_DataSpc->Open(strSQL.GetBuffer(strSQL.GetLength()+1),
		_variant_t((IDispatch*)conSORT,true),adOpenStatic,
		adLockOptimistic,adCmdText);
		m_DBComboSpec.LoadList();

		
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}
	//7/19���õ�ǰֵ����λ��¼��
	m_SelDsgnID=(m_SelDsgnID<=0)?4:m_SelDsgnID;
	if(m_DataDsgn->GetRecordCount()==0||m_DataSpc->GetRecordCount()==0)
	{
		m_DBGvlm.SetRefDataSource(NULL);
		m_DBComboDSGN.ResetContent();
		m_DBComboSpec.ResetContent();
		m_DBComboDSGN.SetCurSel(-1);
		m_DBComboSpec.SetCurSel(-1);
		return;
	}
	m_DataDsgn->MoveFirst();
	CString str_sjjdID;
	str_sjjdID.Format("%ld",m_SelDsgnID);
	cTmp="sjjdid="+str_sjjdID+"";
	m_DataDsgn->Find(_bstr_t(cTmp), 0, adSearchForward);
	m_SelSpecID=(m_SelSpecID<=0)?3:m_SelSpecID;
	m_DataSpc->MoveFirst();
	CString str_zyid;
	str_zyid.Format("%ld",m_SelSpecID);
	cTmp="zyid="+str_zyid+"";
	m_DataSpc->Find(_bstr_t(cTmp),0,adSearchForward);
	m_DBComboDSGN.RefLst();
	m_DBComboSpec.RefLst();
	changeDsgnSpeGoRY();

	
}

void CDlgSelEngVol::OnSelchangeComboDsgn() 
{
	m_DBComboDSGN.OnSelchange();
	changeDsgnSpeGoRY();
	
}

void CDlgSelEngVol::OnSelchangeComboSpec() 
{
	m_DBComboSpec.OnSelchange();
	changeDsgnSpeGoRY();
	
}
 

//��"��ҵ"��"��ƽ׶�","רҵ"�����ı�ʱ,��Ӧ��"���"��Ӧ�ط����ı�
void CDlgSelEngVol::changeDsgnSpeGoRY()
{
	try
	{
		_variant_t key;
		key=m_DataEng->GetCollect("enginid");
		if(key.vt==VT_EMPTY)//���̵�ID
		{
			m_SelPrjID="";
		}
		else
		{
			m_SelPrjID=key.bstrVal;
		}
	//	key=m_DataEng->GetCollect("gcmc");
//		m_SelPrjName=(key.vt==VT_EMPTY)?"":key.bstrVal;
		key=m_DataDsgn->GetCollect("sjjdid");
		m_SelDsgnID=(key.vt==VT_EMPTY)?0:key.lVal;//��ƽ׶ε�ID
		key=m_DataSpc->GetCollect("zyid");//רҵID
		m_SelSpecID=(key.vt==VT_EMPTY)?0:key.lVal;
		key=m_DataCategory->GetCollect("sjhyid");
		m_SelHyID=(key.vt==VT_EMPTY)?0:key.lVal;//��ҵ��ID
		m_DBGvlm.SetRefDataSource(NULL);
	//	if(m_SelPrjID==""||m_SelDsgnID==0||m_SelSpecID==0)
	//	{
	//		return;
	//	}
		CString strSQL;
		CString str_HyID;
		str_HyID.Format("%ld",m_SelHyID );
		CString str_Dsgn;
		str_Dsgn.Format("%ld",m_SelDsgnID);
		CString str_SpecID;
		str_SpecID.Format("%ld",m_SelSpecID);
		strSQL="select * from Volume where SJHYID="+str_HyID+" and EnginID='"+m_SelPrjID+"'"; 
		strSQL=strSQL+" and SJJDID="+str_Dsgn+" and ZYID="+str_SpecID+" order by jcdm ";
		if(m_DataVlm->State==adStateOpen)
		{
			m_DataVlm->Close();
		}
		m_DataVlm->Open(strSQL.GetBuffer(strSQL.GetLength()+1),
		_variant_t((IDispatch*)conPRJDB,true),adOpenStatic,
		adLockOptimistic,adCmdText);
	}
	catch(_com_error e)
	{
		AfxMessageBox(e.ErrorMessage());
		return;
	}
/*	if(m_DataVlm->GetRecordCount()==0)
	{
		return;
	}*/
	m_DBGvlm.SetRefDataSource(m_DataVlm->GetDataSource());
	m_DBGvlm.GetColumns().GetItem(_variant_t(_T("enginid"))).SetVisible(false);
	m_DBGvlm.GetColumns().GetItem(_variant_t(_T("VolumeID"))).SetVisible(FALSE);
	m_DBGvlm.GetColumns().GetItem(_variant_t(_T("SJJDID"))).SetVisible(FALSE);
	m_DBGvlm.GetColumns().GetItem(_variant_t(_T("SJHYID"))).SetVisible(FALSE);
	m_DBGvlm.GetColumns().GetItem(_variant_t(_T("ZYID"))).SetVisible(FALSE);
	SetDBGridColumnCaptionAndWidth (m_DBGvlm, _T("tvlm"));

}

void CDlgSelEngVol::OnOK() 
{
	_variant_t key, var;
	try
	{
		//ȡ��ҵ��id
		if(m_DataCategory->State!=adStateOpen||m_DataCategory->adoEOF)
		{
			AfxMessageBox("û�к���'��ҵ'��ѡ��,������!");
			return;
		}
		key=m_DataCategory->GetCollect("sjhyid");
		if(key.vt==VT_EMPTY||key.vt==VT_NULL)
		{
			AfxMessageBox("'��ҵ'����Ϣ������,������!");
			return;
		}
		else
		{
			m_SelHyID=key.lVal;	
		}
		//ȡ��ƽ׶ε�id
		if(m_DataDsgn->State!=adStateOpen||m_DataDsgn->adoEOF)
		{
			AfxMessageBox("û�к���'��ƽ׶�'��ѡ��,������!");
			return;
		}
		key=m_DataDsgn->GetCollect("sjjdid");
		if(key.vt==VT_EMPTY||key.vt==VT_NULL)
		{
			AfxMessageBox("'��ƽ׶�'����Ϣ������,������!");
			return;
		}
		else
		{
			m_SelDsgnID=key.lVal;
		}
		//ȡ���̵�id
		if(m_DataEng->State!=adStateOpen||m_DataDsgn->adoEOF)
		{
			AfxMessageBox("û�к���'����'��ѡ��,������!");
			return;
		}
		key=m_DataEng->GetCollect("enginid");
		var=m_DataEng->GetCollect("gcmc");//2005.2.23
		if(key.vt==VT_EMPTY||key.vt==VT_NULL || var.vt==VT_EMPTY||var.vt==VT_NULL)
		{
			AfxMessageBox("'����'��Ϣ������,������!");
			return;
		}
		else
		{
			m_SelPrjID=key.bstrVal;
			m_SelPrjName = var.bstrVal;
		}
		//ȡרҵ��id
		if(m_DataSpc->State!=adStateOpen||m_DataSpc->adoEOF)
		{
			AfxMessageBox("û�к���'רҵ'��ѡ��,������!");
			return;
		}
		key=m_DataSpc->GetCollect("zyid");
		if(key.vt==VT_EMPTY||key.vt==VT_NULL)
		{
			AfxMessageBox("'רҵ'��Ϣ��������������!");
			return;
		}
		else
		{
			m_SelSpecID=key.lVal;
		}
		//ȡ���id
		if(m_DataVlm->State!=adStateOpen||m_DataVlm->adoEOF)
		{
			AfxMessageBox("û�к���'���'��ѡ��,������!");
			return;
		}
		key=m_DataVlm->GetCollect("volumeid");
		if(key.vt==VT_EMPTY||key.vt==VT_NULL)
		{
			AfxMessageBox("'���'����Ϣ������,������!");
			return;
		}
		else
		{
			m_SelVlmID=key.lVal;
		}
		key = m_DataVlm->GetCollect("jcdm");  //2005.2.23
		var = m_DataVlm->GetCollect("jcmc");
		if(key.vt==VT_EMPTY||key.vt==VT_NULL || var.vt==VT_EMPTY||var.vt==VT_NULL)
		{
			AfxMessageBox("'���'����Ϣ������,������!");
			return;
		}
		else
		{
			m_SelVlmCODE = key.bstrVal;
			m_SelVlmName = var.bstrVal;
		}


	}
	catch(_com_error e)
	{
		AfxMessageBox("�뵥���༭���е������е��κεط�!");
		return;
	}
	//����ǰ��ֵ���浽��SaveCurrent;
	//ע���SaveCurrent����һ����¼!!!!!!!
	CString str_SelDsgnID;
	str_SelDsgnID.Format("%ld",m_SelDsgnID);
	CString str_SelHyID;
	str_SelHyID.Format("%ld",m_SelHyID);
	CString str_SelSpecID;
	str_SelSpecID.Format("%ld",m_SelSpecID);
	CString str_SelVlmID;
	str_SelVlmID.Format("%ld",m_SelVlmID);
	_bstr_t sql;
	CString strSQL;
	strSQL="update savecurrent set selspecid="+str_SelSpecID+",seldsgnid="+str_SelDsgnID+"";
	strSQL=strSQL+" , selprjid='"+m_SelPrjID+"',selhyid="+str_SelHyID+",selvlmid="+str_SelVlmID+"";
	sql=strSQL;
	try
	{
		condbPRJ->Execute(sql,NULL,adCmdText);
	}
	catch(_com_error e)
	{
		AfxMessageBox("��������ʧ��!");
		return;
	}
//	ReleaseMemoryEngVol();//7/25
	CDialog::OnOK();
}

//�ر����ݿ�ͼ�¼��
void CDlgSelEngVol::ReleaseMemoryEngVol()
{
	if(m_DataVlm->State==adStateOpen)
	{
		m_DataVlm->Close();
	}
	if(m_DataSpc->State==adStateOpen)
	{
		m_DataSpc->Close();
	}
	if(m_DataDsgn->State==adStateOpen)
	{
		m_DataDsgn->Close();
	}
	if(m_DataCategory->State==adStateOpen)
	{
		m_DataCategory->Close();
	}
	if(m_DataEng->State==adStateOpen)
	{
		m_DataEng->Close();
	}
	if(conPRJDB->State==adStateOpen)
	{
		conPRJDB->Close();
	}
	if(conSORT->State==adStateOpen)
	{
		conSORT->Close();
	}
}

void CDlgSelEngVol::OnCancel() 
{
	CDialog::OnCancel();

}

//�����̴��ڱ༭״̬ʱ����ȡ����ȷ��
void CDlgSelEngVol::OnColEditDbgeng(short ColIndex) //7/25
{
	m_stateEditEng=true;

//	if(m_bIsNew)
//	{
//		m_cancalSelEng.EnableWindow(false);
//	}
//	m_okSelEng.EnableWindow(false);
}
//�����ڱ༭״̬ʱ����ȡ����ȷ��
void CDlgSelEngVol::OnColEditDbgvim(short ColIndex) //7/25
{
	m_stateEditVlm=true;	
}

//�����ݿ���û�б�SaveCurrent ʱ����һ����
BOOL CDlgSelEngVol::CreateTableSaveCurrent()
{
	CString strSQL;
	_bstr_t sql;
	strSQL="create table SaveCurrent (selspecid long,seldsgnid long,selprjid char,selhyid long,selvlmid long) ";
	sql=strSQL;
	try
	{
		condbPRJ->Execute(sql,NULL,adCmdText);
	}
	catch(_com_error e)
	{
		AfxMessageBox("�������ݿ�û�гɹ�!");
		return false;
	}
	return true;
}

//�жϸñ���ָ�������ݿ����Ƿ����
bool CDlgSelEngVol::tdfExists(_ConnectionPtr pConn, CString tbn)
{
	if(pConn==NULL || tbn=="")
		return false;
	_RecordsetPtr rs;
	if(tbn.Left(1)!="[")
		tbn="["+tbn;
	if(tbn.Right(1)!="]")
		tbn+="]";
	try{
		rs=pConn->Execute(_bstr_t(tbn),NULL,adCmdTable);
		rs->Close();
		return true;
	}
	catch(_com_error e)
	{
		return false;
	}
}

void CDlgSelEngVol::OnKeyDownDbgeng(short FAR* KeyCode, short Shift) 
{
/*	_variant_t key;//8/14
	CString str;
	//����������ģʽʱ������Ӧ�˺���
	//1:��ʾ��������״̬����û���޸�
	//2:��ʾ��������״̬�Ѿ��༭, ����û�����ӵ����ݿ���
	if( m_DBGeng.GetAddNewMode()==2||m_DBGeng.GetAddNewMode()==1)
	{
		return;
	}
	//VK_UP��VK_DOWNΪ���̵��������
	if(!(*KeyCode==VK_UP||*KeyCode==VK_DOWN))
	{
		return;
	}
	int i;
	_variant_t var = m_DataEng->GetBookmark();
	i = (int)var.dblVal;
	int count = m_DataEng->RecordCount;
	if(i>count)
	{	
		return;
	}

	if( *KeyCode == VK_UP )
	{
		if( i != 1 )
		{
			m_DataEng->MovePrevious();
			VARIANT FAR* LastRow; short LastCol;
			OnRowColChangeDbgeng(LastRow, LastCol);   //�ı䵱ǰ�л��С�
		}
	}
	if( *KeyCode == VK_DOWN )
	{
		if( i != count )
		{
			m_DataEng->MoveNext();
			VARIANT FAR* LastRow; short LastCol;
			OnRowColChangeDbgeng(LastRow, LastCol);   //�ı䵱ǰ�л��С�
		}
	}

*/
}

void CDlgSelEngVol::OnKeyUpDbgeng(short FAR* KeyCode, short Shift) 
{

		if(!(*KeyCode==VK_UP||*KeyCode==VK_DOWN))
		{
			return;
		}
		if(m_DBGeng.GetAddNewMode()==2||m_DBGeng.GetAddNewMode()==1)
		{
			m_DBGvlm.SetAllowAddNew(false);
		}

		if(!(m_DBGeng.GetAddNewMode()==1||m_DBGeng.GetAddNewMode()==2))
		
		{
			VARIANT FAR* LastRow; short LastCol;
			OnRowColChangeDbgeng(LastRow, LastCol);   
		}
	
}
