#include"stdafx.h"
#include "InstallLang.h"


Language::Language ( )
{
 m_hCurrentInstance = NULL;
 m_hOldInstance = NULL;
}

Language :: Language(int option,CString str)
{
	m_hCurrentInstance=m_hOldInstance=NULL;
 Create(option,str);
}

VOID Language::InstallRunLanguage(CString str)
{//安装和设置资源
	 HINSTANCE hInst = LoadLibrary (str);//
	 m_hOldInstance = AfxGetResourceHandle ();
	 AfxSetResourceHandle (hInst);
	 m_hCurrentInstance = hInst;
}


VOID Language::Create(int option,CString str/* ="" */)
{
	
	if(option == INSTALL&&str!="")InstallRunLanguage(str);
    if(option == UNINSTALL)UnInstallRunLanguage();
}



VOID Language::UnInstallRunLanguage()///设置成保存的资源进程
{//卸载资源
 // ASSERT(m_hOldInstance!=NULL);
  if(m_hOldInstance==NULL)return;
  HINSTANCE hInstal;
  hInstal=AfxGetResourceHandle();
  ASSERT(FreeLibrary(hInstal));
  AfxSetResourceHandle(m_hOldInstance);
  m_hCurrentInstance=m_hOldInstance;
  m_hOldInstance=NULL;  
}


VOID Language::SwitchLangResProcess()
{//切换资源
	if(m_hCurrentInstance==NULL||m_hOldInstance==NULL)
	{
		ASSERT(FALSE);
		return;
	}
	else
	{
		m_hCurrentInstance = m_hOldInstance;
		AfxSetResourceHandle(m_hOldInstance);
		m_hOldInstance=m_hCurrentInstance;
	}
}
