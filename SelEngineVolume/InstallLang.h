#ifndef _INSTALLLNG
#define _INSTALLLNG

///using installed language  to set the running language

/*class InstallLanguage
{
public:	
	VOID InstallRunLanguage();//加载运行语言库
	VOID UnInstallRunLanguage();//卸载运行语言库
protected:
	VOID InstallDll(CString dllFile);//加载dll
	VOID GetInstallLanguage();//获取安装语言
	VOID UnInstallDll(CString dllFile);//卸载dll

};*/
/*installed or uninstall Current language*/

class Language
{
public:
	enum{INSTALL=0,UNINSTALL};
	Language(int option,CString="");
	Language();
	VOID Create(int option,CString="");
	VOID InstallRunLanguage(CString str);//加载运行语言库
	VOID UnInstallRunLanguage();//卸载运行语言库
	VOID SwitchLangResProcess();//切换存储语言资源的当前和存储进程。

private:
	HINSTANCE m_hOldInstance;//原来资源句柄。
public:
	HINSTANCE m_hCurrentInstance;//当前资源句柄。
	HINSTANCE m_hAddtion;///添加
};
#endif