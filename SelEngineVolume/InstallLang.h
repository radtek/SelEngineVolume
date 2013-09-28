#ifndef _INSTALLLNG
#define _INSTALLLNG

///using installed language  to set the running language

/*class InstallLanguage
{
public:	
	VOID InstallRunLanguage();//�����������Կ�
	VOID UnInstallRunLanguage();//ж���������Կ�
protected:
	VOID InstallDll(CString dllFile);//����dll
	VOID GetInstallLanguage();//��ȡ��װ����
	VOID UnInstallDll(CString dllFile);//ж��dll

};*/
/*installed or uninstall Current language*/

class Language
{
public:
	enum{INSTALL=0,UNINSTALL};
	Language(int option,CString="");
	Language();
	VOID Create(int option,CString="");
	VOID InstallRunLanguage(CString str);//�����������Կ�
	VOID UnInstallRunLanguage();//ж���������Կ�
	VOID SwitchLangResProcess();//�л��洢������Դ�ĵ�ǰ�ʹ洢���̡�

private:
	HINSTANCE m_hOldInstance;//ԭ����Դ�����
public:
	HINSTANCE m_hCurrentInstance;//��ǰ��Դ�����
	HINSTANCE m_hAddtion;///���
};
#endif