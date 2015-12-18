#pragma once

class AFX_CLASS_EXPORT ZSPDStartupManager
{
public:
	ZSPDStartupManager();
	ZSPDStartupManager(const CString& sModulePath, const CString& sArxLoad, const CString& sCADVersion );
	~ZSPDStartupManager(void);

private:
	BOOL RegZWSTProfile();
	BOOL GetCADCurProfile( const CString& sCADVersion,CString & sProfileName );
	BOOL SetCADProfile( const CString& sCADVersion, const CString & sProfileName );
	BOOL WriteStartupRx();
	BOOL RegSoftware();
	BOOL GetProfileRecordByReg( CString & sProfileRecord );
	BOOL GetCADVersionByReg( CString & sCADVersion );
	//�ⲿ�ӿ�
public:
	//��������Ŀ¼��ZWST.exeĿ¼��
	void SetModulePath( const CString& sPath );
	//���������汾
	void SetACADVersion( const CString& sVersion = _T("R18.0") );
	//����CAD
	BOOL Startup();
	BOOL Startup( const CString & sArxLoad, const CString& sPath, const CString& sCADVersion, const CString& sScrFile );

	BOOL clear();

	//ע����л�ȡCAD·��
	BOOL GetCADExePath( CString & sSetupPath, CString & sCADVersion );
	//ע����л�ȡARX·��
	BOOL GetAppPathByReg( CString & sAppPath );
	//ע����л�ȡ�˵��ļ���·��
	BOOL GetMnuPathByReg( CString & sMnuPath );
	//ע����л�ȡ���ݿ�·��
	BOOL GetDBPathByReg( CString & sDBPath );
	////ע����л�ȡ��װ·��
	BOOL GetSetupPathByReg( CString & sSetupPath );
	//ע����ҪItem
	BOOL RegItem( const CString & sKey, const CString & sRegValue , CString& sError);
	//��ȡItemֵ
	BOOL GetRegItem( const CString& sKey, CString & sRegVaule );
	//�������ã����ã�
	BOOL ResetProfileRecord();
	//ȡģ���ļ�Ŀ¼
	BOOL GetTempPathByReg( CString& sTempPath );
	//���ü���scr�ű��ļ�
	void setLoadScr( const CString& sScrFile);

	//��ȡ��ά��ͨ���·��
	CString getZwstPath(const CString&strFilePath,const CString& strPathName);

	//ֻ��ARXʹ��
    CString getZwstPath(const CString& strPathName);

private:
	CString m_sArxLoad;								//����ARX
	CString m_sModulePath;							//����Ŀ¼
	CString m_sACADVersion;							//����CAD�汾
	CString m_sScrFile;								//�����ű�
	CString m_strCurPofie;
};