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
	//外部接口
public:
	//设置启动目录（ZWST.exe目录）
	void SetModulePath( const CString& sPath );
	//设置启动版本
	void SetACADVersion( const CString& sVersion = _T("R18.0") );
	//启动CAD
	BOOL Startup();
	BOOL Startup( const CString & sArxLoad, const CString& sPath, const CString& sCADVersion, const CString& sScrFile );

	BOOL clear();

	//注册表中获取CAD路径
	BOOL GetCADExePath( CString & sSetupPath, CString & sCADVersion );
	//注册表中获取ARX路径
	BOOL GetAppPathByReg( CString & sAppPath );
	//注册表中获取菜单文件夹路径
	BOOL GetMnuPathByReg( CString & sMnuPath );
	//注册表中获取数据库路径
	BOOL GetDBPathByReg( CString & sDBPath );
	////注册表中获取安装路径
	BOOL GetSetupPathByReg( CString & sSetupPath );
	//注册需要Item
	BOOL RegItem( const CString & sKey, const CString & sRegValue , CString& sError);
	//获取Item值
	BOOL GetRegItem( const CString& sKey, CString & sRegVaule );
	//重置配置（慎用）
	BOOL ResetProfileRecord();
	//取模板文件目录
	BOOL GetTempPathByReg( CString& sTempPath );
	//设置加载scr脚本文件
	void setLoadScr( const CString& sScrFile);

	//获取中维数通相关路径
	CString getZwstPath(const CString&strFilePath,const CString& strPathName);

	//只供ARX使用
    CString getZwstPath(const CString& strPathName);

private:
	CString m_sArxLoad;								//启动ARX
	CString m_sModulePath;							//启动目录
	CString m_sACADVersion;							//启动CAD版本
	CString m_sScrFile;								//启动脚本
	CString m_strCurPofie;
};