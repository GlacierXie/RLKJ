#include "StdAfx.h"
#include "ZSPDStartupManager.h"
#include <fstream>
#include "IcRegistry.h"

#define ACAD_CADCFGPATH		_T("Software\\Autodesk\\AutoCAD")

CStringA CStringToCStringA(const CString& str)
{
#ifdef UNICODE
	int iBufferSize = (str.GetLength() + 1) * 2;
	char *szBuffer = new char[iBufferSize];
	WideCharToMultiByte(CP_ACP, 0, str, -1, szBuffer, iBufferSize, NULL, NULL);
	CStringA sRet = szBuffer;
	delete[] szBuffer;
	return sRet;
#else
	return str;
#endif
}

ZSPDStartupManager::ZSPDStartupManager()
{
	m_sArxLoad = _T("rlkeBackEdge.arx");
}

ZSPDStartupManager::ZSPDStartupManager( const CString& sModulePath, const CString& sArxLoad, const CString& sCADVersion )
{
	SetModulePath(sModulePath);
	SetACADVersion(sCADVersion);
}

ZSPDStartupManager::~ZSPDStartupManager(void)
{
	
}

BOOL ZSPDStartupManager::RegZWSTProfile()
{

	CString sCADPath, sCADVersio, strCurPofie;
	GetCADExePath(sCADPath, sCADVersio);
	if (GetCADCurProfile(sCADVersio, strCurPofie))
	{
		;
	}
	CString sCADRegPath = CString(ACAD_CADCFGPATH) + _T("\\") + m_sACADVersion + _T("\\ACAD-B001:804\\Profiles\\") + strCurPofie + _T("\\General");

	IcRegistry regUser( HKEY_CURRENT_USER);
	HKEY hRegApp = regUser.OpenSubKey(sCADRegPath, false);
	if(hRegApp == NULL)
		return FALSE;

	IcRegistry regAppPath(hRegApp);
	CString sCur;
	regAppPath.Read(_T("ACAD"), sCur);
	if (sCur.Find(_T("rlkebackedge")) >= 0)
		return TRUE;
// 	CString strTmp = _T("UserDataCache\\zh-CN\\Support");
// 	strTmp.MakeLower();
// 	if (sCur.Find(strTmp) > 0)
// 	{
// 		int nLength = sCur.ReverseFind(';');
// 		sCur = sCur.Left(nLength);
// 		nLength = sCur.ReverseFind(';');
// 		sCur = sCur.Left(nLength + 1);
// 	}
	CString strTmp = m_sModulePath;
	sCur = sCur + strTmp.MakeLower();
	sCur = sCur + _T(";");
	regAppPath.Write(_T("ACAD"), sCur);
	//regAppPath.Write(_T("ShowFullPathInTitle"), 1);

// 	sCADRegPath = CString(ACAD_CADCFGPATH) + _T("\\") + m_sACADVersion + _T("\\ACAD-B001:804\\Profiles\\") + m_strCurPofie + _T("\\General Configuration");
// 	IcRegistry regUser1(HKEY_CURRENT_USER);
// 	HKEY hRegApp1 = regUser1.OpenSubKey(sCADRegPath, false);
// 	if (hRegApp1 == NULL)
// 		return FALSE;

// 	IcRegistry regAppPath1(hRegApp1);
// 	regAppPath1.Write(_T("MenuFile"), m_sModulePath + _T("\\acad"));

	return TRUE;
}

BOOL ZSPDStartupManager::clear()
{
	CString sCurProfile = _T("");
	CString sCADPath;
	CString sCADVersion;
	if (GetCADExePath(sCADPath, sCADVersion))
	{
		if (GetCADCurProfile(sCADVersion, sCurProfile))
		{

		}
	}
	
	int nLength = sCADPath.ReverseFind('\\');
	sCADPath = sCADPath.Left(nLength);
	CString sCADPathTmp = sCADPath;
	sCADPathTmp.MakeLower();

	sCADPath += _T("\\UserDataCache\\zh-CN\\Support\\acad");

	CString sCADRegPath = CString(ACAD_CADCFGPATH) + _T("\\") + m_sACADVersion + _T("\\ACAD-B001:804\\Profiles\\") + sCurProfile + _T("\\General Configuration");
	IcRegistry regUser1(HKEY_CURRENT_USER);
	HKEY hRegApp1 = regUser1.OpenSubKey(sCADRegPath, false);
	if (hRegApp1 == NULL)
		return FALSE;

	IcRegistry regAppPath1(hRegApp1);
	regAppPath1.Write(_T("MenuFile"), sCADPath);

	 sCADRegPath = CString(ACAD_CADCFGPATH) + _T("\\") + m_sACADVersion + _T("\\ACAD-B001:804\\Profiles\\") + sCurProfile + _T("\\General");
	IcRegistry regUser(HKEY_CURRENT_USER);
	HKEY hRegApp = regUser.OpenSubKey(sCADRegPath, false);
	if (hRegApp == NULL)
		return FALSE;


	IcRegistry regAppPath(hRegApp);
	CString sCur;
	regAppPath.Read(_T("ACAD"), sCur);
	if (sCur.Find(_T("rlkebackedge")) < 0)
		return FALSE;
	nLength = sCur.ReverseFind(';');
	sCur = sCur.Left(nLength);
	nLength = sCur.ReverseFind(';');
	sCur = sCur.Left(nLength+1);
// 	CString strTmp = _T("UserDataCache\\zh-CN\\Support");
// 	strTmp.MakeLower();
// 	if (sCur.Find(strTmp) < 0)
// 		sCur = sCur + sCADPathTmp + strTmp + _T(";");
	
	regAppPath.Write(_T("ACAD"), sCur);

	CFile::Remove(m_sModulePath + _T("\\acad.rx"));

	return TRUE;
}

BOOL ZSPDStartupManager::GetAppPathByReg( CString & sAppPath )
{
	sAppPath = m_sModulePath;
	return TRUE;
}

BOOL ZSPDStartupManager::GetMnuPathByReg( CString & sMnuPath )
{
	return GetAppPathByReg(sMnuPath);
}

BOOL ZSPDStartupManager::GetDBPathByReg( CString & sDBPath )
{
	return GetAppPathByReg(sDBPath);
}

BOOL ZSPDStartupManager::GetSetupPathByReg( CString & sSetupPath )
{
	return GetAppPathByReg(sSetupPath);
}

BOOL ZSPDStartupManager::GetCADExePath( CString & sSetupPath, CString & sCADVersion )
{
	if(m_sACADVersion.IsEmpty())
		return FALSE;

	if(m_sACADVersion.CompareNoCase(_T("R18.0"))!=0 && m_sACADVersion.CompareNoCase(_T("R19.0"))!=0)
		return FALSE;

	IcRegistry reg(HKEY_LOCAL_MACHINE);
	HKEY hKey = reg.OpenSubKey(_T("SOFTWARE\\Autodesk\\AutoCAD\\") + m_sACADVersion+_T("\\ACAD-B001:804"));//read
	if (NULL == hKey)
		return FALSE;
	
	IcRegistry subReg(hKey);
	if (!subReg.Read(_T("AcadLocation"), sSetupPath) || !subReg.Read(_T("ProductName"), sCADVersion))
	     return FALSE;
	sSetupPath = sSetupPath + _T("\\acad.exe");
	sCADVersion = _T("ACAD-B001:804");
	return TRUE;
}

BOOL ZSPDStartupManager::GetCADCurProfile( const CString& sCADVersion,CString & sProfileName )
{
	IcRegistry regUser( HKEY_CURRENT_USER);
	HKEY hRegApp = regUser.OpenSubKey(CString(ACAD_CADCFGPATH) + _T("\\R19.0\\") + sCADVersion + _T("\\profiles"), true);
	if(hRegApp == NULL)
		return FALSE;

	IcRegistry regAppPath(hRegApp);
	return regAppPath.Read(_T(""), sProfileName);
}

BOOL ZSPDStartupManager::SetCADProfile( const CString& sCADVersion, const CString & sProfileName )
{
	IcRegistry regUser( HKEY_CURRENT_USER);
	HKEY hRegApp = regUser.OpenSubKey(CString(ACAD_CADCFGPATH) + _T("\\R19.0\\") + sCADVersion + _T("\\profiles"), false);
	if(hRegApp == NULL)
		return FALSE;

	IcRegistry regAppPath(hRegApp);
	return regAppPath.Write(_T(""), sProfileName);
}

void ZSPDStartupManager::setLoadScr( const CString& sScrFile)
{
	m_sScrFile = sScrFile;
}

void ZSPDStartupManager::SetModulePath( const CString& sPath )
{
	m_sModulePath = sPath;
	m_sModulePath.TrimRight('\\');
}

BOOL ZSPDStartupManager::Startup()
{
	//获取ACAD安装路径

	CString sACADVersion;
	CString sACADPath;
	if (!GetCADExePath(sACADPath, sACADVersion))
	{
		AfxMessageBox(CString(_T("CAD ")) + m_sACADVersion + _T(" 安装不正确！"));
		return FALSE;
	}
	if (GetCADCurProfile(sACADVersion, m_strCurPofie))
	{
		;
	}
	if(!RegSoftware())
	{
		AfxMessageBox(_T("注册失败！"));
		return FALSE;
	}

	if(!WriteStartupRx())
	{
		AfxMessageBox(_T("设置启动文件失败！"));
		return FALSE;
	}

	if(!RegZWSTProfile())
	{
		AfxMessageBox(_T("读取配置文件失败！"));
		return FALSE;
	}

// 	//获取配置数据
// 
// 	CString sProfileOld;
// 	GetCADCurProfile(sACADVersion, sProfileOld );
// 	SetCADProfile(sACADVersion, sProfileOld);

	return TRUE;
}


BOOL ZSPDStartupManager::Startup( const CString& sArxLoad, const CString& sPath, const CString& sCADVersion, const CString& sScrFile )
{
	SetModulePath(sPath);
	SetACADVersion(sCADVersion);
	setLoadScr(sScrFile);

	return Startup();
}


void ZSPDStartupManager::SetACADVersion( const CString& sVersion /*= _T("R18.0") */ )
{
	m_sACADVersion = sVersion;
}

BOOL ZSPDStartupManager::WriteStartupRx()
{
	CString sArxPath;
	GetAppPathByReg( sArxPath );

	CFileFind finder;
	if (m_sArxLoad.IsEmpty())
	{
		CFile::Remove(sArxPath + _T("\\acad.rx"));
		return TRUE;
	}

	if(!finder.FindFile(sArxPath + _T("\\") + m_sArxLoad))
		return FALSE;

	if(finder.FindFile( sArxPath + _T("\\acad.rx")))
		CFile::Remove(sArxPath + _T("\\acad.rx"));

	/*jd 2015 3 19 */
	finder.Close();

	CString sRxFileName = sArxPath + _T("\\acad.rx");
	std::ofstream rxfile(CStringToCStringA(sRxFileName));
	CString sTemp;
	if(sArxPath.Right(1).CompareNoCase(_T("\\"))==0)
		sTemp = m_sArxLoad + _T("\n");
	else
		sTemp = CString(_T("\\")) + m_sArxLoad + _T("\n");
	rxfile.clear();
	rxfile<<CStringToCStringA(sArxPath) << CStringToCStringA(sTemp);
	rxfile.close();

	return TRUE;
}

BOOL ZSPDStartupManager::RegSoftware()
{
	//解析路径

	if(m_sModulePath.IsEmpty())
		return FALSE;

	CString sModulePath = m_sModulePath;
	CFileFind finder;

	//程序路径
	if(m_sArxLoad.IsEmpty())
	{
		AfxMessageBox(_T("软件安装不正确，请检查更新！"));
		return FALSE;
	}

	CString sArxPath = _T("");
	while(!finder.FindFile(sModulePath + _T("\\") + m_sArxLoad))
	{
		int nFind = sModulePath.ReverseFind('\\');
		if(nFind==-1)
		{
			sModulePath = _T("");
			break;
		}

		sModulePath = sModulePath.Left( nFind );
	}

	sArxPath = sModulePath;
	if( sModulePath.IsEmpty() )
	{
		AfxMessageBox(_T("软件安装不正确，请检查更新！"));
		return FALSE;
	}
	sModulePath = m_sModulePath;

	return TRUE;
}


BOOL ZSPDStartupManager::ResetProfileRecord()
{
	CString sProfileRec = m_strCurPofie;
	CString sCADPath;
	CString sCADVersion;
	GetCADExePath(sCADPath, sCADVersion);
	SetCADProfile(sCADVersion, sProfileRec);
	return TRUE;
}


