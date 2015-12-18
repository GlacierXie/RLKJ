#include "StdAfx.h"
#include "IcRegistry.h" 

#ifdef _WIN64
BOOL g_bOpenRegistryX64=TRUE;//为TRUE的情况：32位app在64位操作系统下，需要访问64位注册表
#else
BOOL g_bOpenRegistryX64=FALSE;//为TRUE的情况：32位app在64位操作系统下，需要访问64位注册表
#endif


IcRegistry::IcRegistry(HKEY hKeyRoot)
 : m_hKey(hKeyRoot)
{
	
}

IcRegistry::~IcRegistry(void)
{
	Close(); 
}

HKEY IcRegistry::OpenSubKey( LPCTSTR pszPath ,bool bReadOnly )
{
	ASSERT (m_hKey); 
	ASSERT (pszPath); 

	if ( NULL == m_hKey )
	{
		return NULL;
	}

	HKEY hRet = NULL;
	REGSAM access=bReadOnly?KEY_READ:KEY_READ|KEY_WRITE;
	if(g_bOpenRegistryX64)
		access|=KEY_WOW64_64KEY;

	LONG ReturnValue = RegOpenKeyEx (m_hKey, pszPath, 0L, access, &hRet); 

	m_Info.lMessage = ReturnValue; 
	m_Info.dwSize = 0L; 
	m_Info.dwType = 0L;

	if(ReturnValue == ERROR_SUCCESS)
	{
		return hRet; 
	} 

	return NULL; 
}


BOOL IcRegistry::VerifyKey (LPCTSTR pszPath) 
{ 
	ASSERT (m_hKey); 

	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	HKEY hTmpKey = NULL;
	REGSAM access=KEY_READ;
	if(g_bOpenRegistryX64)
		access|=KEY_WOW64_64KEY;

	LONG ReturnValue = RegOpenKeyEx (m_hKey, pszPath, 0L, access, &hTmpKey); 

	m_Info.lMessage = ReturnValue; 
	m_Info.dwSize = 0L; 
	m_Info.dwType = 0L; 

	if(ReturnValue == ERROR_SUCCESS) 
	{
		RegCloseKey (hTmpKey); 
		return TRUE; 
	}

	return FALSE; 
} 

BOOL IcRegistry::VerifyValue (LPCTSTR pszValue) 
{ 
	ASSERT(m_hKey); 

	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	LONG lReturn = RegQueryValueEx(m_hKey, pszValue, NULL, 
		NULL, NULL, NULL); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = 0L; 
	m_Info.dwType = 0L; 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::VerifyValue (LPCTSTR pszValue, DWORD dwType )
{
	ASSERT(m_hKey); 

	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	DWORD dwTypeRet;
	LONG lReturn = RegQueryValueEx(m_hKey, pszValue, NULL, 
		&dwTypeRet, NULL, NULL);

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = 0L; 
	m_Info.dwType = 0L; 

	if(lReturn == ERROR_SUCCESS && dwType==dwTypeRet) 
		return TRUE; 

	return FALSE; 
}

HKEY IcRegistry::CreateKey (LPCTSTR pszPath)
{
	ASSERT(m_hKey); 

	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	DWORD dw; 
	HKEY hTmpKey = NULL;

	REGSAM access=KEY_READ|KEY_WRITE;
	if(g_bOpenRegistryX64)
		access|=KEY_WOW64_64KEY;

	LONG ReturnValue = RegCreateKeyEx (m_hKey, pszPath, 0L, NULL, 
		/*REG_OPTION_VOLATILE*/REG_OPTION_NON_VOLATILE, //modified by whl on 101114
		access, NULL, 
		&hTmpKey, &dw); 
	if (ReturnValue != ERROR_SUCCESS)
	{
		return FALSE;
	}
	m_Info.lMessage = ReturnValue; 
	m_Info.dwSize = 0L; 
	m_Info.dwType = 0L; 

	if(ReturnValue == ERROR_SUCCESS) 
		return hTmpKey; 

	return FALSE; 
}

void IcRegistry::Close() 
{  
	if (NULL!=m_hKey
		&& HKEY_CLASSES_ROOT!=m_hKey
		&& HKEY_CURRENT_USER!=m_hKey
		&& HKEY_LOCAL_MACHINE!=m_hKey
		&& HKEY_USERS!=m_hKey
		) 
	{ 
		RegCloseKey (m_hKey); 
		m_hKey = NULL; 
	} 
} 

BOOL IcRegistry::Write (LPCTSTR pszKey, int iVal) 
{ 
	DWORD dwValue; 

	ASSERT(m_hKey); 
	ASSERT(pszKey); 

	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	dwValue = (DWORD)iVal; 
	LONG ReturnValue = RegSetValueEx (m_hKey, pszKey, 0L, REG_DWORD, 
		(CONST BYTE*) &dwValue, sizeof(DWORD)); 

	m_Info.lMessage = ReturnValue; 
	m_Info.dwSize = sizeof(DWORD); 
	m_Info.dwType = REG_DWORD; 

	if(ReturnValue == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Write (LPCTSTR pszKey, DWORD dwVal) 
{ 
	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	ASSERT(m_hKey); 
	ASSERT(pszKey); 

	return RegSetValueEx (m_hKey, pszKey, 0L, REG_DWORD, 
		(CONST BYTE*) &dwVal, sizeof(DWORD)); 
} 

BOOL IcRegistry::Write (LPCTSTR pszKey, LPCTSTR pszData) 
{ 
	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	ASSERT(pszData); 
	ASSERT(AfxIsValidAddress(pszData, _tcslen(pszData), FALSE)); 

	LONG ReturnValue = RegSetValueEx (m_hKey, pszKey, 0L, REG_SZ, 
		(CONST BYTE*) pszData, ((DWORD)_tcslen(pszData) + 1)*sizeof(TCHAR) ); 

	m_Info.lMessage = ReturnValue; 
	m_Info.dwSize = (DWORD)_tcslen(pszData) + 1; 
	m_Info.dwType = REG_SZ; 

	if(ReturnValue == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Write (LPCTSTR pszKey, CStringList& scStringList) 
{ 
	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	CMemFile file(byData, iMaxChars, 16); 
	CArchive ar(&file, CArchive::store); 
	ASSERT(scStringList.IsSerializable()); 
	scStringList.Serialize(ar); 
	ar.Close(); 
	const ULONGLONG dwLen = file.GetLength(); 
	ASSERT(dwLen < iMaxChars); 
	LONG lReturn = RegSetValueEx(m_hKey, pszKey, 0, REG_BINARY, 
		file.Detach(), (DWORD)dwLen); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = (DWORD)dwLen; 
	m_Info.dwType = REG_BINARY; 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
}  

BOOL IcRegistry::Write (LPCTSTR pszKey, CByteArray& bcArray) 
{ 
	if ( NULL == m_hKey )
	{
		return FALSE;
	}

	ASSERT(m_hKey); 
	ASSERT(pszKey); 

	const int iMaxChars = 4096; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	CMemFile file(byData, iMaxChars, 16); 
	CArchive ar(&file, CArchive::store); 
	ASSERT(bcArray.IsSerializable()); 
	bcArray.Serialize(ar); 
	ar.Close(); 
	const ULONGLONG dwLen = file.GetLength(); 
	ASSERT(dwLen < iMaxChars); 
	LONG lReturn = RegSetValueEx(m_hKey, pszKey, 0, REG_BINARY, 
		file.Detach(), (DWORD)dwLen); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = (DWORD)dwLen; 
	m_Info.dwType = REG_BINARY; 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Write (LPCTSTR pszKey, CDWordArray& dwcArray) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	CMemFile file(byData, iMaxChars, 16); 
	CArchive ar(&file, CArchive::store); 
	ASSERT(dwcArray.IsSerializable()); 
	dwcArray.Serialize(ar); 
	ar.Close(); 
	const ULONGLONG dwLen = file.GetLength(); 
	ASSERT(dwLen < iMaxChars); 
	LONG lReturn = RegSetValueEx(m_hKey, pszKey, 0, REG_BINARY, 
		file.Detach(), (DWORD)dwLen); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = (DWORD)dwLen; 
	m_Info.dwType = REG_BINARY; 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Write (LPCTSTR pszKey, CWordArray& wcArray) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	CMemFile file(byData, iMaxChars, 16); 
	CArchive ar(&file, CArchive::store); 
	ASSERT(wcArray.IsSerializable()); 
	wcArray.Serialize(ar); 
	ar.Close(); 
	const ULONGLONG dwLen = file.GetLength(); 
	ASSERT(dwLen < iMaxChars); 
	LONG lReturn = RegSetValueEx(m_hKey, pszKey, 0, REG_BINARY, 
		file.Detach(), (WORD)dwLen); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = (WORD)dwLen; 
	m_Info.dwType = REG_BINARY; 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Write (LPCTSTR pszKey, CStringArray& scArray) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	CMemFile file(byData, iMaxChars, 16); 
	CArchive ar(&file, CArchive::store); 
	ASSERT(scArray.IsSerializable()); 
	scArray.Serialize(ar); 
	ar.Close(); 
	const ULONGLONG dwLen = file.GetLength(); 
	ASSERT(dwLen < iMaxChars); 
	LONG lReturn = RegSetValueEx(m_hKey, pszKey, 0, REG_BINARY, 
		file.Detach(), (DWORD)dwLen); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = (DWORD)dwLen; 
	m_Info.dwType = REG_BINARY; 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Write(LPCTSTR pszKey, LPCRECT rcRect) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 30; 
	CDWordArray dwcArray; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	dwcArray.SetSize(5); 
	dwcArray.SetAt(0, rcRect->top); 
	dwcArray.SetAt(1, rcRect->bottom); 
	dwcArray.SetAt(2, rcRect->left); 
	dwcArray.SetAt(3, rcRect->right); 

	CMemFile file(byData, iMaxChars, 16); 
	CArchive ar(&file, CArchive::store); 
	ASSERT(dwcArray.IsSerializable()); 
	dwcArray.Serialize(ar); 
	ar.Close(); 
	const ULONGLONG dwLen = file.GetLength(); 
	ASSERT(dwLen < iMaxChars); 
	LONG lReturn = RegSetValueEx(m_hKey, pszKey, 0, REG_BINARY, 
		file.Detach(), (DWORD)dwLen); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = (DWORD)dwLen; 
	m_Info.dwType = REG_RECT; 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Write(LPCTSTR pszKey, LPPOINT& lpPoint) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 20; 
	CDWordArray dwcArray; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	dwcArray.SetSize(5); 
	dwcArray.SetAt(0, lpPoint->x); 
	dwcArray.SetAt(1, lpPoint->y); 

	CMemFile file(byData, iMaxChars, 16); 
	CArchive ar(&file, CArchive::store); 
	ASSERT(dwcArray.IsSerializable()); 
	dwcArray.Serialize(ar); 
	ar.Close(); 
	const ULONGLONG dwLen = file.GetLength(); 
	ASSERT(dwLen < iMaxChars); 
	LONG lReturn = RegSetValueEx(m_hKey, pszKey, 0, REG_BINARY, 
		file.Detach(), (DWORD)dwLen); 

	m_Info.lMessage = lReturn; 
	m_Info.dwSize = (DWORD)dwLen; 
	m_Info.dwType = REG_POINT; 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read(LPCTSTR pszKey, int& iVal) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 

	DWORD dwType; 
	DWORD dwSize = sizeof (DWORD); 
	DWORD dwDest; 

	LONG lReturn = RegQueryValueEx (m_hKey, (LPCTSTR) pszKey, NULL, 
		&dwType, (BYTE *) &dwDest, &dwSize); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwSize; 

	if(lReturn == ERROR_SUCCESS) 
	{ 
		iVal = (int)dwDest; 
		return TRUE; 
	} 

	return FALSE; 
} 

BOOL IcRegistry::Read (LPCTSTR pszKey, DWORD& dwVal) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 

	DWORD dwType; 
	DWORD dwSize = sizeof (DWORD); 
	DWORD dwDest; 

	LONG lReturn = RegQueryValueEx (m_hKey, (LPCTSTR) pszKey, NULL, 
		&dwType, (BYTE *) &dwDest, &dwSize); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwSize; 

	if(lReturn == ERROR_SUCCESS) 
	{ 
		dwVal = dwDest; 
		return TRUE; 
	} 

	return FALSE; 
} 

BOOL IcRegistry::Read (LPCTSTR pszKey, CString& sVal) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 

	DWORD dwType; 
	DWORD dwSize = 1000; 
	char szString[1024]; 

	memset(szString, 0, sizeof(szString)); 

	LONG lReturn = RegQueryValueEx (m_hKey, (LPCTSTR) pszKey, NULL, 
		&dwType, (BYTE *) szString, &dwSize); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwSize; 

	if(lReturn == ERROR_SUCCESS) 
	{ 
		sVal.Format( _T("%s"), szString );
		//sVal = szString; 
		return TRUE; 
	} 

	return FALSE; 
} 

BOOL IcRegistry::Read (LPCTSTR pszKey, CStringList& scStringList) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	DWORD dwType; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	LONG lReturn = RegQueryValueEx(m_hKey, pszKey, NULL, &dwType, 
		byData, &dwData); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwData; 

	if(lReturn == ERROR_SUCCESS && dwType == REG_BINARY) 
	{ 
		ASSERT(dwData < iMaxChars); 
		CMemFile file(byData, dwData); 
		CArchive ar(&file, CArchive::load); 
		ar.m_bForceFlat = FALSE; 
		ASSERT(ar.IsLoading()); 
		ASSERT(scStringList.IsSerializable()); 
		scStringList.RemoveAll(); 
		scStringList.Serialize(ar); 
		ar.Close(); 
		file.Close(); 
	} 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read (LPCTSTR pszKey, CByteArray& bcArray) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	int OldSize = (int)bcArray.GetSize(); 
	DWORD dwType; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	LONG lReturn = RegQueryValueEx(m_hKey, pszKey, NULL, &dwType, 
		byData, &dwData); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwData; 

	if(lReturn == ERROR_SUCCESS && dwType == REG_BINARY) 
	{ 
		ASSERT(dwData < iMaxChars); 
		CMemFile file(byData, dwData); 
		CArchive ar(&file, CArchive::load); 
		ar.m_bForceFlat = FALSE; 
		ASSERT(ar.IsLoading()); 
		ASSERT(bcArray.IsSerializable()); 
		bcArray.RemoveAll(); 
		bcArray.SetSize(10); 
		bcArray.Serialize(ar); 
		bcArray.SetSize(OldSize); 
		ar.Close(); 
		file.Close(); 
	} 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read (LPCTSTR pszKey, CDWordArray& dwcArray) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	int OldSize = (int)dwcArray.GetSize(); 
	DWORD dwType; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	LONG lReturn = RegQueryValueEx(m_hKey, pszKey, NULL, &dwType, 
		byData, &dwData); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwData; 

	if(lReturn == ERROR_SUCCESS && dwType == REG_BINARY) 
	{ 
		ASSERT(dwData < iMaxChars); 
		CMemFile file(byData, dwData); 
		CArchive ar(&file, CArchive::load); 
		ar.m_bForceFlat = FALSE; 
		ASSERT(ar.IsLoading()); 
		ASSERT(dwcArray.IsSerializable()); 
		dwcArray.RemoveAll(); 
		dwcArray.SetSize(10); 
		dwcArray.Serialize(ar); 
		dwcArray.SetSize(OldSize); 
		ar.Close(); 
		file.Close(); 
	} 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read (LPCTSTR pszKey, CWordArray& wcArray) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	int OldSize = (int)wcArray.GetSize(); 
	DWORD dwType; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	LONG lReturn = RegQueryValueEx(m_hKey, pszKey, NULL, &dwType, 
		byData, &dwData); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwData; 

	if(lReturn == ERROR_SUCCESS && dwType == REG_BINARY) 
	{ 
		ASSERT(dwData < iMaxChars); 
		CMemFile file(byData, dwData); 
		CArchive ar(&file, CArchive::load); 
		ar.m_bForceFlat = FALSE; 
		ASSERT(ar.IsLoading()); 
		ASSERT(wcArray.IsSerializable()); 
		wcArray.RemoveAll(); 
		wcArray.SetSize(10); 
		wcArray.Serialize(ar); 
		wcArray.SetSize(OldSize); 
		ar.Close(); 
		file.Close(); 
	} 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read (LPCTSTR pszKey, CStringArray& scArray) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 4096; 
	int OldSize = (int)scArray.GetSize(); 
	DWORD dwType; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	LONG lReturn = RegQueryValueEx(m_hKey, pszKey, NULL, &dwType, 
		byData, &dwData); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = dwType; 
	m_Info.dwSize = dwData; 

	if(lReturn == ERROR_SUCCESS && dwType == REG_BINARY) 
	{ 
		ASSERT(dwData < iMaxChars); 
		CMemFile file(byData, dwData); 
		CArchive ar(&file, CArchive::load); 
		ar.m_bForceFlat = FALSE; 
		ASSERT(ar.IsLoading()); 
		ASSERT(scArray.IsSerializable()); 
		scArray.RemoveAll(); 
		scArray.SetSize(10); 
		scArray.Serialize(ar); 
		scArray.SetSize(OldSize); 
		ar.Close(); 
		file.Close(); 
	} 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read(LPCTSTR pszKey, LPRECT& rcRect) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 30; 
	CDWordArray dwcArray; 
	DWORD dwType; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	LONG lReturn = RegQueryValueEx(m_hKey, pszKey, NULL, &dwType, 
		byData, &dwData); 

	if(lReturn == ERROR_SUCCESS && dwType == REG_BINARY) 
	{ 
		ASSERT(dwData < iMaxChars); 
		CMemFile file(byData, dwData); 
		CArchive ar(&file, CArchive::load); 
		ar.m_bForceFlat = FALSE; 
		ASSERT(ar.IsLoading()); 
		ASSERT(dwcArray.IsSerializable()); 
		dwcArray.RemoveAll(); 
		dwcArray.SetSize(5); 
		dwcArray.Serialize(ar); 
		ar.Close(); 
		file.Close(); 
		rcRect->top = dwcArray.GetAt(0); 
		rcRect->bottom = dwcArray.GetAt(1); 
		rcRect->left = dwcArray.GetAt(2); 
		rcRect->right = dwcArray.GetAt(3); 
	} 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = REG_RECT; 
	m_Info.dwSize = sizeof(RECT); 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read(LPCTSTR pszKey, LPPOINT& lpPoint) 
{ 
	ASSERT(m_hKey); 
	ASSERT(pszKey); 
	const int iMaxChars = 20; 
	CDWordArray dwcArray; 
	DWORD dwType; 
	DWORD dwData = iMaxChars; 
	BYTE* byData = (BYTE*)::calloc(iMaxChars, sizeof(TCHAR)); 
	ASSERT(byData); 

	LONG lReturn = RegQueryValueEx(m_hKey, pszKey, NULL, &dwType, 
		byData, &dwData); 

	if(lReturn == ERROR_SUCCESS && dwType == REG_BINARY) 
	{ 
		ASSERT(dwData < iMaxChars); 
		CMemFile file(byData, dwData); 
		CArchive ar(&file, CArchive::load); 
		ar.m_bForceFlat = FALSE; 
		ASSERT(ar.IsLoading()); 
		ASSERT(dwcArray.IsSerializable()); 
		dwcArray.RemoveAll(); 
		dwcArray.SetSize(5); 
		dwcArray.Serialize(ar); 
		ar.Close(); 
		file.Close(); 
		lpPoint->x = dwcArray.GetAt(0); 
		lpPoint->y = dwcArray.GetAt(1); 
	} 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = REG_POINT; 
	m_Info.dwSize = sizeof(POINT); 

	if(byData) 
	{ 
		free(byData); 
		byData = NULL; 
	} 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::Read( LPCTSTR lpszValueName, LPVOID lpReturnBuffer, int nSize ) 
{ 

	if( m_hKey == NULL ) 
		return( FALSE ); 

	DWORD dwSize = (DWORD) nSize; 
	BOOL bRet = ( ::RegQueryValueEx( m_hKey, lpszValueName, NULL, NULL, (unsigned char *) lpReturnBuffer, &dwSize ) == ERROR_SUCCESS ); 

	m_dwLastError = GetLastError(); 

	return( bRet ); 

} 

BOOL IcRegistry::ReadDWORD( LPCTSTR lpszValueName, DWORD *pdwData, DWORD *pdwLastError ) 
{ 

	if( m_hKey == NULL ) 
		return( FALSE ); 

	BOOL bRet = Read( lpszValueName, pdwData, sizeof( DWORD ) ); 

	if( pdwLastError != NULL ) 
		*pdwLastError = m_dwLastError; 

	return( bRet ); 

} 

BOOL IcRegistry::DeleteValue (LPCTSTR pszValue) 
{ 
	ASSERT(m_hKey); 
	LONG lReturn = RegDeleteValue(m_hKey, pszValue); 


	m_Info.lMessage = lReturn; 
	m_Info.dwType = 0L; 
	m_Info.dwSize = 0L; 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
} 

BOOL IcRegistry::ReadString( const TCHAR *lpszValueName, LPVOID lpReturnBuffer, int nSize, DWORD *pdwLastError ) 
{ 

	if( m_hKey == NULL ) 
		return( FALSE ); 

	char *lpWork = (char *) lpReturnBuffer; 
	lpWork[0] = 0; 
	BOOL bRet = Read( lpszValueName, lpReturnBuffer, nSize ); 

	if( pdwLastError != NULL ) 
		*pdwLastError = m_dwLastError; 

	return( bRet ); 

} 

BOOL IcRegistry::DeleteValueKey (HKEY hKeyRoot, LPCTSTR pszPath) 
{ 
	ASSERT(pszPath); 
	ASSERT(hKeyRoot); 

	LONG lReturn = RegDeleteKey(hKeyRoot, pszPath); 

	m_Info.lMessage = lReturn; 
	m_Info.dwType = 0L; 
	m_Info.dwSize = 0L; 

	if(lReturn == ERROR_SUCCESS) 
		return TRUE; 

	return FALSE; 
}

BOOL IcRegistry::QuerySubKeyNameList( CStringArray& scArray ) //获取子项名称列表
{
	ASSERT(m_hKey); 

	if ( NULL == m_hKey )
	{
		return FALSE;
	}
	
	TCHAR    achKey[MAX_KEY_LENGTH];   // buffer for subkey name
	DWORD    cbName;                   // size of name string 
	TCHAR    achClass[MAX_PATH] = TEXT("");  // buffer for class name 
	DWORD    cchClassName = MAX_PATH;  // size of class string 
	DWORD    cSubKeys=0;               // number of subkeys 
	DWORD    cbMaxSubKey;              // longest subkey size 
	DWORD    cchMaxClass;              // longest class string 
	DWORD    cValues;              // number of values for key 
	DWORD    cchMaxValue;          // longest value name 
	DWORD    cbMaxValueData;       // longest value data 
	DWORD    cbSecurityDescriptor; // size of security descriptor 
	FILETIME ftLastWriteTime;      // last write time 

	DWORD i, retCode; 

	DWORD cchValue = MAX_VALUE_NAME; 

	CString sTemp;
	// Get the class name and the value count. 
	retCode = RegQueryInfoKey(
		m_hKey,                    // key handle 
		achClass,                // buffer for class name 
		&cchClassName,           // size of class string 
		NULL,                    // reserved 
		&cSubKeys,               // number of subkeys 
		&cbMaxSubKey,            // longest subkey size 
		&cchMaxClass,            // longest class string 
		&cValues,                // number of values for this key 
		&cchMaxValue,            // longest value name 
		&cbMaxValueData,         // longest value data 
		&cbSecurityDescriptor,   // security descriptor 
		&ftLastWriteTime);       // last write time 

	if ( ERROR_SUCCESS != retCode )
	{
		return FALSE;
	}

	// Enumerate the subkeys, until RegEnumKeyEx fails.
	if (cSubKeys > 0)
	{
		for (i=0; i<cSubKeys; i++) 
		{ 
			cbName = MAX_KEY_LENGTH;
			retCode = RegEnumKeyEx(m_hKey, i,
				achKey, 
				&cbName, 
				NULL, 
				NULL, 
				NULL, 
				&ftLastWriteTime); 
			if (retCode == ERROR_SUCCESS) 
			{
				scArray.Add( CString( achKey ) );
			}
		}
	} 

	return TRUE;
}