#pragma once
#include <winreg.h>

#define REG_RECT 0x0001 
#define REG_POINT 0x0002 

extern BOOL g_bOpenRegistryX64;

// 注册表读取类
class IcRegistry
{ 
	// Construction 
public: 
	IcRegistry(HKEY hKeyRoot = HKEY_LOCAL_MACHINE); 
	virtual ~IcRegistry(); 

	struct REGINFO 
	{ 
		LONG lMessage; 
		DWORD dwType; 
		DWORD dwSize; 
	} m_Info; 
	// Operations 
public: 
	//BOOL VerifyKey (HKEY hKeyRoot, LPCTSTR pszPath); 
	BOOL VerifyKey (LPCTSTR pszPath); 
	BOOL VerifyValue (LPCTSTR pszValue); 
	BOOL VerifyValue (LPCTSTR pszValue, DWORD dwType ); 
	//BOOL CreateKey (HKEY hKeyRoot, LPCTSTR pszPath); 
	HKEY CreateKey (LPCTSTR pszPath);
	//BOOL Open (HKEY hKeyRoot, LPCTSTR pszPath);

	HKEY OpenSubKey( LPCTSTR pszPath ,bool bReadOnly=true); // NULL为打开失败
	void Close(); 

	BOOL QuerySubKeyNameList( CStringArray& scArray ); //获取子项名称列表

	BOOL DeleteValue (LPCTSTR pszValue); 
	BOOL DeleteValueKey (HKEY hKeyRoot, LPCTSTR pszPath); 

	BOOL Write (LPCTSTR pszKey, int iVal); 
	BOOL Write (LPCTSTR pszKey, DWORD dwVal); 
	BOOL Write (LPCTSTR pszKey, LPCTSTR pszVal); 
	BOOL Write (LPCTSTR pszKey, CStringList& scStringList); 
	BOOL Write (LPCTSTR pszKey, CByteArray& bcArray); 
	BOOL Write (LPCTSTR pszKey, CStringArray& scArray); 
	BOOL Write (LPCTSTR pszKey, CDWordArray& dwcArray); 
	BOOL Write (LPCTSTR pszKey, CWordArray& wcArray); 
	BOOL Write (LPCTSTR pszKey, LPCRECT rcRect); 
	BOOL Write (LPCTSTR pszKey, LPPOINT& lpPoint); 

	BOOL Read (LPCTSTR pszKey, int& iVal); 
	BOOL Read (LPCTSTR pszKey, DWORD& dwVal); 
	BOOL Read (LPCTSTR pszKey, CString& sVal); 
	BOOL Read (LPCTSTR pszKey, CStringList& scStringList); 
	BOOL Read (LPCTSTR pszKey, CStringArray& scArray); 
	BOOL Read (LPCTSTR pszKey, CDWordArray& dwcArray); 
	BOOL Read (LPCTSTR pszKey, CWordArray& wcArray); 
	BOOL Read (LPCTSTR pszKey, CByteArray& bcArray); 
	BOOL Read (LPCTSTR pszKey, LPPOINT& lpPoint); 
	BOOL Read (LPCTSTR pszKey, LPRECT& rcRect); 
	//add for wbwa 
	BOOL Read( LPCTSTR lpszValueName, LPVOID lpReturnBuffer, int nSize ); 
	BOOL ReadDWORD( LPCTSTR, DWORD *, DWORD *pdwLastError = NULL ); 
	BOOL ReadString( const TCHAR *, LPVOID, int, DWORD *pdwLastError = NULL ); 

private: 
	HKEY m_hKey; 
	CString m_sPath; 
	DWORD m_dwLastError; 
	
	enum
	{ 
		MAX_KEY_LENGTH = 255,
		MAX_VALUE_NAME = 16383
	};
};