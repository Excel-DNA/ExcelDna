/*
	CDetectDotNet Version 2.2
	Last updated : Jul 21 2005

	Copyright(C) Nishant Sivakumar 2005. All Rights Reserved.	
	Contact Email : nish@voidnish.com
	URLs : www.voidnish.com, blog.voidnish.com	
*/


#include "StdAfx.h"
#include "detectdotnet.h"

const int g_cPathSize = 512;

CDetectDotNet::CDetectDotNet() : m_szInstallRootPath(NULL)
{
	m_bDotNetPresent = IsDotNetPresentInternal();
	if(m_bDotNetPresent)
	{
		TCHAR szRootPath[g_cPathSize];
		if(GetInstallRootPathInternal(szRootPath, g_cPathSize))
		{
			size_t siz = _tcslen(szRootPath);
			m_szInstallRootPath = new TCHAR[siz + 1];
			_tcsncpy(m_szInstallRootPath, szRootPath, siz);
			m_szInstallRootPath[siz] = NULL;
		}
	}
}

CDetectDotNet::~CDetectDotNet()
{
	delete[] m_szInstallRootPath;
}

//Public functions
bool CDetectDotNet::IsDotNetPresent()
{
	return m_bDotNetPresent;
}

size_t CDetectDotNet::EnumerateCLRVersions(vector<string>& CLRVersions)
{
	CLRVersions.clear();
	USES_CONVERSION;
	vector<string> PossibleCLRVersions;
	EnumeratePossibleCLRVersionsInternal(PossibleCLRVersions);
	for(vector<string>::iterator it = PossibleCLRVersions.begin(); 
		it < PossibleCLRVersions.end(); it++)
	{				
		if(CheckForSpecificCLRVersionInternal(A2CW((*it).c_str())))
		{			
			CLRVersions.push_back(GetVersionFromFolderName(*it));
		}		
	}
	return CLRVersions.size();
}

bool CDetectDotNet::GetInstallRootPath(TCHAR* szRootPath, DWORD dwBufferSize)
{
	bool bRet = false;
	if(m_szInstallRootPath)
	{
		size_t siz = _tcslen(m_szInstallRootPath);
		if(dwBufferSize > siz)
		{
			_tcsncpy(szRootPath, m_szInstallRootPath, siz);
			szRootPath[siz] = NULL;
			bRet = true;
		}
	}
	return bRet;
}

//Protected functions

bool CDetectDotNet::CheckForSpecificCLRVersionInternal(LPCWSTR pszVersion)
{
	bool bRet = false;
	if( m_bDotNetPresent )
	{
		UINT prevErrMode = SetErrorMode(SEM_FAILCRITICALERRORS);
		HMODULE hModule = LoadLibrary(_T("mscoree"));
		if(hModule)
		{	
			FPGetRequestedRuntimeInfo pGetRequestedRuntimeInfo =
				reinterpret_cast<FPGetRequestedRuntimeInfo>(
				GetProcAddress(hModule, "GetRequestedRuntimeInfo"));
			if(pGetRequestedRuntimeInfo)
			{
				LPWSTR dirBuff = NULL;
				DWORD dwDir = 0;
				LPWSTR verBuff = NULL;
				DWORD dwVer = 0;

				pGetRequestedRuntimeInfo(NULL, pszVersion,
					NULL,0,0,
					dirBuff, dwDir, &dwDir,
					verBuff, dwVer, &dwVer);

				dirBuff = new WCHAR[dwDir + 1];
				verBuff = new WCHAR[dwVer + 1];

				HRESULT hr = pGetRequestedRuntimeInfo(NULL, pszVersion,
					NULL,0,0,dirBuff, dwDir, &dwDir,verBuff, dwVer, &dwVer);

				bRet = (hr == S_OK);

				delete[] verBuff;
				delete[] dirBuff;
			}
			FreeLibrary(hModule);			
		}
		SetErrorMode(prevErrMode);
	}
	return bRet;
}

size_t CDetectDotNet::EnumeratePossibleCLRVersionsInternal(
	vector<string>& PossibleCLRVersions)
{
	PossibleCLRVersions.clear();
	if(m_bDotNetPresent)
	{
		TCHAR szRootBuff[g_cPathSize];
		if(GetInstallRootPath(szRootBuff, g_cPathSize))
		{
			WIN32_FIND_DATA finddata = {0};
			_tcsncat(szRootBuff, _T("*"), 1);
			HANDLE hFind = FindFirstFile(szRootBuff, &finddata);
			if(hFind != INVALID_HANDLE_VALUE)
			{
				do
				{
					if( finddata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY )
					{
						PossibleCLRVersions.push_back(finddata.cFileName);
					}
				}while(FindNextFile(hFind, &finddata));
				FindClose(hFind);
			}
		}
	}
	return PossibleCLRVersions.size();
}

string CDetectDotNet::GetDotNetVersion(LPCTSTR szFolder)
{
	string strRet = _T("");
	LPVOID m_lpData = NULL;
	TCHAR buff[MAX_PATH + 1] = {0};
	_tcsncpy(buff, szFolder, MAX_PATH);
	DWORD dwHandle = 0;
	DWORD dwVerInfoSize = GetFileVersionInfoSize(buff, &dwHandle);	

	if(dwVerInfoSize != 0) //Success
	{
		m_lpData = malloc(dwVerInfoSize);
		if(GetFileVersionInfo(buff, dwHandle, dwVerInfoSize, m_lpData) == FALSE)
		{
			free(m_lpData);
			m_lpData = NULL;
		}
		else
		{
			UINT cbTranslate = 0;

			struct LANGANDCODEPAGE 
			{
				WORD wLanguage;
				WORD wCodePage;
			} *lpTranslate;	

			if(VerQueryValue(m_lpData,_T("\\VarFileInfo\\Translation"),
				(LPVOID*)&lpTranslate,&cbTranslate))
			{
				int count = (int)(cbTranslate/sizeof(struct LANGANDCODEPAGE));

				for(int i=0; i < count; i++ )
				{
					TCHAR SubBlock[128];
					HRESULT hr = StringCchPrintf(SubBlock, 127,_T("\\StringFileInfo\\%04x%04x\\%s"),
						lpTranslate[i].wLanguage,lpTranslate[i].wCodePage,_T("FileVersion"));	

					if(SUCCEEDED(hr))
					{
						UINT dwBytes = 0;
						TCHAR* lpBuffer;

						if(VerQueryValue(m_lpData, SubBlock, (LPVOID*)&lpBuffer, &dwBytes))
						{	
							USES_CONVERSION;
							strRet = T2A(lpBuffer);
							for(unsigned int x = 0, j = 0; j < strRet.size(); j++)
							{
								if(strRet[j] == '.')
								{
									if(++x == 3)
									{
										strRet.erase(j,strRet.size() - j);				
										break;
									}
								}
							}
							break;
						}
					}		
				}
			}			
		}
	}
	return strRet;
}

string CDetectDotNet::GetVersionFromFolderName(string szFolderName)
{
	string strRet = "<Version could not be extracted from mscorlib>";
	TCHAR szRootPath[g_cPathSize];
	if(GetInstallRootPath(szRootPath, g_cPathSize))
	{
		USES_CONVERSION;
		string szFilepath = T2A(szRootPath);
		szFilepath += (szFolderName + "\\mscorlib.dll");
		string s = GetDotNetVersion(A2CT(szFilepath.c_str()));
		if(s.size() > 0)
			strRet = s;
	}
	return strRet;
}

bool CDetectDotNet::IsDotNetPresentInternal()
{
	bool bRet = false;
	//Attempt to LoadLibrary "mscoree.dll" (the CLR EE shim)
	HMODULE hModule = LoadLibrary(_T("mscoree"));
	if(hModule)
	{	
		//Okay - that worked, but just to ensure that this is
		//not a placeholder DLL shipped with some earlier OS versions,
		//we attempt to GetProcAddress "GetCORVersion".
		bRet = (GetProcAddress(hModule, "GetCORVersion") != NULL);
		FreeLibrary(hModule);
	}
	return bRet;
}

bool CDetectDotNet::GetInstallRootPathInternal(TCHAR* szRootPath, DWORD dwBufferSize)
{
	bool bRet = false;
	TCHAR szRegPath[] = _T("SOFTWARE\\Microsoft\\.NETFramework");
	HKEY hKey = NULL;
	if(RegOpenKeyEx(HKEY_LOCAL_MACHINE, szRegPath, 0, 
		KEY_READ, &hKey) == ERROR_SUCCESS)
	{	
		DWORD dwSize = dwBufferSize;
		if(RegQueryValueEx(hKey, _T("InstallRoot"), NULL, NULL,
			reinterpret_cast<LPBYTE>(szRootPath), 
			&dwSize) == ERROR_SUCCESS)
		{
			bRet = (dwSize <= dwBufferSize);
		}
		RegCloseKey(hKey);
	}
	return bRet;
}