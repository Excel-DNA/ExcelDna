// Copied from somewhere - GvD

#include "stdafx.h"
#include "DetectFx.h"

#define CountOf(x) sizeof(x)/sizeof(*x)

const TCHAR *g_szNetfx20RegKeyName = _T("Software\\Microsoft\\NET Framework Setup\\NDP\\v2.0.50727");
const TCHAR *g_szNetfxStandardRegValueName = _T("Install");

bool IsNetfx20Installed();
bool RegistryGetValue(HKEY, const TCHAR*, const TCHAR*, DWORD, LPBYTE, DWORD);


/******************************************************************
Function Name:	IsNetfx20Installed
Description:	Uses the detection method recommended at
                http://msdn2.microsoft.com/en-us/library/aa480243.aspx
                to determine whether the .NET Framework 2.0 is
                installed on the machine
Inputs:         NONE
Results:        true if the .NET Framework 2.0 is installed
                false otherwise
******************************************************************/
bool IsNetfx20Installed()
{
	bool bRetValue = false;
    DWORD dwRegValue=0;

	if (RegistryGetValue(HKEY_LOCAL_MACHINE, g_szNetfx20RegKeyName, g_szNetfxStandardRegValueName, NULL, (LPBYTE)&dwRegValue, sizeof(DWORD)))
	{
		if (1 == dwRegValue)
			bRetValue = true;
	}

	return bRetValue;
}


/******************************************************************
Function Name:  RegistryGetValue
Description:    Get the value of a reg key
Inputs:         HKEY hk - The hk of the key to retrieve
                TCHAR *pszKey - Name of the key to retrieve
                TCHAR *pszValue - The value that will be retrieved
                DWORD dwType - The type of the value that will be retrieved
                LPBYTE data - A buffer to save the retrieved data
                DWORD dwSize - The size of the data retrieved
Results:        true if successful, false otherwise
******************************************************************/
bool RegistryGetValue(HKEY hk, const TCHAR * pszKey, const TCHAR * pszValue, DWORD dwType, LPBYTE data, DWORD dwSize)
{
	HKEY hkOpened;

	// Try to open the key
	if (RegOpenKeyEx(hk, pszKey, 0, KEY_READ, &hkOpened) != ERROR_SUCCESS)
	{
		return false;
	}

	// If the key was opened, try to retrieve the value
	if (RegQueryValueEx(hkOpened, pszValue, 0, &dwType, (LPBYTE)data, &dwSize) != ERROR_SUCCESS)
	{
		RegCloseKey(hkOpened);
		return false;
	}
	
	// Clean up
	RegCloseKey(hkOpened);

	return true;
}

// ExcelDna - Added exports
bool DetectFxIsNet20Installed()
{
	return IsNetfx20Installed();// && CheckNetfxVersionUsingMscoree(g_szNetfx20VersionString);
}

int  DetectFxReadMajorVersion(TCHAR* pszVersion)
{
	TCHAR *pszToken = NULL;
	int iVersionPartCounter = 0;
	int iVersionMajor = 0;
	int iVersionMinor = 0;
	int iVersionBuild = 0;
	int iVersionRevision = 0;

	// This registry value should be of the format
	// v#.#.#####.##.  Try to parse the 4 parts of
	// the version here
	TCHAR* tok_context;
	pszToken = _tcstok_s(pszVersion+1, _T("."), &tok_context);
	while (NULL != pszToken)
	{
		iVersionPartCounter++;

		switch (iVersionPartCounter)
		{
		case 1:
			// Convert the major version value to an integer
			iVersionMajor = _tstoi(pszToken);
			break;
		case 2:
			// Convert the minor version value to an integer
			iVersionMinor = _tstoi(pszToken);
			break;
		case 3:
			// Convert the build number value to an integer
			iVersionBuild = _tstoi(pszToken);
			break;
		case 4:
			// Convert the revision number value to an integer
			iVersionRevision = _tstoi(pszToken);
			break;
		default:
			break;

		}

		// Get the next part of the version number
		pszToken = _tcstok_s(NULL, _T("."), &tok_context);
	}
	return iVersionMajor;
}