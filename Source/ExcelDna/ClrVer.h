#pragma once

#include <Windows.h>

struct ClrVerInfo
{
    BOOL ClrInstalled;     // Is some version of the Clr installed
	BOOL v2PlusInstalled; // Is v2.0 or later installed
    WCHAR LatestRuntime[30]; // Latest runtime installed
    DWORD LatestRuntimeNumChars;
};

void InitializeClrVerInfo(ClrVerInfo& clrVerInfo);
void GetClrVerInfo(ClrVerInfo& clrVerInfo);