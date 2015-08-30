//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

#pragma once
#include <string>

void LoaderInitialize(HMODULE hModule);
void LoaderUnload();

bool XlLibraryInitialize(XlAddInExportInfo* pExportInfo);
void XlLibraryUnload();

std::wstring GetAddInFullPath();
HRESULT GetAddInName(std::wstring& addInName);

