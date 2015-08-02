//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

#pragma once
#include <string>

void LoaderInitialize(HMODULE hModule);
void LoaderUnload();

bool XlLibraryInitialize(XlAddInExportInfo* pExportInfo);
void XlLibraryUnload();

std::wstring GetAddInFullPath();
HRESULT GetAddInName(std::wstring& addInName);

