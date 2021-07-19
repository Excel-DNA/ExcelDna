//  Copyright (c) Govert van Drimmelen. All rights reserved.

#pragma once

#include "exports.h"

#include <string>

void LoaderInitialize(HMODULE hModule);
void LoaderUnload(bool processTerminating);

bool XlLibraryInitialize(XlAddInExportInfo* pExportInfo);
void XlLibraryUnload();

void LockModule();
void UnlockModule();

std::wstring GetAddInFullPath();
// HRESULT GetAddInName(std::wstring& addInName);
