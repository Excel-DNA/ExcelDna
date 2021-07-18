//  Copyright (c) Govert van Drimmelen. All rights reserved.

#pragma once

#include "exports.h"

int load_runtime_and_run(LPCWSTR basePath, XlAddInExportInfo* pExportInfo, HMODULE hModuleXll, LPCWSTR pathXll);
