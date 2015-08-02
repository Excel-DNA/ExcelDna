//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace ExcelDna.Integration
{
	public interface IExcelAddIn
	{
		void AutoOpen();
		void AutoClose();
	}
}
