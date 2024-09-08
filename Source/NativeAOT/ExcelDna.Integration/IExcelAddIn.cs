//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

namespace ExcelDna.Integration
{
	public interface IExcelAddIn
	{
		void AutoOpen();
		void AutoClose();
	}
}
