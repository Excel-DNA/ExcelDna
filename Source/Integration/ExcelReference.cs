/*
  Copyright (C) 2005, 2006, 2007 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDna.Integration
{
	public class ExcelReference
	{
		List<XlOper.XlRectangle> rectangles = new List<XlOper.XlRectangle>();
		int sheetId;

		public ExcelReference(int row, int column)
			: this(row, row, column, column)
		{
		}

		// DOCUMENT: If no SheetId is given, assume the Active Sheet
		public ExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast) :
			this(rowFirst, rowLast, columnFirst, columnLast, 0)
		{
			try
			{
				ExcelReference r = (ExcelReference)XlCall.Excel(XlCall.xlSheetId);
				sheetId = r.sheetId;
			}
			catch
			{
				// CONSIDER: throw or 'default' behaviour?
			}
		}

		public ExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast, int sheetId)
		{
			this.sheetId = sheetId;
			XlOper.XlRectangle rect = new XlOper.XlRectangle(rowFirst, rowLast, columnFirst, columnLast);
			rectangles.Add(rect);
		}

		// THROWS: OverFlowException if the arguments exceed the allowed size
		// or if the number of Inner References exceeds 65000
		public void AddReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
		{
			if (rectangles.Count < ushort.MaxValue)
				rectangles.Add(new XlOper.XlRectangle(rowFirst, rowLast, columnFirst, columnLast));
			else 
				throw new OverflowException("Maximum number of references exceeded");
		}

		public int RowFirst
		{
			get { return rectangles[0].RowFirst; }
		}

		public int RowLast
		{
			get { return rectangles[0].RowLast; }
		}

		public int ColumnFirst
		{
			get { return rectangles[0].ColumnFirst; }
		}

		public int ColumnLast
		{
			get { return rectangles[0].ColumnLast; }
		}

		public int SheetId
		{
			get { return sheetId; }
		}

		public List<ExcelReference> InnerReferences
		{
			get 
			{
				List<ExcelReference> inner = new List<ExcelReference>();
				foreach (XlOper.XlRectangle rect in rectangles)
				{
					inner.Add(new ExcelReference(rect.RowFirst, rect.RowLast, 
						rect.ColumnFirst, rect.ColumnLast, sheetId));
				}
				return inner;
			}
		}

		public object GetValue()
		{
			return XlCall.Excel(XlCall.xlCoerce, this);
		}

		public bool SetValue(object value)
		{
			return (bool)XlCall.Excel(XlCall.xlSet, this, value);
		}
	}
}
