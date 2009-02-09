/*
  Copyright (C) 2005-2008 Govert van Drimmelen

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
    // CAUTION: The ExcelReference class is also called via reflection by the ExcelDna.Loader marshaler.
	public class ExcelReference
	{
        internal class ExcelRectangle
        {
            // TODO: EXCEL2007 - Sort out sizes and compatibility
            public ushort RowFirst;
            public ushort RowLast;
            public byte ColumnFirst;
            public byte ColumnLast;

            internal ExcelRectangle(int rowFirst, int rowLast, int columnFirst, int columnLast)
            {
                RowFirst = (ushort)Math.Max(0, Math.Min(rowFirst, ushort.MaxValue));
                RowLast = (ushort)Math.Max(0, Math.Min(rowLast, ushort.MaxValue));
                ColumnFirst = (byte)Math.Max(0, Math.Min(columnFirst, byte.MaxValue)); ;
                ColumnLast = (byte)Math.Max(0, Math.Min(columnLast, byte.MaxValue)); ; ;
            }
        }
		List<ExcelRectangle> rectangles = new List<ExcelRectangle>();
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
			ExcelRectangle rect = new ExcelRectangle(rowFirst, rowLast, columnFirst, columnLast);
			rectangles.Add(rect);
		}

		// THROWS: OverFlowException if the arguments exceed the allowed size
		// or if the number of Inner References exceeds 65000
		public void AddReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
		{
			if (rectangles.Count < ushort.MaxValue)
				rectangles.Add(new ExcelRectangle(rowFirst, rowLast, columnFirst, columnLast));
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
				foreach (ExcelRectangle rect in rectangles)
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

        // CAUTION: These 'private' functions are called via reflection by the ExcelDna.Loader marshaler
        // Returns arrays containing all the inner rectangles (including the one we pretend is outside).
        private int[][] GetRectangles()
        {
            int[][] intRects = new int[rectangles.Count][];
            for (int i = 0; i < rectangles.Count; i++)
            {
                ExcelRectangle rect = rectangles[i];
                intRects[i] = new int[] {rect.RowFirst, rect.RowLast,
                    rect.ColumnFirst, rect.ColumnLast};
            }
            return intRects;
        }

        private int GetRectangleCount()
        {
            return rectangles.Count;
        }
	}
}
