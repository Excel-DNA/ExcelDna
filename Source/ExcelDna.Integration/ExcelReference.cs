//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelDna.Integration
{
    // CAUTION: The ExcelReference class is also called via reflection by the ExcelDna.Loader marshaler.
	public class ExcelReference : IEquatable<ExcelReference>
	{
	    struct ExcelRectangle : IEquatable<ExcelRectangle>
        {
            public readonly int RowFirst;
            public readonly int RowLast;
            public readonly int ColumnFirst;
            public readonly int ColumnLast;

            internal ExcelRectangle(int rowFirst, int rowLast, int columnFirst, int columnLast)
            {
                // CONSIDER: Throw or truncate for errors
                RowFirst    = Clamp(rowFirst, 0, ExcelDnaUtil.ExcelLimits.MaxRows - 1);
                RowLast     = Clamp(rowLast, 0, ExcelDnaUtil.ExcelLimits.MaxRows - 1);
                ColumnFirst = Clamp(columnFirst, 0, ExcelDnaUtil.ExcelLimits.MaxColumns - 1);
                ColumnLast  = Clamp(columnLast, 0, ExcelDnaUtil.ExcelLimits.MaxColumns - 1);
                
                // CONSIDER: Swap or truncate rect ??
                //if (RowLast < RowFirst) RowLast = RowFirst;
                //if (ColumnLast < ColumnFirst) ColumnLast = RowFirst;
            }

            static int Clamp(int value, int min, int max)
            {
                Debug.Assert(min <= max);
                if (value < min) return min;
                if (value > max) return max;
                return value;
            }

            public override bool Equals(object obj)
            {
                if (obj.GetType() != typeof (ExcelRectangle)) return false;
                return Equals((ExcelRectangle) obj);
            }

            public bool Equals(ExcelRectangle other)
            {
                return other.RowFirst == RowFirst && other.RowLast == RowLast && other.ColumnFirst == ColumnFirst && other.ColumnLast == ColumnLast;
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int result = RowFirst;
                    result = (result*397) ^ RowLast;
                    result = (result*397) ^ ColumnFirst;
                    result = (result*397) ^ ColumnLast;
                    return result;
                }
            }
        }

		readonly IntPtr sheetId;
        // Save first rectangle explicitly, to avoid array if we don't need one.
        readonly ExcelRectangle rectangle;
	    ExcelRectangle[] furtherRectangles; // Can't make readonly (yet) since we still have the obsolete 'AddReference'

		public ExcelReference(int row, int column)
			: this(row, row, column, column)
		{
		}

		// DOCUMENT: If no SheetId is given, assume the Active (Front) Sheet
		public ExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast) :
			this(rowFirst, rowLast, columnFirst, columnLast, IntPtr.Zero)
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

		public ExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast, IntPtr sheetId)
		{
			this.sheetId = sheetId;
			rectangle = new ExcelRectangle(rowFirst, rowLast, columnFirst, columnLast);
		}

        // TODO: Consider how to deal with invalid sheetName. I presume xlSheetId will fail.
        // Perhaps throw a custom exception...?
        public ExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast, string sheetName)
        {
            ExcelReference sheetRef = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);
            this.sheetId = sheetRef.SheetId;
            rectangle = new ExcelRectangle(rowFirst, rowLast, columnFirst, columnLast);
        }

        public ExcelReference(int[][] rectangles, IntPtr sheetId)
        {
            this.sheetId = sheetId;
            int rectCount = rectangles.Length;
            int[] rect = rectangles[0];
            rectangle = new ExcelRectangle(rect[0], rect[1], rect[2], rect[3]);

            int furtherRectangleCount = rectCount - 1;
            if (furtherRectangleCount > 0)
            {
                furtherRectangles = new ExcelRectangle[furtherRectangleCount];
                for (int i = 0; i < furtherRectangleCount; i++)
                {
                    rect = rectangles[i + 1];
                    Debug.Assert(rect.Length == 4);
                    furtherRectangles[i] = new ExcelRectangle(rect[0], rect[1], rect[2], rect[3]);
                }
            }
        }

		// THROWS: OverFlowException if the arguments exceed the allowed size
		// or if the number of Inner References exceeds 65000
        [Obsolete("An ExcelReference should never be modified.")]
		public void AddReference(int rowFirst, int rowLast, int columnFirst, int columnLast)
		{
            ExcelRectangle rect = new ExcelRectangle(rowFirst, rowLast, columnFirst, columnLast);
            if (furtherRectangles == null)
            {
                furtherRectangles = new ExcelRectangle[1];
                furtherRectangles[0] = rect;
                return;
            }
			if (furtherRectangles.Length >= ushort.MaxValue - 1)
				throw new OverflowException("Maximum number of references exceeded");

            ExcelRectangle[] newRectangles = new ExcelRectangle[furtherRectangles.Length + 1];
            Array.Copy(furtherRectangles, newRectangles, furtherRectangles.Length);
            newRectangles[newRectangles.Length - 1] = rect;
            furtherRectangles = newRectangles;
		}

		public int RowFirst
		{
			get { return rectangle.RowFirst; }
		}

		public int RowLast
		{
			get { return rectangle.RowLast; }
		}

		public int ColumnFirst
		{
			get { return rectangle.ColumnFirst; }
		}

		public int ColumnLast
		{
			get { return rectangle.ColumnLast; }
		}

		public IntPtr SheetId
		{
			get { return sheetId; }
		}

        // TODO: Document the fact that the returned list is a copy, and does not modify the list of Rectangles in the ExcelReference.
		public List<ExcelReference> InnerReferences
		{
			get 
			{
				List<ExcelReference> inner = new List<ExcelReference>();
                // Add ourselves - which is the first one
                inner.Add(new ExcelReference(RowFirst, RowLast, ColumnFirst, ColumnLast, SheetId));
                // And then all the others
                if (furtherRectangles != null)
                {
                    foreach (ExcelRectangle rect in furtherRectangles)
                    {
                        inner.Add(new ExcelReference(rect.RowFirst, rect.RowLast, rect.ColumnFirst, rect.ColumnLast, SheetId));
                    }
                }
				return inner;
			}
		}

		public object GetValue()
		{
			return XlCall.Excel(XlCall.xlCoerce, this);
		}

        // DOCUMENT: Strange behaviour with SetValue...???
		public bool SetValue(object value)
		{
			return (bool)XlCall.Excel(XlCall.xlSet, this, value);
		}

        // CAUTION: These 'private' functions are called via reflection by the ExcelDna.Loader marshaler
        // Returns arrays containing all the inner rectangles (including the one we pretend is outside).
        private int[][] GetRectangles()
        {
            int rectangleCount = GetRectangleCount();
            int furtherRectangleCount = rectangleCount - 1;
            int[][] intRects = new int[rectangleCount][];
            intRects[0] = new int[] { RowFirst, RowLast, ColumnFirst, ColumnLast };
            for (int i = 0; i < furtherRectangleCount; i++)
            {
                // We only get here if there are further rectangles
                Debug.Assert(furtherRectangles != null);
                ExcelRectangle rect = furtherRectangles[i];
                intRects[i+1] = new int[] {rect.RowFirst, rect.RowLast, rect.ColumnFirst, rect.ColumnLast};
            }
            return intRects;
        }

        private int GetRectangleCount()
        {
            if (furtherRectangles == null)
                return 1;
            return furtherRectangles.Length + 1;
        }

        // Structural equality implementation
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
                return false;
            if (ReferenceEquals(this, obj))
                return true;
            if (obj.GetType() != typeof (ExcelReference))
                return false;
            return Equals((ExcelReference) obj);
        }

	    public bool Equals(ExcelReference other)
	    {
	        if (ReferenceEquals(null, other))
                return false;
	        if (ReferenceEquals(this, other))
                return true;
            // Implement equality check based on contents.
            if (!rectangle.Equals(other.rectangle))
                return false;
            if (!sheetId.Equals(other.sheetId))
                return false;

            // Count will implicitly check null furtherRectangle arrays on either side too
            int rectangleCount = GetRectangleCount();
            if (rectangleCount != other.GetRectangleCount())
                return false;

            int furtherRectangleCount = rectangleCount - 1;
            for (int i = 0; i < furtherRectangleCount; i++)
            {
                if (!furtherRectangles[i].Equals(other.furtherRectangles[i]))
                    return false; 
            }
	        return true;
	    }

        // We need to take some care with the Hash Code here, since we use the ExcelReference with structural comparison
        // in some Dictionaries.
	    public override int GetHashCode()
	    {
            // One of the ideas from http://stackoverflow.com/questions/263400/what-is-the-best-algorithm-for-an-overridden-system-object-gethashcode
            const int b = 378551;
            int a = 63689;
            int hash = 0;
            
            unchecked
	        {
                hash = hash * a + rectangle.GetHashCode();
                a = a * b;
                if (furtherRectangles != null)
                {
                    for (int i = 0; i < furtherRectangles.Length; i++)
                    {
                        hash = hash * a + furtherRectangles[i].GetHashCode();
                        a = a * b;
                    }
                }
                hash *= 397;
            }
	        return hash ^ sheetId.GetHashCode();
	    }

        public override string ToString()
        {
            return string.Format("({0},{1} : {2},{3}) - {4}", RowFirst, ColumnFirst, RowLast, ColumnLast, SheetId);
        }

        public static bool operator ==(ExcelReference left, ExcelReference right) { return left.Equals(right); }
        public static bool operator !=(ExcelReference left, ExcelReference right) { return !left.Equals(right); }
	}
}
