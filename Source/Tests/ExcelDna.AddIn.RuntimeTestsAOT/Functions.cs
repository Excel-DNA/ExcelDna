﻿using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Functions
    {
        [ExcelFunction]
        public static string NativeHello(string name)
        {
            return $"Hello {name}!";
        }

        [ExcelFunction]
        public static int NativeSum(int i1, int i2)
        {
            return i1 + i2;
        }

        [ExcelAsyncFunction]
        public static async Task<string> NativeAsyncTaskHello(string name, int msDelay)
        {
            await Task.Delay(msDelay);
            return $"Hello native async task {name}";
        }

        [ExcelFunction]
        public static string NativeApplicationName()
        {
            return (string)ExcelDnaUtil.DynamicApplication.Get("Name")!;
        }

        [ExcelFunction]
        public static double NativeApplicationGetCellValue(string cell)
        {
            var workbook = (IDynamic)ExcelDnaUtil.DynamicApplication.Get("ActiveWorkbook")!;
            var sheets = (IDynamic)workbook.Get("Sheets")!;
            var sheet = (IDynamic)sheets[1]!;
            var range = (IDynamic)sheet.Get("Range", [cell])!;
            return (double)range.Get("Value")!;
        }

        [ExcelFunction]
        public static double NativeApplicationGetCellValueT(string cell)
        {
            var workbook = ExcelDnaUtil.DynamicApplication.Get<IDynamic>("ActiveWorkbook");
            var sheets = workbook.Get<IDynamic>("Sheets");
            var sheet = (IDynamic)sheets[1]!;
            var range = sheet.Get<IDynamic>("Range", [cell]);
            return range.Get<double>("Value");
        }

        [ExcelFunction]
        public static int NativeApplicationAlignCellRight(string cell)
        {
            var workbook = ExcelDnaUtil.DynamicApplication.Get<IDynamic>("ActiveWorkbook");
            var sheets = workbook.Get<IDynamic>("Sheets");
            var sheet = (IDynamic)sheets[1]!;
            var range = sheet.Get<IDynamic>("Range", [cell]);
            range.Set("HorizontalAlignment", -4152);
            return range.Get<int>("HorizontalAlignment");
        }

        [ExcelFunction]
        public static string NativeApplicationAddCellComment(string cell, string comment)
        {
            var workbook = ExcelDnaUtil.DynamicApplication.Get<IDynamic>("ActiveWorkbook");
            var sheets = workbook.Get<IDynamic>("Sheets");
            var sheet = (IDynamic)sheets[1]!;
            var range = sheet.Get<IDynamic>("Range", [cell]);
            var newComment = (IDynamic)range.Invoke("AddComment", [comment])!;
            return newComment.Invoke<string>("Text", []);
        }

        [ExcelFunction]
        public static object NativeRangeConcat2(object[,] values)
        {
            string result = "";
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    object value = values[i, j];
                    result += value.ToString();
                }
            }
            return result;
        }

        [ExcelFunction]
        public static string NativeNullableDouble(double? d)
        {
            return "Native Nullable VAL: " + (d.HasValue ? d : "NULL");
        }

        [ExcelFunction]
        public static string NativeOptionalDouble(double d = 1.23)
        {
            return "Native Optional VAL: " + d.ToString();
        }

        [ExcelFunction]
        public static string NativeRangeAddress(IRange r)
        {
            return "Native Address: " + r.Get<string>("Address");
        }
    }
}
