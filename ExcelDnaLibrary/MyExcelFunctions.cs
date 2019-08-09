using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;

namespace ExcelDnaLibrary
{
    public class MyExcelFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string HelloDna(string name)
        {
            return "Hello " + name;
        }


        [ExcelFunction(Description = "GetLastRow")]
        public static int GetLastRow()
        {
            //returns active cell
            var excelref = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            return excelref.RowLast;

        }

        public static void WriteDataFromArrayToRange()
        {
            var xlApp = (Application)ExcelDnaUtil.Application;

            var wb = xlApp.ActiveWorkbook;
            if (wb == null)
                return;

            Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            ws.Range["A1"].Value = "Date";
            ws.Range["B1"].Value = "Value";

            var headerRow = ws.Range["A1", "B1"];
            headerRow.Font.Size = 12;
            headerRow.Font.Bold = true;

            // Generally it's faster to write an array to a range
            var values = new object[100, 2];
            var startDate = new DateTime(2007, 1, 1);
            var rand = new Random();
            for (var i = 0; i < 100; i++)
            {
                values[i, 0] = startDate.AddDays(i);
                values[i, 1] = rand.NextDouble();
            }

            ws.Range["A2"].Resize[100, 2].Value = values;
            ws.Columns["A:A"].EntireColumn.AutoFit();

            // Add a chart
            var dataRange = ws.Range["A1:B101"];
            dataRange.Select();
            ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();

        }

        public static void GetTopTen()
        {

            var xlApp = (Application)ExcelDnaUtil.Application;

            var wb = xlApp.ActiveWorkbook;
            if (wb == null)
                return;

            Worksheet ws = wb.ActiveSheet;

            var dataRange = ws.Range["A1:B101"];
            dataRange.Select();
            ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();

            // add autofilter
            //dataRange.AutoFilter(Field: 1, Criteria1: "t", VisibleDropDown: true);

            dataRange.AutoFilter(Field: 1, Criteria1: 10, Operator: XlAutoFilterOperator.xlTop10Items);
        }



    }
}