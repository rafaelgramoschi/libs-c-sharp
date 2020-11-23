/*
Tools > Get Tools and Features...
Check Office/sharepoint dev & modify installation

manually add Microsoft.Office.interop.excel
add references.. > browse > C:\Windows\assembly\(GAC\ OR \GAC_MSIL)\microsoft.office.interop.excel\n.n.n.n\Microsoft.Office.Interop.Excel.dll

in solution explorer > assemblies > microsoft.office.interop.excel > properties > set: Embed Interop Types to YES/TRUE
*/

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace SheetsQuickstart
{
    class ExcelReader
    {
        const string FILE = @"C:\excelfile.xlsx";
        Excel.Application xlApp;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range range;

        public void TestExcelReader()
        {
            this.PrintData();
            this.ReleaseMem();
        }

        public ExcelReader()
        {
            // instanciate excel Application object
            xlApp = new Excel.Application();

            // select "book" (file)
            // https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
            
            wb = xlApp.Workbooks.Open(FILE, 0, true, 5, "", "", true,
                                      Excel.XlPlatform.xlWindows, "\t",
                                      false, false, 0, true, 1, 0);
            // select which sheet
            ws = (Excel.Worksheet)wb.Worksheets.get_Item(1); //wb.Sheets[1];

            /*
                For reading entire content of an Excel file in C#,
                we have to know how many cells used in the Excel file.
            
                ("UsedRange" property of xlWorkSheet)

                It includes any cell that has ever been used.
                It will return the last cell of used area. 
            */
            range = ws.UsedRange; //full file

            /* MULTIPLE RANGES
            range = (Excel.Range)ws.Cells[ 
                                            ws.Cells[1,1],
                                            ws.Cells[3,3]
                                         ];
            */
        }

        public ExcelReader(string filepath)
        {
            xlApp = new Excel.Application();
            wb = xlApp.Workbooks.Open(filepath, 0, true, 5, "", "", true,
                                      Excel.XlPlatform.xlWindows, "\t",
                                      false, false, 0, true, 1, 0);
            ws = (Excel.Worksheet)wb.Worksheets.get_Item(1); //wb.Sheets[1];
            range = ws.UsedRange; //full file
        }

        public void PrintData()
        {
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            // _xVal contains the value for each single cell
            object _xVal; // used to not use .. notation on COM objects

            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= cols; col++)
                {
                    _xVal = ((Excel.Range)range.Cells[row, col]).Value2;
                    
                    //each new line
                    if (col == 1)
                    {
                        Console.Write("\r\n");
                    }

                    // first row (columns' names)
                    if (row == 1)
                    {
                        Console.Write($"{_xVal.ToString().ToUpper()}\t");
                    }

                    if (row > 1 && range.Cells[row, col] != null && _xVal != null)
                    {
                        Console.Write( _xVal.ToString() + "\t");
                    }
                    //add useful things here!   
                }
            }
        }

        public void ReleaseMem()
        {
            // garbage collector
            GC.Collect();
            // wait for pending stuff going on to finish
            GC.WaitForPendingFinalizers();

            // Don't use ReleaseComObject([smthing].[smthing].[smthing])
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(ws);

            wb.Close();
            Marshal.ReleaseComObject(wb);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
