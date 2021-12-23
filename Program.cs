using Microsoft.Office.Interop.Excel;
using System;
namespace ConsoleApp1
{
    class program
    {
        static void Main(string[] args)

        {

          Application excelApp = new Application(); 
            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Bismillah\FinacialSample1.xlsx");
            Worksheet  excelsheet = ( Worksheet)excelBook.Sheets[1];
            Range excelRange = excelsheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;
            for (int i = 0; i < rows; i++)
            {
                Console.Write("");
                for (int j = 0; j < cols; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        Console.Write(excelRange.Cells[i, j].Value2 + "\t");
                }
            }
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.ReadLine();
               
        }
    }
}


