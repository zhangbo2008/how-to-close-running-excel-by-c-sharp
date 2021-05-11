using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


using Excel = Microsoft.Office.Interop.Excel;
namespace Microsoft.Office.Interop

{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp;
     
            Excel.Workbook xlWorkBook;
          //  xlApp = new Excel.Application();
          xlApp= (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            xlApp.Visible = true;
            xlApp.ActiveWorkbook.Close();
           
     /*       xlWorkBook = xlApp.Workbooks.Open("c:/1.xlsx");
            xlWorkBook.Close();*/

            /*    Excel._Application objExcel;

                objExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                objExcel.ActiveWorkbook.Save();
                int a = 3333;
                //  objExcel.Workbooks.Open("c://1.xlsx");
                MyExcel.Application EXC1 = new Excel.Application();*/


            // System.Console.Write("dsafdsfsad");
        }
















    }
}
