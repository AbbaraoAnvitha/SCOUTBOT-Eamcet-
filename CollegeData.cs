using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Runtime.InteropServices;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace Scout
{


    public class CollegeData
    {

        
        

        public string findBranches(string college)
        {
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Workbook workBook;
            Microsoft.Office.Interop.Excel.Worksheet workSheet;
            DataTable dt;
            string expression;
            app = new Microsoft.Office.Interop.Excel.Application();
            workBook = app.Workbooks.Open(@"C:\\Users\\Sathya\\Pictures\\SCOUT-master\\Dialogs\\eamcetdataset.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;

            int index = 0;
            object rowIndex = 2;
            dt = new DataTable();
            dt.Columns.Add("College");
            dt.Columns.Add("Branch");
            dt.Columns.Add("ocb");
            dt.Columns.Add("ocg");
            /*dt.Columns.Add("scb");
            dt.Columns.Add("scg");
            dt.Columns.Add("stb");
            dt.Columns.Add("bcab");
            dt.Columns.Add("bcag");
            dt.Columns.Add("bcbb");
            dt.Columns.Add("bcbg");
            dt.Columns.Add("bccb");
            dt.Columns.Add("bccg");
            dt.Columns.Add("bcdb");
            dt.Columns.Add("bcdg");
            dt.Columns.Add("bceb");
            dt.Columns.Add("bceg");*/


            DataRow row;

            while (((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 1]).Value2 != null)
            {
                rowIndex = 2 + index;
                row = dt.NewRow();
                row[0] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 1]).Value2);
                row[1] = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 2]).Value2);
                row[2] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 3]).Value2);
                row[3] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 4]).Value2);
               /* row[4] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 5]).Value2);
                row[5] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 6]).Value2);
                row[6] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 7]).Value2);
                row[7] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 8]).Value2);
                row[8] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 9]).Value2);
                row[9] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 10]).Value2);
                row[10] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 11]).Value2);
                row[11] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 12]).Value2);
                row[12] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 13]).Value2);
                row[13] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 14]).Value2);
                row[14] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 15]).Value2);
                row[15] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 16]).Value2);
                row[16] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 17]).Value2);
                /*row[17] = Convert.ToInt32(((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[rowIndex, 18]).Value2);*/
                index++;
                dt.Rows.Add(row);
            }
            app.Workbooks.Close();

            expression = "College  = 'KSHATRIYA COLLEGE OF ENGINEERING'";
            DataRow[] foundRows;

            // Use the Select method to find all rows matching the filter.
            foundRows = dt.Select(expression);

            // Print column 0 of each returned row.
            string result = "";
            
            for (int i = 0; i < foundRows.Length; i++)
            {
                result += foundRows[i][1]+Environment.NewLine;
            }
            return result;
        }

        public string getRandom()
        {
            return "HELOOOFGHRKGHKE";
        }


        
    }
}