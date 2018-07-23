//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using OfficeExcel = Microsoft.Office.Interop.Excel;


namespace SolicitudesConstructor.Logica
{
    class LeerExcelForma1
    {
        //Application _excelApp;

        //public LeerExcelForma1(Application excelApp)
        //{
        //    //_excelApp = new Application();
        //}

        public DataSet LeerExcel()
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            try
            {                
                string path = @"C:\Users\adria\Documents\Visual Studio 2017\Projects\SolicitudesConstructor\Archivos\ArchivoXlsx.xlsx";

                ds.DataSetName = "Solicitudes";

                OfficeExcel.Application application = new OfficeExcel.Application();
                application.DisplayAlerts = false;

                OfficeExcel.Workbook workbook = application.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //This opens the file

                OfficeExcel.Worksheet sheet = (OfficeExcel.Worksheet)workbook.Sheets.get_Item("Solicitudes"); //Get the first sheet in the file

                int lastRow = sheet.Cells.SpecialCells(OfficeExcel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                int lastColumn = sheet.Cells.SpecialCells(OfficeExcel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

                OfficeExcel.Range oRange = sheet.UsedRange;//sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[lastRow, lastColumn]);//("A1",lastColumnIndex + lastRow.ToString());

                oRange.EntireColumn.AutoFit();
               
                object[,] cellValues = (object[,])oRange.Value2;
                object[] values = new object[lastColumn];

                bool AgregarRows = false;

                for (int i = 1; i <= lastRow; i++)
                {

                    for (int j = 0; j < lastColumn; j++)
                    {
                        if (i == 1)
                        {
                            //dt.Columns.Add("a" + i.ToString());
                            dt.Columns.Add(cellValues[i, j + 1].ToString());
                        }
                        else
                        {
                            values[j] = cellValues[i, j + 1];
                            AgregarRows = true;
                        }

                        
                    }
                    if (AgregarRows)
                    {
                        dt.Rows.Add(values);
                    }
                    
                }
                dt.TableName = "Table";
                ds.Tables.Add(dt);

                workbook.Close(false, Type.Missing, Type.Missing);
                application.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                //if ()
                //{

                //}
            }

            return ds;
        }

        public void LeerExcel2()
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            string path = @"C:\Users\adria\Documents\Visual Studio 2017\Projects\SolicitudesConstructor\Archivos\ArchivoXlsx.xlsx";

            OfficeExcel.Application application = new OfficeExcel.Application();
            application.DisplayAlerts = false;

            OfficeExcel.Workbook workbook = application.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //This opens the file

            OfficeExcel.Worksheet sheet = (OfficeExcel.Worksheet)workbook.Sheets.get_Item("Solicitudes"); //Get the first sheet in the file

            int lastRow = sheet.Cells.SpecialCells(OfficeExcel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            int lastColumn = sheet.Cells.SpecialCells(OfficeExcel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

            OfficeExcel.Range oRange = sheet.UsedRange;//sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[lastRow, lastColumn]);//("A1",lastColumnIndex + lastRow.ToString());

            int rowCount = oRange.Rows.Count;
            int colCount = oRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (oRange.Cells[i, j] != null && oRange.Cells[i, j].Value2 != null)
                        Console.Write(oRange.Cells[i, j].Value2.ToString() + "\t");
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            //Marshal.ReleaseComObject(oRange);
            //Marshal.ReleaseComObject(workbook);

            //close and release
            //workbook.Close();
            workbook.Close(false, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(workbook);

            //quit and release
            application.Quit();
            Marshal.ReleaseComObject(application);
        }
    }
}
