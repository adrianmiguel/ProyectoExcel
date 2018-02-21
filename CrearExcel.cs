using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using OfficeExcel = Microsoft.Office.Interop.Excel;

namespace ProyectoExcel
{
    class CrearExcel
    {
        public void CrearExcelXLS()
        {
            var appSettings = ConfigurationManager.AppSettings;

            string RutaEntradas_BB = appSettings["Ruta_Entrdas_BB"];
            string NombreHojaExcel = appSettings["NombreHojaExcel"];

            DateTime FechaActual = DateTime.Now;

            string Dia = FechaActual.ToString("dd");
            string Mes = FechaActual.ToString("MM");
            string Anio = FechaActual.ToString("yyyy");
            string Ruta_Archivo = Path.Combine(RutaEntradas_BB, Dia + Mes + Anio + ".xls");

            string xConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Ruta_Archivo + "; Extended Properties='Excel 8.0;HDR=YES'";
            var conn = new OleDbConnection(xConnStr);
            string ColumnName = "[columename] varchar(255)"; 
            conn.Open();
            var cmd = new OleDbCommand("CREATE TABLE [" + NombreHojaExcel + "] (F1 number, F2 char(255), F3 char(128))", conn);
            cmd.ExecuteNonQuery();

            //now we insert the values into the existing sheet...no new sheet is added.
            cmd.CommandText = "INSERT INTO [" + NombreHojaExcel + "] (F1, F2, F3) VALUES(4,\"Tampa\",\"Florida\")";
            cmd.ExecuteNonQuery();

            //insert another row into the sheet...
            cmd.CommandText = "INSERT INTO [" + NombreHojaExcel + "] (F1, F2, F3) VALUES(5,\"Pittsburgh\",\"Pennsylvania\")";
            cmd.ExecuteNonQuery();

            conn.Close();
        }

        public void CrearExcelXLSX()
        {

        }

        public void CrearExcel3()
        {           
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            dt.Columns.Add("Sl No");//Define Columns
            dt.Columns.Add("Novel Name");
            dt.Columns.Add("Author");
            dt.Columns.Add("Genres");
            dt.Columns.Add("Published Date");
            dt.Columns.Add("Price");
            dt.Columns.Add("Rating");

            dt.Rows.Add("1", "In Search of Lost Time", "Marcel Proust", "Literary modernism", "01-01-1913", "348", "4.3");//Adding Rows
            dt.Rows.Add("2", "Ulysses", "James Joyce", "Modernism", "22-02-1922", "58", "3.7");
            dt.Rows.Add("3", "Moby Dick", "Herman Melville", "Adventure fiction", "18-10-1851", "131", "3.4");
            dt.Rows.Add("4", "Hamlet", "William Shakespeare", "Tragedy", "01-01-1603", "225", "3.9");
            dt.Rows.Add("5", "War and Peace", "Leo Tolstoy", "Historical fiction", "01-01-1869", "133.95", "4.1");
            dt.TableName = "Table1";

            DataTable dtbl2 = dt.Copy();//Created copies of first table
            dtbl2.TableName = "Table2";
            ds.Tables.Add(dtbl2);
            DataTable dtbl3 = dt.Copy();//Created copies of first table
            dtbl3.TableName = "Table3";
            ds.Tables.Add(dtbl3);

            int inHeaderLength = 3, inColumn = 0, inRow = 0;
            System.Reflection.Missing Default = System.Reflection.Missing.Value;

            string strPath = @"C:\Users\adria\Documents\Visual Studio 2017\Projects\ProyectoExcel\Salidas_BB\" + "2.xls";
            OfficeExcel.Application excelApp = new OfficeExcel.Application();
            OfficeExcel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);

            foreach (DataTable dtbl in ds.Tables)
            {
                //Create Excel WorkSheet
                OfficeExcel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add(Default, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], 1, Default);
                excelWorkSheet.Name = dtbl.TableName;//Name worksheet

                //Write Column Name
                for (int i = 0; i < dtbl.Columns.Count; i++)
                    excelWorkSheet.Cells[inHeaderLength + 1, i + 1] = dtbl.Columns[i].ColumnName.ToUpper();

                //Write Rows
                for (int m = 0; m < dtbl.Rows.Count; m++)
                {
                    for (int n = 0; n < dtbl.Columns.Count; n++)
                    {
                        inColumn = n + 1;
                        inRow = inHeaderLength + 2 + m;
                        excelWorkSheet.Cells[inRow, inColumn] = dtbl.Rows[m].ItemArray[n].ToString();
                        if (m % 2 == 0)
                            excelWorkSheet.get_Range("A" + inRow.ToString(), "G" + inRow.ToString()).Interior.Color = System.Drawing.ColorTranslator.FromHtml("#FCE4D6");
                    }
                }

                //Excel Header
                OfficeExcel.Range cellRang = excelWorkSheet.get_Range("A1", "G3");
                cellRang.Merge(false);
                cellRang.Interior.Color = System.Drawing.Color.White;
                cellRang.Font.Color = System.Drawing.Color.Gray;
                cellRang.HorizontalAlignment = OfficeExcel.XlHAlign.xlHAlignCenter;
                cellRang.VerticalAlignment = OfficeExcel.XlVAlign.xlVAlignCenter;
                cellRang.Font.Size = 26;
                excelWorkSheet.Cells[1, 1] = "Greate Novels Of All Time";

                //Style table column names
                cellRang = excelWorkSheet.get_Range("A4", "G4");
                cellRang.Font.Bold = true;
                cellRang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                cellRang.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ED7D31");
                excelWorkSheet.get_Range("F4").EntireColumn.HorizontalAlignment = OfficeExcel.XlHAlign.xlHAlignRight;
                //Formate price column
                excelWorkSheet.get_Range("F5").EntireColumn.NumberFormat = "0.00";
                //Auto fit columns
                excelWorkSheet.Columns.AutoFit();
            }

            //Delete First Page
            excelApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Worksheet lastWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[1];
            lastWorkSheet.Delete();
            excelApp.DisplayAlerts = true;

            //Set Defualt Page
            (excelWorkBook.Sheets[1] as OfficeExcel._Worksheet).Activate();

            excelWorkBook.SaveAs(strPath, OfficeExcel.XlFileFormat.xlExcel8, Default, Default, false, Default, OfficeExcel.XlSaveAsAccessMode.xlNoChange, Default, Default, Default, Default, Default);
            excelWorkBook.Close();
            excelApp.Quit();
        }
    }
}
