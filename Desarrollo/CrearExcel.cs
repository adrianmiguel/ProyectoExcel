using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Configuration;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using Security.Bpm.Ws; 

namespace Poliza
{
    class CrearExcel
    {
        public static string cadenaConexionBDMantiz { get; set; }

        public CrearExcel()
        {
            string userId = Cryptography.Decrypt(ConfigurationManager.AppSettings.Get("userId"), "BancoBogota");          //Usuario
            string key = Cryptography.Decrypt(ConfigurationManager.AppSettings.Get("key"), "BancoBogota");                //Clave
            string instancia = Cryptography.Decrypt(ConfigurationManager.AppSettings.Get("instancia"), "BancoBogota");    //Instancia SQL
            string dbMantiz = Cryptography.Decrypt(ConfigurationManager.AppSettings.Get("db"), "BancoBogota");            //Nombre Base de Datos
            cadenaConexionBDMantiz = "Server=" + instancia + ";Database=" + dbMantiz + ";User Id=" + userId + ";Password=" + key + ";";
        }

        public void CrearExcelXLS()
        {
            var appSettings = ConfigurationManager.AppSettings;

            string RutaSalidas_BB = appSettings["Ruta_Salidas_BB"];
            string NombreHojaExcel = appSettings["NombreHojaExcel"];

            DateTime FechaActual = DateTime.Now;

            string Dia = FechaActual.ToString("dd");
            string Mes = FechaActual.ToString("MM");
            string Anio = FechaActual.ToString("yyyy");
            string Ruta_Archivo = Path.Combine(RutaSalidas_BB, Dia + Mes + Anio + ".xls");

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
            }
            else
            {
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "ID";
                xlWorkSheet.Cells[1, 2] = "Name";
                xlWorkSheet.Cells[2, 1] = "1";
                xlWorkSheet.Cells[2, 2] = "One";
                xlWorkSheet.Cells[3, 1] = "2";
                xlWorkSheet.Cells[3, 2] = "Two";

                xlWorkBook.SaveAs(Ruta_Archivo, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }

        }

        public void CrearExcelXLSX()
        {
            var appSettings = ConfigurationManager.AppSettings;

            string RutaSalidas_BB = appSettings["Ruta_Salidas_BB"];
            string NombreHojaExcel = appSettings["NombreHojaExcel"];

            DateTime FechaActual = DateTime.Now;

            string Dia = FechaActual.ToString("dd");
            string Mes = FechaActual.ToString("MM");
            string Anio = FechaActual.ToString("yyyy");
            string Ruta_Archivo = Path.Combine(RutaSalidas_BB, Dia + Mes + Anio + ".xls");

            string xConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Ruta_Archivo + "; Extended Properties='Excel 8.0'";
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

        public void excel()
        {
            var appSettings = ConfigurationManager.AppSettings;

            string RutaSalidas_BB = appSettings["Ruta_Salidas_BB"];
            string NombreHojaExcel = appSettings["NombreHojaExcel"];

            DateTime FechaActual = DateTime.Now;

            string Dia = FechaActual.ToString("dd");
            string Mes = FechaActual.ToString("MM");
            string Anio = FechaActual.ToString("yyyy");
            string Ruta_Archivo = Path.Combine(RutaSalidas_BB, Dia + Mes + Anio + ".xls");

            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Sex", typeof(string));
            table.Columns.Add("Subject1", typeof(int));
            table.Columns.Add("Subject2", typeof(int));
            table.Columns.Add("Subject3", typeof(int));
            table.Columns.Add("Subject4", typeof(int));
            table.Columns.Add("Subject5", typeof(int));
            table.Columns.Add("Subject6", typeof(int));
            table.Rows.Add(1, "Amar", "M", 78, 59, 72, 95, 83, 77);
            table.Rows.Add(2, "Mohit", "M", 76, 65, 85, 87, 72, 90);
            table.Rows.Add(3, "Garima", "F", 77, 73, 83, 64, 86, 63);
            table.Rows.Add(4, "jyoti", "F", 55, 77, 85, 69, 70, 86);
            table.Rows.Add(5, "Avinash", "M", 87, 73, 69, 75, 67, 81);
            table.Rows.Add(6, "Devesh", "M", 92, 87, 78, 73, 75, 72);

            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE; 

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "StudentRepoertCard";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "Student Report Card";
                worKsheeT.Cells.Font.Size = 15;

                int rowcount = 2;

                foreach (DataRow datarow in table.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= table.Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = table.Columns[i - 1].ColumnName;
                            worKsheeT.Cells.Font.Color = System.Drawing.Color.Red;

                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == table.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, table.Columns.Count]];
                                }

                            }
                        }

                    }

                }

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, table.Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, table.Columns.Count]];

                worKbooK.SaveAs(Ruta_Archivo);
                worKbooK.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);

            }
            finally
            {
                worKsheeT = null;
                celLrangE = null;
                worKbooK = null;
            }  
        }

        public void ExcelV3()
        {
            var appSettings = ConfigurationManager.AppSettings;

            string RutaSalidas_BB = appSettings["Ruta_Salidas_BB"];
            string NombreHojaExcel = appSettings["NombreHojaExcel"];

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

            string strPath = Path.Combine(RutaSalidas_BB, "4.xlsx");
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);

            foreach (DataTable dtbl in ds.Tables)
            {
                //Create Excel WorkSheet
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add(Default, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], 1, Default);
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
                Excel.Range cellRang = excelWorkSheet.get_Range("A1", "G3");
                cellRang.Merge(false);
                cellRang.Interior.Color = System.Drawing.Color.White;
                cellRang.Font.Color = System.Drawing.Color.Gray;
                cellRang.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                cellRang.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                cellRang.Font.Size = 26;
                excelWorkSheet.Cells[1, 1] = "Greate Novels Of All Time";

                //Style table column names
                cellRang = excelWorkSheet.get_Range("A4", "G4");
                cellRang.Font.Bold = true;
                cellRang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                cellRang.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ED7D31");
                excelWorkSheet.get_Range("F4").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
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
            (excelWorkBook.Sheets[1] as Excel._Worksheet).Activate();

            excelWorkBook.SaveAs(strPath, Default, Default, Default, false, Default, Excel.XlSaveAsAccessMode.xlNoChange, Default, Default, Default, Default, Default);
            excelWorkBook.Close();
            excelApp.Quit();
        }

        public void ExcelV4()
        {
            var appSettings = ConfigurationManager.AppSettings;

            string RutaSalidas_BB = appSettings["Ruta_Salidas_BB"];
            string NombreHojaExcel = appSettings["NombreHojaExcel"];
            DataTable dt = new DataTable();

            DateTime FechaActual = DateTime.Now;

            string Dia = FechaActual.ToString("dd");
            string Mes = FechaActual.ToString("MM");
            string Anio = FechaActual.ToString("yyyy");
            string Ruta_Archivo = Path.Combine(RutaSalidas_BB, "Poliza_Inc_" + Dia + Mes + Anio + ".xlsx");

            ObtenerSolicitudes solicitudes = new ObtenerSolicitudes();
            dt = solicitudes.ConsultarSolicitudes(cadenaConexionBDMantiz, FechaActual);            

            dt.TableName = "Poliza";
            //dt.Columns.Add("Crédito");
            //dt.Columns.Add("Línea");
            //dt.Columns.Add("Tipo Documento");
            //dt.Columns.Add("No Documento");
            //dt.Columns.Add("Primer Apellido");
            //dt.Columns.Add("Segundo Apellido");
            //dt.Columns.Add("Primer Nombre");
            //dt.Columns.Add("Segundo Nombre");

            //dt.Columns.Add("Dirección");

            //dt.Columns.Add("Ciudad");
            //dt.Columns.Add("Departamento");
            //dt.Columns.Add("Telefono");
            //dt.Columns.Add("Fecha Desembolso");
            //dt.Columns.Add("Fecha Primer Pago");
            //dt.Columns.Add("Valor Asegurado");

            //dt.Columns.Add("Tipo Poliza");
            //dt.Columns.Add("Caso");
            //dt.Columns.Add("Valor Prima");
            //dt.Columns.Add("Periodo póliza");
            //dt.Columns.Add("Direcion Correspondenia");
            //dt.Columns.Add("Barrio");
            //dt.Columns.Add("Ciudad correspondencia");
            //dt.Columns.Add("Departamento de correspondencia");

            ////
            //dt.Columns.Add("NO. POLIZA");
            //dt.Columns.Add("ASEGURADO");
            //dt.Columns.Add("Nombre Arhivo Poliza");

            //dt.Rows.Add("359550834", "119", "CC", "79806071", "MARTINEZ", "AMADO", "RUBEN", "DARIO", "CALLE 14 359", "FUSAGASUGA", "CUNDINAMARCA", "517455555", "3/01/2018", "3/02/2018", "145260200", "PROPIA", "9310770", "44044", "MNSUAL", "CALLE 28 CR 39 - 77", "la guarda", "BOGOTA, DC", "BOGOTA");
            //dt.Rows.Add("359550834", "119", "CC", "79806071", "MARTINEZ", "AMADO", "RUBEN", "DARIO", "CALLE 14 359", "FUSAGASUGA", "CUNDINAMARCA", "517455555", "3/01/2018", "3/02/2018", "145260200", "PROPIA", "9310770", "44044", "MNSUAL", "CALLE 28 CR 39 - 77", "la guarda", "BOGOTA, DC", "BOGOTA1");
            //dt.Rows.Add("359550834", "119", "CC", "79806071", "MARTINEZ", "AMADO", "RUBEN", "DARIO", "CALLE 14 359", "FUSAGASUGA", "CUNDINAMARCA", "517455555", "3/01/2018", "3/02/2018", "145260200", "PROPIA", "9310770", "44044", "MNSUAL", "CALLE 28 CR 39 - 77", "la guarda", "BOGOTA, DC", "BOGOTA2");
            //dt.Rows.Add("359550834", "119", "CC", "79806071", "MARTINEZ", "AMADO", "RUBEN", "DARIO", "CALLE 14 359", "FUSAGASUGA", "CUNDINAMARCA", "517455555", "3/01/2018", "3/02/2018", "145260200", "PROPIA", "9310770", "44044", "MNSUAL", "CALLE 28 CR 39 - 77", "la guarda", "BOGOTA, DC", "BOGOTA3");
            //dt.Rows.Add("359550834", "119", "CC", "79806071", "MARTINEZ", "AMADO", "RUBEN", "DARIO", "CALLE 14 359", "FUSAGASUGA", "CUNDINAMARCA", "517455555", "3/01/2018", "3/02/2018", "145260200", "PROPIA", "9310770", "44044", "MNSUAL", "CALLE 28 CR 39 - 77", "la guarda", "BOGOTA, DC", "BOGOTA4");

            dt.Columns.Add("Numero POLIZA");
            dt.Columns.Add("ASEGURADO");
            dt.Columns.Add("Riesgo");
            dt.Columns.Add("Nombre Arhivo Poliza");

            int inColumn = 0, inRow = 0;

            System.Reflection.Missing Default = System.Reflection.Missing.Value;
            //string strPath = Path.Combine(RutaSalidas_BB, "5.xlsx");

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);

            //Create Excel WorkSheet
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add(Default, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], 1, Default);
           // excelWorkSheet.Name = dt.TableName;//"Poliza";//Name worksheet

            //Write Column Name
            for (int i = 0; i < dt.Columns.Count; i++)
                excelWorkSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;//.ToUpper();

            //Write Rows
            for (int m = 0; m < dt.Rows.Count; m++)
            {
                for (int n = 0; n < dt.Columns.Count; n++)
                {
                    inColumn = n + 1;
                    inRow = 2 + m;//1 + 2 + m;
                    excelWorkSheet.Cells[inRow, inColumn] = dt.Rows[m].ItemArray[n].ToString();
                    if (m % 2 == 0)
                        excelWorkSheet.get_Range("A" + inRow.ToString(), "Z" + inRow.ToString()).Interior.Color = System.Drawing.ColorTranslator.FromHtml("#D6EAF8");
                }
            }

            ////Excel Header
            //OfficeExcel.Range cellRang = excelWorkSheet.get_Range("A1", "O1");
            //cellRang.Merge(false);
            //cellRang.Interior.Color = System.Drawing.Color.Blue;
            //cellRang.Font.Color = System.Drawing.Color.Black;
            //cellRang.HorizontalAlignment = OfficeExcel.XlHAlign.xlHAlignCenter;
            //cellRang.VerticalAlignment = OfficeExcel.XlVAlign.xlVAlignCenter;
            //cellRang.Font.Size = 16;
            //excelWorkSheet.Cells[1, 1] = "Greate Novels Of All Time";

            //Style table column names
            Excel.Range cellRang = excelWorkSheet.get_Range("A1", "W1");
            cellRang.Font.Bold = true;
            cellRang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            cellRang.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#022C4D");
            cellRang.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            cellRang = excelWorkSheet.get_Range("X1", "Z1");
            cellRang.Font.Bold = true;
            cellRang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            cellRang.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#69C0BC");

            excelWorkSheet.get_Range("F4").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            //Formate price column
            excelWorkSheet.get_Range("O2").EntireColumn.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"; //.NumberFormat = "0.00";
            excelWorkSheet.get_Range("O2").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            //Auto fit columns
            excelWorkSheet.Columns.AutoFit();

            //Delete First Page
            excelApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Worksheet lastWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[1];
            lastWorkSheet.Delete();
            excelApp.DisplayAlerts = true;

            //Set Defualt Page
            (excelWorkBook.Sheets[1] as Excel._Worksheet).Activate();

            excelWorkBook.SaveAs(Ruta_Archivo, Default, Default, Default, false, Default, Excel.XlSaveAsAccessMode.xlNoChange, Default, Default, Default, Default, Default);
            excelWorkBook.Close();
            excelApp.Quit();
        }
    }
}
