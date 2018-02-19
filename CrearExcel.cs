using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;

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
    }
}
