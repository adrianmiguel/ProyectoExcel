using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Configuration;
using System.IO;
using System.Data.OleDb;
using System.Data; 

namespace Poliza
{
    class LeerExcel
    {
        public void CargarExcel()
        {

            var appSettings = ConfigurationManager.AppSettings;

            string RutaEntradas_BB = appSettings["Ruta_Entradas_BB"];
            string ExtensionArchivoEntrada = appSettings["ExtensionExcelEntrada"];

            DateTime FechaActual = DateTime.Now;
            string Dia = FechaActual.ToString("dd");
            string Mes = FechaActual.ToString("MM");
            string Anio = FechaActual.ToString("yyyy");
            string Ruta_Archivo = "";
            string CadenaConexionArchivoExcel = "";

            if (ExtensionArchivoEntrada == ".xls")
            {
                Ruta_Archivo = Path.Combine(RutaEntradas_BB, Dia + Mes + Anio, "Poliza_Inc_" + Dia + Mes + Anio + ".xls");
                CadenaConexionArchivoExcel = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Ruta_Archivo + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            }
            else if (ExtensionArchivoEntrada == ".xlsx")
            {
                Ruta_Archivo = Path.Combine(RutaEntradas_BB, Dia + Mes + Anio, "Poliza_Inc_" + Dia + Mes + Anio + ".xlsx");
                CadenaConexionArchivoExcel = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Ruta_Archivo + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }

            string Query = "SELECT Caso, [Numero POLIZA], ASEGURADO, Riesgo, [Nombre Arhivo Poliza] FROM [Hoja2$]";
            //string Query = "SELECT * FROM [Hoja1$]";
            OleDbConnection con = new OleDbConnection(CadenaConexionArchivoExcel);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            OleDbCommand cmd = new OleDbCommand(Query, con);
            OleDbDataAdapter Adaptador = new OleDbDataAdapter();
            Adaptador.SelectCommand = cmd;            

            DataTable dt = new DataTable();
            Adaptador.Fill(dt);
            Adaptador.Dispose();
            con.Close();
            con.Dispose();

            foreach (DataRow dr in dt.Rows)
            {
                string Solicitud = dr["Caso"].ToString();
                string Poliza = dr["Numero POLIZA"].ToString();
                string Asegurado = dr["ASEGURADO"].ToString();
                string Riesgo = dr["Riesgo"].ToString();
                string NombreArchivo = dr["Nombre Arhivo Poliza"].ToString();

                string RutaPolizas = Path.Combine(RutaEntradas_BB, Dia + Mes + Anio, NombreArchivo);

            }

            //DataSet ds = new DataSet();
            //Adaptador.Fill(ds);
            //Adaptador.Dispose();
            //con.Close();
            //con.Dispose();

            //foreach (DataRow dr in ds.Tables[0].Rows)
            //{
            //    string Solicitud = dr["Caso"].ToString();
            //    string Poliza = dr["Numero POLIZA"].ToString();
            //    string Asegurado = dr["ASEGURADO"].ToString();
            //    string NombreArchivo = dr["Nombre Arhivo Poliza"].ToString();
            //}
        }
    }
}
