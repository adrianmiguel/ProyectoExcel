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
    class LeerExcel
    {
        public void LeerExcelXLS()
        {
            var appSettings = ConfigurationManager.AppSettings;

            string RutaEntradas_BB = appSettings["Ruta_Entrdas_BB"];

            DateTime FechaActual = DateTime.Now;

            string Dia = FechaActual.ToString("dd");
            string Mes = FechaActual.ToString("MM");
            string Anio = FechaActual.ToString("yyyy");

            string Ruta_Archivo = Path.Combine(RutaEntradas_BB, Dia + Mes + Anio, Dia + Mes + Anio + ".xls");

            string CadenaConexionArchivoExcel = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Ruta_Archivo + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";

            OleDbConnection connection = new OleDbConnection();
            connection = new OleDbConnection(CadenaConexionArchivoExcel);
            connection.Open();

            OleDbCommand consulta = default(OleDbCommand);
            string Query = "SELECT * FROM [Hoja1$]";
            consulta = new OleDbCommand(Query, connection);

            OleDbDataAdapter Adaptador = new OleDbDataAdapter();
            Adaptador.SelectCommand = consulta;

            DataSet ds = new DataSet();
            Adaptador.Fill(ds);

            connection.Close();
        }
        public void LeerExcelXLSX()
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
                Ruta_Archivo = Path.Combine(RutaEntradas_BB, Dia + Mes + Anio, Dia + Mes + Anio + ".xls");
                CadenaConexionArchivoExcel = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Ruta_Archivo + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
            }
            else if (ExtensionArchivoEntrada == ".xlsx")
            {
                Ruta_Archivo = Path.Combine(RutaEntradas_BB, Dia + Mes + Anio, Dia + Mes + Anio + ".xlsx");
                CadenaConexionArchivoExcel = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Ruta_Archivo + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }

            //string Query = "select [Empleado Id], [Nombre Compañia], [Contacto], [Telefono] from [Hoja1$]";
            string Query = "SELECT * FROM [Hoja1$]";
            OleDbConnection con = new OleDbConnection(CadenaConexionArchivoExcel);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            OleDbCommand cmd = new OleDbCommand(Query, con);
            OleDbDataAdapter Adaptador = new OleDbDataAdapter();
            Adaptador.SelectCommand = cmd;

            DataSet ds = new DataSet();
            Adaptador.Fill(ds);
            Adaptador.Dispose();
            con.Close();
            con.Dispose();

            //using (MuDatabaseEnties dc = new MuDatabaseEnties)
            //{
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                string Emp = dr["Empleado Id"].ToString();
            }

            //OleDbConnection connection = new OleDbConnection();
            //try
            //{

            //    connection = new OleDbConnection(CadenaConexionArchivoExcel);
            //    connection.Open();
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //}


            //OleDbCommand consulta = default(OleDbCommand);
            //string Query = "SELECT * FROM [Hoja1$]";
            //consulta = new OleDbCommand(Query, connection);

            //OleDbDataAdapter Adaptador = new OleDbDataAdapter();
            //Adaptador.SelectCommand = consulta;

            //DataSet ds = new DataSet();
            //Adaptador.Fill(ds);

            //foreach (DataRow dr in ds.Tables[0].Rows)
            //{
            //    string Emp = dr["F1"].ToString();
            //}

        }
    }
}
