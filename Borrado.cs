using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;

namespace ProyectoExcel
{
    class Borrado
    {
        public void IniciarBorrado()
        {
            //try
            //{
                var appSettings = ConfigurationManager.AppSettings;

                string RutaEntradas_BB = appSettings["Ruta_Entrdas_BB"];
                string RutaSalidas_BB = appSettings["Ruta_Salidas_BB"];
            //Console.WriteLine(RutaEntradas_BB);
            //}
            //catch (ConfigurationErrorsException)
            //{
            //    Console.WriteLine("Error leyendo parametros de onfiguración");
            //}

            try
            {
                DateTime FechaActual = DateTime.Now;
                Console.WriteLine(FechaActual);

                DateTime FechaBorrar = DateTime.Today.AddDays(-2);

                string Dia = FechaBorrar.ToString("dd");
                string Mes = FechaBorrar.ToString("MM");
                string Anio = FechaBorrar.ToString("yyyy");

                string Ruta_Borrar = Path.Combine(RutaSalidas_BB, Dia + Mes + Anio);

                Directory.Delete(Ruta_Borrar, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }

            
        }
        
    }
}
