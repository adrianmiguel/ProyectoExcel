using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProyectoExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Borrado borrado = new Borrado();
            //borrado.IniciarBorrado();

            //CargarExcel excel = new CargarExcel();
            //excel.CrearExcel();

            //CrearExcel crearExcel = new CrearExcel();
            //crearExcel.CrearExcelXLS();

            LeerExcel LeerExcel = new LeerExcel();
            LeerExcel.LeerExcelXLSX();
        }
    }
}
