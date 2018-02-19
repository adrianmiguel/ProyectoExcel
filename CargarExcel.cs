using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ProyectoExcel
{
    class CargarExcel
    {
        public void CrearExcel()
        {
            ExcelPackage Excel = new ExcelPackage(new System.IO.FileInfo(@"C:\Users\adria\Documents\Visual Studio 2017\Projects\ProyectoExcel\Entradas_BB\ExcelEPPlus.xlsx"));
            Excel.Workbook.Worksheets.Add("Hoja1");
            Excel.Save();

            //return Excel;
        }
    }
}
