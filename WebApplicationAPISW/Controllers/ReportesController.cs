using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApplicationAPISW.Models;
using WebApplicationAPISW.Tools;

namespace WebApplicationAPISW.Controllers
{
    public class ReportesController : Controller
    {
        // GET: Reportes
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public  ActionResult GenerarExcelReporte()
        {
            //Datos Dummy
            List<Empleado> lista = new Empleado().GetDataDummy();

            if (lista != null)
            {
                string fileName = $"ReporteUsers.xlsx";
                string sheetName = "Reporte";
                List<Tuple<int, string>> formatos = new List<Tuple<int, string>>();
                formatos.Add(new Tuple<int, string>(5, "$#,##0.00"));

                byte[] resultado = lista.GenerarExcel(sheetName, formatos);
                return File((byte[])resultado, MimeMapping.GetMimeMapping(fileName), fileName);
            }
            else
            {
                return null;
            }
            
        }

    }
}