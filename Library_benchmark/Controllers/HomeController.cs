using Library_benchmark.Helpers;
using Library_benchmark.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Library_benchmark.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult NPOIResult()
        {
            return View();
        }

        [HttpPost]
        public JsonResult NPOIResult(Parametros parametros)
        {
            return Json(new { }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult EPPLUSResult()
        {
            return View();
        }

        [HttpPost]
        public ActionResult EPPLUSResult(Parametros parametros)
        {
            IList<Elemento1> informacion = new Consultas(parametros.Rows).GetInformacion();
            if (informacion != null) {
                var excel = new EPPLUSServicio(informacion, parametros.Sheets).GetExcelExample();
                var fileStream = new MemoryStream();
                excel.SaveAs(fileStream);
                fileStream.Position = 0;


                var fileDownloadName = "sample.xlsx";
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                var fsr = new FileStreamResult(fileStream, contentType);
                fsr.FileDownloadName = fileDownloadName;

                return fsr;
            }
            return PartialView("About");
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return PartialView();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}