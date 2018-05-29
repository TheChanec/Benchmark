using Library_benchmark.Helpers;
using Library_benchmark.Models;
using System;
using System.Collections.Generic;
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
            var cabeceras = new Consultas().Cabeceras();
            var informacion = new Consultas().Informacion();
            
            var excel = new EPPLUSServicio(cabeceras, informacion);
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