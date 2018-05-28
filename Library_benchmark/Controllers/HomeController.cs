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
        public ActionResult NPOIResult(Parametros parametros)
        {
            return View();
        }

        public ActionResult EPPLUSResult()
        {
            return View();
        }

        [HttpPost]
        public ActionResult EPPLUSResult(Parametros parametros)
        {
            return View();
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}