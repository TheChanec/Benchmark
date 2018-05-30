using Library_benchmark.Helpers;
using Library_benchmark.Models;
using NPOI.HSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Library_benchmark.Controllers
{
    public class HomeController : Controller
    {

        #region GET
        public ActionResult Index()
        {
            return View();
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

        public ActionResult NPOIResult()
        {
            return View();
        }

        public ActionResult EPPLUSResult()
        {
            return View();
        }

        public ActionResult Tiempos()
        {
            return PartialView(Resultado.Instance);
        }

        #endregion

        #region Post
        [HttpPost]
        public FileResult NPOIResult(Parametros parametros)
        {
            IList<Elemento1> informacion = new Consultas(parametros.Rows).GetInformacion();


            if (informacion != null)
            {
                Resultado res = Resultado.Instance;
                Stopwatch stopWatch = Stopwatch.StartNew();
                FileContentResult file;

                var excel = new NPOIService(informacion, parametros.Sheets).GetExcelExample();
                if (parametros.Design)
                {
                    var design = new NPOIDesign(excel).GetExcelExample();
                    file = NPOIdownload(design);
                }
                else {
                    file = NPOIdownload(excel);
                }
                

                stopWatch.Stop();
                res.Tiempos.Add(new Tiempo { Parametro = parametros, Libreria = "NPOI", TiempoDeEjecucion = stopWatch.Elapsed.ToString() });
                return file;
            }
            return null;
        }

       

        [HttpPost]
        public ActionResult EPPLUSResult(Parametros parametros)
        {
            IList<Elemento1> informacion = new Consultas(parametros.Rows).GetInformacion();
            if (informacion != null)
            {

                Resultado res = Resultado.Instance;

                Stopwatch stopWatch = Stopwatch.StartNew();

                FileStreamResult file;
                var excel = new EPPLUSServicio(informacion, parametros.Sheets).GetExcelExample();
                if (parametros.Design)
                {
                    var design = new EPPlusDesign(excel, 1, true, "", 12).GetExcelExample();
                    file = EPPlusDownload(design);
                }
                else
                    file = EPPlusDownload(excel);



                stopWatch.Stop();
                res.Tiempos.Add(new Tiempo { Parametro = parametros, Libreria = "EPPlus", TiempoDeEjecucion = stopWatch.Elapsed.ToString() });

                return file;
            }
            return PartialView("About");
        }
        #endregion

        #region Helpers
        private FileStreamResult EPPlusDownload(ExcelPackage excel)
        {
            var fileStream = new MemoryStream();
            excel.SaveAs(fileStream);
            fileStream.Position = 0;


            var fileDownloadName = "sample.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var fsr = new FileStreamResult(fileStream, contentType);
            fsr.FileDownloadName = fileDownloadName;

            return fsr;
        }

        private FileContentResult NPOIdownload(HSSFWorkbook excel)
        {
            using (var exportData = new MemoryStream())
            {

                excel.Write(exportData);

                string saveAsFileName = string.Format("MembershipExport-{0:d}.xls", DateTime.Now).Replace("/", "-");

                byte[] bytes = exportData.ToArray();
                
                return File(bytes, "application/vnd.ms-excel", saveAsFileName);
            }
        }

        #endregion

    }
}