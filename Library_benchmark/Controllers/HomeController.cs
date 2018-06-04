using Library_benchmark.Helpers;
using Library_benchmark.Models;
using Library_benchmark.Properties;
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
            return PartialView(Singleton.Instance.Resultados);
        }

        #endregion

        #region Post
        [HttpPost]
        public FileResult NPOIResult(Parametros parametros)
        {
            Singleton res = Singleton.Instance;
            Stopwatch stopWatch = Stopwatch.StartNew();
            IList<Dummy> informacion = new Consultas(parametros.Rows).GetInformacion();

            Resultado result = new Resultado();
            result.Parametro = parametros;
            result.Libreria = "NPOI";

            HSSFWorkbook excel;

            if (informacion != null)
            {
                Stopwatch watchCreation = Stopwatch.StartNew();

                if (parametros.Resource)
                    excel = new NPOIService(Resources.Book1, informacion, parametros.Sheets).GetExcelExample();

                else
                    excel = new NPOIService(informacion, parametros.Sheets).GetExcelExample();


                watchCreation.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = "Creation",
                    Value = watchCreation.Elapsed.ToString()
                });
                
                if (parametros.Design)
                {
                    Stopwatch watchDesign = Stopwatch.StartNew();
                    excel = new NPOIDesign(excel, true).GetExcelExample();
                    
                    watchDesign.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = "Design",
                        Value = watchDesign.Elapsed.ToString()
                    });

                }
                Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                FileContentResult file = NPOIdownload(excel);
                watchFiletoDonwload.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = "File to download",
                    Value = watchFiletoDonwload.Elapsed.ToString()
                });



                stopWatch.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = "Total",
                    Value = stopWatch.Elapsed.ToString()
                });
                res.Resultados.Add(result);
                return file;
            }
            return null;
        }

        [HttpPost]
        public ActionResult EPPLUSResult(Parametros parametros)
        {
            Singleton res = Singleton.Instance;
            Stopwatch stopWatch = Stopwatch.StartNew();
            IList<Dummy> informacion = new Consultas(parametros.Rows).GetInformacion();

            Resultado result = new Resultado();
            result.Parametro = parametros;
            result.Libreria = "EPPLUS";

            
            ExcelPackage excel;

            if (informacion != null)
            {
                Stopwatch watchCreation = Stopwatch.StartNew();

                if (parametros.Resource)
                    excel = new EPPLUSServicio(informacion, parametros.Sheets).GetExcelExample();
                else
                    excel = new EPPLUSServicio(informacion, parametros.Sheets).GetExcelExample();

                watchCreation.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = "Creation",
                    Value = watchCreation.Elapsed.ToString()
                });
                

                if (parametros.Design)
                {
                    Stopwatch watchDesign = Stopwatch.StartNew();
                    excel = new EPPlusDesign(excel, true).GetExcelExample();
                    watchDesign.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = "Design",
                        Value = watchDesign.Elapsed.ToString()
                    });
                }
                Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                FileStreamResult file = EPPlusDownload(excel);
                watchFiletoDonwload.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = "File to download",
                    Value = watchFiletoDonwload.Elapsed.ToString()
                });

                stopWatch.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = "Total",
                    Value = stopWatch.Elapsed.ToString()
                });
                res.Resultados.Add(result);
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