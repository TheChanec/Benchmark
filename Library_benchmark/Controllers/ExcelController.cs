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
    public class ExcelController : Controller
    {

        #region GET
        public ActionResult Index()
        {
            return View();
        }
        
        public ActionResult NPOI()
        {
            return View();
        }

        public ActionResult EPPLUS()
        {
            return View();
        }

        public ActionResult Tiempos()
        {
            return PartialView(Singleton.Instance.Resultados);
        }

        public FileResult DownloadSingleton()
        {
            Singleton res = Singleton.Instance;
            ExcelPackage excel;
            ExcelWorksheet currentsheet;

            excel = new ExcelPackage();
            currentsheet = excel.Workbook.Worksheets.Add("Result");

            ICollection<DummyView> respuesta = new List<DummyView>();

            respuesta = res
                .Resultados
                .Select(x => new DummyView
                {
                    Libreria = x.Libreria,
                    Registros = x.Parametro.Rows,
                    Sheet = x.Parametro.Sheets,
                    Recurso = x.Parametro.Resource,
                    TiempoCreacionDeExcel = x.Tiempos.Where(t => t.Descripcion == "Creacion").Select(t => t.Value).FirstOrDefault(),
                    TiempoDiseno = x.Tiempos.Where(t => t.Descripcion == "Diseno").Select(t => t.Value).FirstOrDefault(),
                    TiempoCreardescarga = x.Tiempos.Where(t => t.Descripcion == "File to download").Select(t => t.Value).FirstOrDefault(),
                    Total = x.Tiempos.Where(t => t.Descripcion == "Total").Select(t => t.Value).FirstOrDefault()
                })
                .ToList();

            currentsheet.Cells[1, 1].LoadFromCollection(respuesta, true );
            FileStreamResult file = EPPlusDownload(excel);


            return file;
        }

        #endregion

        #region Post
        [HttpPost]
        public FileResult NPOI(Parametros parametros)
        {
            Singleton res = Singleton.Instance;
            for (int i = 0; i < parametros.Iteraciones; i++)
            {
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
                        excel = new NPOIService(Resources.BookNPOI, informacion, parametros.Sheets).GetExcelExample();

                    else
                        excel = new NPOIService(informacion, parametros.Design, parametros.Sheets).GetExcelExample();


                    watchCreation.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = "Creacion",
                        Value = watchCreation.Elapsed.ToString()
                    });

                    if (parametros.Design)
                    {
                        Stopwatch watchDesign = Stopwatch.StartNew();
                        excel = new NPOIDesign(excel, parametros.Resource).GetExcelExample();

                        watchDesign.Stop();
                        result.Tiempos.Add(new Tiempo
                        {
                            Descripcion = "Diseno",
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
                    result.Intento = i;
                    res.Resultados.Add(result);

                    excel = null;
                    if (i != (parametros.Iteraciones - 1))
                        file = null;
                    else
                        return file;
                    
                }
            }
            
            return null;
        }

        [HttpPost]
        public ActionResult EPPLUS(Parametros parametros)
        {
            Singleton res = Singleton.Instance;
            FileStreamResult file;
            for (int i = 0; i < parametros.Iteraciones; i++)
            {
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
                        excel = new EPPLUSServicio(@"C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Resource/BookEPPLUS.xlsx", informacion, parametros.Sheets).GetExcelExample();
                    else
                        excel = new EPPLUSServicio(informacion, parametros.Design, parametros.Sheets).GetExcelExample();

                    watchCreation.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = "Creacion",
                        Value = watchCreation.Elapsed.ToString()
                    });


                    if (parametros.Design)
                    {
                        Stopwatch watchDesign = Stopwatch.StartNew();
                        excel = new EPPlusDesign(excel, parametros.Resource).GetExcelExample();
                        watchDesign.Stop();
                        result.Tiempos.Add(new Tiempo
                        {
                            Descripcion = "Diseno",
                            Value = watchDesign.Elapsed.ToString()
                        });
                    }
                    Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                    file = EPPlusDownload(excel);
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
                    result.Intento = i;
                    res.Resultados.Add(result);

                    excel = null;
                    if (i != (parametros.Iteraciones -1))
                        file = null;
                    else
                        return file;
                }

            }


            return null;
            
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

                string saveAsFileName = string.Format("{0:d}.xls", DateTime.Now).Replace("/", "-");

                byte[] bytes = exportData.ToArray();

                return File(bytes, "application/vnd.ms-excel", saveAsFileName);
            }
        }

        #endregion

    }
}