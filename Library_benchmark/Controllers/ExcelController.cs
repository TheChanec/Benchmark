using Library_benchmark.Helpers;
using Library_benchmark.Models;
using Library_benchmark.Properties;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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
            Parametros parametros = new Parametros();
            ViewBag.IdLibreria = new SelectList(parametros.Exceles, "Id", "Nombre");

            return View(parametros);
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

            currentsheet.Cells[1, 1].LoadFromCollection(respuesta, true);
            FileStreamResult file = EPPlusDownload(excel);


            return file;
        }

        #endregion

        #region Post

        [HttpPost]
        public ActionResult Index(Parametros parametros)
        {

            if (parametros.IdExcel == 1)
            {
                var res = NPOI(parametros);
                return res;
            }
            else if (parametros.IdExcel == 2)
            {
                return EPPLUS(parametros);
            }
            else
            {

                ViewBag.IdLibreria = new SelectList(parametros.Exceles, "Id", "Nombre");

                return View(parametros);
            }


        }


        #endregion

        #region Helpers
        private FileStreamResult EPPlusDownload(ExcelPackage excel)
        {
            var ms = new MemoryStream();
            excel.SaveAs(ms);
            ms.Position = 0;


            var fileDownloadName = "EPPLUS.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var fsr = new FileStreamResult(ms, contentType);
            fsr.FileDownloadName = fileDownloadName;

            return fsr;
        }

        private FileContentResult NPOIdownload(XSSFWorkbook excel)
        {
            using (MemoryStream ms = new MemoryStream())
            {

                excel.Write(ms);

                var fileDownloadName = "NPOI.xlsx";
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                byte[] bytes = ms.ToArray();


                return File(bytes, contentType, fileDownloadName);
            }


        }

        private FileResult NPOI(Parametros parametros)
        {
            Singleton res = Singleton.Instance;
            FileContentResult file = null;
            for (int i = 0; i < parametros.Iteraciones; i++)
            {
                Stopwatch stopWatch = Stopwatch.StartNew();
                IList<Dummy> informacion = new Consultas(parametros.Rows).GetInformacion();

                Resultado result = new Resultado();
                result.Parametro = parametros;
                result.Libreria = "NPOI";

                XSSFWorkbook excel;

                if (informacion != null)
                {
                    Stopwatch watchCreation = Stopwatch.StartNew();

                    if (parametros.Resource)
                        excel = new NPOIService(Resources.DummyReport, informacion, parametros.Mascaras, parametros.Sheets).GetExcelExample();

                    else
                        excel = new NPOIService(informacion, parametros.Design, parametros.Mascaras, parametros.Sheets).GetExcelExample();


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
                    file = NPOIdownload(excel);
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


                }
                else
                {
                    return null;
                }
            }

            return file;
        }

        private FileStreamResult EPPLUS(Parametros parametros)
        {
            Singleton res = Singleton.Instance;
            FileStreamResult file = null;
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
                        excel = new EPPLUSServicio(Resources.DummyReport, informacion, parametros.Mascaras, parametros.Sheets).GetExcelExample();
                    else
                        excel = new EPPLUSServicio(informacion, parametros.Design, parametros.Mascaras, parametros.Sheets).GetExcelExample();

                    watchCreation.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = "Creacion",
                        Value = watchCreation.Elapsed.ToString()
                    });


                    if (parametros.Design)
                    {
                        Stopwatch watchDesign = Stopwatch.StartNew();
                        excel = new EPPlusDesign(excel, parametros.Resource, null).GetExcelExample();
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
                    if (i != (parametros.Iteraciones - 1))
                        file = null;

                }

            }


            return file;

        }
        #endregion

    }
}