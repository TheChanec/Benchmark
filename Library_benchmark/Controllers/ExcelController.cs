using Library_benchmark.Helpers;
using Library_benchmark.Helpers.EPPlus;
using Library_benchmark.Helpers.NPOI;
using Library_benchmark.Models;
using Library_benchmark.Properties;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
            var res = Singleton.Instance;

            var excel = new ExcelPackage();
            var currentsheet = excel.Workbook.Worksheets.Add("Result");

            ICollection<TimesView> respuesta = res
                .Resultados
                .Select(x => new TimesView
                {
                    Libreria = x.Libreria,
                    Registros = x.Parametro.Rows,
                    Sheet = x.Parametro.Hojas,
                    Recurso = x.Parametro.Template,
                    TiempoCreacionDeExcel = x.Tiempos.Where(t => t.Descripcion == "Creacion").Select(t => t.Value).FirstOrDefault(),
                    TiempoDiseno = x.Tiempos.Where(t => t.Descripcion == "Diseno").Select(t => t.Value).FirstOrDefault(),
                    TiempoCreardescarga = x.Tiempos.Where(t => t.Descripcion == "File to download").Select(t => t.Value).FirstOrDefault(),
                    Total = x.Tiempos.Where(t => t.Descripcion == "Total").Select(t => t.Value).FirstOrDefault()
                })
                .ToList();

            currentsheet.Cells[1, 1].LoadFromCollection(respuesta, true);
            var file = EpplusDownload(excel);


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
        private FileStreamResult EpplusDownload(ExcelPackage excel)
        {
            var ms = new MemoryStream();
            excel.SaveAs(ms);
            ms.Position = 0;

            const string fileDownloadName = "EPPLUS.xlsx";
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var fsr = new FileStreamResult(ms, contentType) {FileDownloadName = fileDownloadName};

            return fsr;
        }

        private FileContentResult NpoiDownload(XSSFWorkbook excel)
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
            var res = Singleton.Instance;
            FileContentResult file = null;
            for (int i = 0; i < parametros.Iteraciones; i++)
            {
                var stopWatch = Stopwatch.StartNew();
                var informacion = new Consultas(parametros.Rows).GetInformacion();

                var result = new Resultado
                {
                    Parametro = parametros,
                    Libreria = "NPOI"
                };

                XSSFWorkbook excel;

                if (informacion != null)
                {
                    Stopwatch watchCreation = Stopwatch.StartNew();

                    if (parametros.Template)
                        excel = new NpoiService(Resources.DummyReport, informacion, parametros.Mascaras, parametros.Hojas).GetExcelExample();

                    else
                        excel = new NpoiService(informacion, parametros.Diseno, parametros.Mascaras, parametros.Hojas).GetExcelExample();


                    watchCreation.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = "Creacion",
                        Value = watchCreation.Elapsed.ToString()
                    });

                    if (parametros.Diseno)
                    {
                        Stopwatch watchDesign = Stopwatch.StartNew();
                        excel = new NpoiDesign(excel, parametros.Template).GetExcelExample();

                        watchDesign.Stop();
                        result.Tiempos.Add(new Tiempo
                        {
                            Descripcion = "Diseno",
                            Value = watchDesign.Elapsed.ToString()
                        });

                    }
                    Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                    file = NpoiDownload(excel);
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
            var res = Singleton.Instance;
            FileStreamResult file = null;
            for (var i = 0; i < parametros.Iteraciones; i++)
            {
                var stopWatch = Stopwatch.StartNew();
                var informacion = new Consultas(parametros.Rows).GetInformacion();

                var result = new Resultado
                {
                    Parametro = parametros,
                    Libreria = "EPPLUS"
                };

                ExcelPackage excel;

                if (informacion == null) continue;
                var watchCreation = Stopwatch.StartNew();

                excel = parametros.Template ? 
                    new EpplusServicio(Resources.DummyReport, informacion, parametros.Mascaras, parametros.Hojas).GetExcelExample() : 
                    new EpplusServicio(informacion, parametros.Diseno, parametros.Mascaras, parametros.Hojas).GetExcelExample();

                watchCreation.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = "Creacion",
                    Value = watchCreation.Elapsed.ToString()
                });


                if (parametros.Diseno)
                {
                    Stopwatch watchDesign = Stopwatch.StartNew();
                    excel = new EpplusDesign(excel, parametros.Template).GetExcelExample();
                    watchDesign.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = "Diseno",
                        Value = watchDesign.Elapsed.ToString()
                    });
                }
                Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                file = EpplusDownload(excel);
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


            return file;

        }
        #endregion

    }
}