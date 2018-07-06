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
            var parametros = new Parametros();
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

        public void LimpiarSingleton()
        {
            var res = Singleton.Instance;
            res.Resultados = null;
            Index();
        }
        #endregion

        #region Post

        [HttpPost]
        public ActionResult Index(Parametros parametros)
        {
            switch (parametros.IdExcel)
            {
                case 1:
                    var res = Npoi(parametros);
                    return res;
                case 2:
                    return Epplus(parametros);
                default:
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

            var fsr = new FileStreamResult(ms, contentType) { FileDownloadName = fileDownloadName };

            return fsr;
        }

        private FileContentResult NpoiDownload(XSSFWorkbook excel)
        {
            using (var ms = new MemoryStream())
            {

                excel.Write(ms);

                const string fileDownloadName = "NPOI.xlsx";
                const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                var bytes = ms.ToArray();


                return File(bytes, contentType, fileDownloadName);
            }


        }

        private FileResult Npoi(Parametros parametros)
        {
            var res = Singleton.Instance;
            FileContentResult file = null;
            for (var i = 0; i < parametros.Iteraciones; i++)
            {
                var stopWatch = Stopwatch.StartNew();
                var informacion = new Consultas(parametros.Rows).GetExcelInformacion();
                var result = new Resultado
                {
                    Parametro = parametros,
                    Libreria = Leyendas.Npoi
                };

                if (informacion == null) continue;
                var watchCreation = Stopwatch.StartNew();

                var excel = parametros.Template ?
                    new NpoiService(Resources.DummyReport, informacion, parametros.Mascaras, parametros.Hojas).GetExcelExample() :
                    new NpoiService(informacion, parametros.Diseno, parametros.Mascaras, parametros.Hojas).GetExcelExample();


                watchCreation.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = Leyendas.Creacion,
                    Value = watchCreation.Elapsed.ToString()
                });

                if (parametros.Diseno)
                {
                    var watchDesign = Stopwatch.StartNew();
                    excel = new NpoiDesign(excel, parametros.Template).GetExcelExample();

                    watchDesign.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = Leyendas.Diseno,
                        Value = watchDesign.Elapsed.ToString()
                    });

                }
                var watchFiletoDonwload = Stopwatch.StartNew();
                file = NpoiDownload(excel);
                watchFiletoDonwload.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = Leyendas.Download,
                    Value = watchFiletoDonwload.Elapsed.ToString()
                });



                stopWatch.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = Leyendas.Total,
                    Value = stopWatch.Elapsed.ToString()
                });
                result.Intento = i;
                res.Resultados.Add(result);

                if (i != (parametros.Iteraciones - 1))
                    file = null;

            }

            return file;
        }

        private FileStreamResult Epplus(Parametros parametros)
        {
            var res = Singleton.Instance;
            FileStreamResult file = null;
            for (var i = 0; i < parametros.Iteraciones; i++)
            {
                var stopWatch = Stopwatch.StartNew();
                var informacion = new Consultas(parametros.Rows).GetExcelInformacion();

                var result = new Resultado
                {
                    Parametro = parametros,
                    Libreria = Leyendas.Epplus
                };

                if (informacion == null) continue;
                var watchCreation = Stopwatch.StartNew();

                var excel = parametros.Template ?
                    new EpplusServicio(Resources.DummyReport, informacion, parametros.Mascaras, parametros.Hojas).GetExcelExample() :
                    new EpplusServicio(informacion, parametros.Diseno, parametros.Mascaras, parametros.Hojas).GetExcelExample();

                watchCreation.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = Leyendas.Creacion,
                    Value = watchCreation.Elapsed.ToString()
                });


                if (parametros.Diseno)
                {
                    var watchDesign = Stopwatch.StartNew();
                    excel = new EpplusDesign(excel, parametros.Template).GetExcelExample();
                    watchDesign.Stop();
                    result.Tiempos.Add(new Tiempo
                    {
                        Descripcion = Leyendas.Diseno,
                        Value = watchDesign.Elapsed.ToString()
                    });
                }
                Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                file = EpplusDownload(excel);
                watchFiletoDonwload.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = Leyendas.Download,
                    Value = watchFiletoDonwload.Elapsed.ToString()
                });

                stopWatch.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = Leyendas.Total,
                    Value = stopWatch.Elapsed.ToString()
                });
                result.Intento = i;
                res.Resultados.Add(result);
                
                if (i != (parametros.Iteraciones - 1))
                    file = null;

            }


            return file;

        }

        private static class Leyendas
        {
            public const string Npoi = "NPOI";
            public const string Epplus = "EPPLUS";
            public const string Total = "Total";
            public const string Download = "File to download";
            public const string Diseno = "Diseno";
            public const string Creacion = "Creacion";
        }

        #endregion

        
    }
}