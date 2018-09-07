using iTextSharp.text;
using iTextSharp.text.pdf;
using Library_benchmark.Helpers;
using Library_benchmark.Models;
using System;
using System.IO;
using System.Web.Mvc;
using Library_benchmark.Helpers.IText;
using System.Diagnostics;
using Library_benchmark.Helpers.Fast;
using Library_benchmark.Properties;
using FastReport;

namespace Library_benchmark.Controllers
{
    public class PDFController : Controller
    {
        // GET: PDF
        public ActionResult Index()
        {
            var parametros = new Parametros();
            ViewBag.IdLibreria = new SelectList(parametros.PDFes, "Id", "Nombre");
            return View(parametros);
        }

        public ActionResult Tiempos()
        {
            return PartialView(Singleton.Instance.Resultados);
        }

        [HttpPost]
        public ActionResult Index(Parametros parametros)
        {
            switch (parametros.IdPdf)
            {
                case 1:
                    return ITextSharp(parametros);
                case 2:
                    return FastReport(parametros);

                default:
                    ViewBag.IdLibreria = new SelectList(parametros.PDFes, "Id", "Nombre");
                    return View(parametros);
            }
        }


        public FileStreamResult ITextSharp(Parametros parametros)
        {
            var informacion = new Consultas().GetPdfInformacion();
            //if (informacion == null) return null;

            FileStreamResult output = null;
            MemoryStream ms = new MemoryStream();
            for (var i = 0; i < parametros.Iteraciones; i++)
            {

                var file = new TextSharpServicio(informacion, parametros.Template).GetFile();

                output = ItextSharpDownload(file);
            }



            return output;
        }

        public FileStreamResult FastReport(Parametros parametros)
        {
            FileStreamResult file = null;
            var res = Singleton.Instance;
            for (var i = 0; i < parametros.Iteraciones; i++)
            {
                var stopWatch = Stopwatch.StartNew();
                var informacion = new Consultas().GetPdfInformacion();

                var result = new Resultado
                {
                    Parametro = parametros,
                    Libreria = Leyendas.Fast
                };

                if (informacion == null) continue;
                var watchCreation = Stopwatch.StartNew();

                var pdf = parametros.Template ?
                    new FastReportServicio(Resources.PdfDummy, informacion).GetExcelExample() :
                    new FastReportServicio(informacion).GetExcelExample();

                watchCreation.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    Descripcion = Leyendas.Creacion,
                    Value = watchCreation.Elapsed.ToString()
                });

                Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                file = FastReportDownload(pdf);
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
        private FileStreamResult ItextSharpDownload(byte[] file)
        {
            MemoryStream output = new MemoryStream();
            output.Write(file, 0, file.Length);
            output.Position = 0;

            HttpContext.Response.AddHeader("content-disposition", "attachment; filename=form.pdf");


            // Return the output stream
            return File(output, "application/pdf");
        }

        private FileStreamResult FastReportDownload(Report pdf)
        {
            if (pdf.Report.Prepare())
            {
                // Set PDF export props
                FastReport.Export.Pdf.PDFExport pdfExport = new FastReport.Export.Pdf.PDFExport();
                pdfExport.ShowProgress = false;
                pdfExport.Subject = "Subject";
                pdfExport.Title = "xxxxxxx";
                pdfExport.Compressed = true;
                pdfExport.AllowPrint = true;
                pdfExport.EmbeddingFonts = true;

                MemoryStream strm = new MemoryStream();
                pdf.Report.Export(pdfExport, strm);
                pdf.Dispose();
                pdfExport.Dispose();
                strm.Position = 0;

                // return stream in browser
                return File(strm, "application/pdf", "report.pdf");
            }
            else
            {
                return null;
            }
        }



        private static class Leyendas
        {
            public const string Fast = "FastReport";
            public const string ITextSharp = "TextSharp";
            public const string Total = "Total";
            public const string Download = "File to download";
            public const string Creacion = "Creacion";
        }

    }


}