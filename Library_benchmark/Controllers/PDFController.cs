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
                    ITextSharp(parametros);
                    return null;
                case 2:
                    FastReport(parametros);
                    return null;
                default:
                    ViewBag.IdLibreria = new SelectList(parametros.PDFes, "Id", "Nombre");
                    return View(parametros);
            }
        }

        public ActionResult ITextSharp()
        {
            return PartialView();
        }

        public void ITextSharp(Parametros parametros)
        {
            var informacion = new Consultas().GetPdfInformacion();
            //if (informacion == null) return null;

            var pdf = new Document();

            for (var i = 0; i < parametros.Iteraciones;)
            {

                var workStream = new MemoryStream();
                PdfWriter.GetInstance(pdf, workStream).CloseStream = false;
                var strFilePath = Server.MapPath("~/PdfUploads/");

                var fileName = "Pdf_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".pdf";
                var doc = new Document(PageSize.A4, 5F, 5F, 73.5F, 70f);

                //var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(strFilePath + fileName, FileMode.Create));
                //pdfWriter.PageEvent = new ITextEvents();

                new TextSharpServicio(informacion, parametros.Template, doc, strFilePath, fileName);


                var contents = System.IO.File.ReadAllBytes(strFilePath + fileName);
                //return File(contents, "application/pdf", fileName);


            }



            //return null;
        }

        public void FastReport(Parametros parametros)
        {

            var res = Singleton.Instance;
            for (var i = 0; i < parametros.Iteraciones; i++)
            {
                var stopWatch = Stopwatch.StartNew();
                var informacion = new Consultas().GetPdfInformacion();

                var result = new Resultado
                {
                    Parametro = parametros,
                    //Libreria = Leyendas.Epplus
                };

                if (informacion == null) continue;
                var watchCreation = Stopwatch.StartNew();

                var pdf = parametros.Template ?
                    new FastReportServicio(Resources.PdfDummy, informacion).GetExcelExample() :
                    new FastReportServicio(informacion).GetExcelExample();

                watchCreation.Stop();
                result.Tiempos.Add(new Tiempo
                {
                    //Descripcion = Leyendas.Creacion,
                    Value = watchCreation.Elapsed.ToString()
                });
            }


            


            //return null;
        }


        

        

       
    }


}