
using iTextSharp.text;
using iTextSharp.text.pdf;
using Library_benchmark.Helpers;
using Library_benchmark.Helpers.ITextSharp;
using Library_benchmark.Models;
using System;
using System.Diagnostics;
using System.IO;
using System.Web.Mvc;

namespace Library_benchmark.Controllers
{
    public class PDFController : Controller
    {
        // GET: PDF
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ITextSharp()
        {
            return PartialView();
        }

        [HttpPost]
        public ActionResult ITextSharp(Parametros parametros)
        {
            var informacion = new Consultas().GetPdfInformacion();
            if (informacion == null) return null;

            var pdf = new Document();
            
            for (var i = 0; i < parametros.Iteraciones;)
            {
                
                var workStream = new MemoryStream();
                PdfWriter.GetInstance(pdf, workStream).CloseStream = false;
                var strFilePath = Server.MapPath("~/PdfUploads/");

                var fileName = "Pdf_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".pdf";
                var doc = new Document(PageSize.A4, 5f, 5f, 73.5f, 70f);

                //var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(strFilePath + fileName, FileMode.Create));
                //pdfWriter.PageEvent = new ITextEvents();

                new TextSharpServicio(informacion, parametros.Template, doc,strFilePath, fileName);

                
                var contents = System.IO.File.ReadAllBytes(strFilePath + fileName);
                return File(contents, "application/pdf", fileName);


            }



            return null;
        }
        

    }
}