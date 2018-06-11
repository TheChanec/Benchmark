
using iTextSharp.text;
using iTextSharp.text.pdf;
using Library_benchmark.Helpers;
using Library_benchmark.Helpers.ITextSharp;
using Library_benchmark.Models;
using System;
using System.Collections.Generic;
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
        public void ITextSharp(Parametros parametros)
        {
            Singleton res = Singleton.Instance;
            for (int i = 0; i < parametros.Iteraciones; i++)
            {
                Stopwatch stopWatch = Stopwatch.StartNew();
                IList<Dummy> informacion = new Consultas(parametros.Rows).GetInformacion();

                Resultado result = new Resultado();
                result.Parametro = parametros;
                result.Libreria = "ITextSharp";

                Document pdf = new Document();

                if (informacion != null)
                {
                    Stopwatch watchCreation = Stopwatch.StartNew();

                    MemoryStream workStream = new MemoryStream();
                    PdfWriter.GetInstance(pdf, workStream).CloseStream = false;
                    pdf = new ITextSharpServicio(informacion, workStream, parametros.Sheets).GetPDFExample();

                    byte[] byteInfo = workStream.ToArray();
                    workStream.Write(byteInfo, 0, byteInfo.Length);
                    workStream.Position = 0;
                    Response.Buffer = true;
                    Response.AddHeader("Content-Disposition", "attachment; filename= " + Server.HtmlEncode("abc.pdf"));
                    Response.ContentType = "APPLICATION/pdf";
                    Response.BinaryWrite(byteInfo);

                    //watchCreation.Stop();
                    //result.Tiempos.Add(new Tiempo
                    //{
                    //    Descripcion = "Creacion",
                    //    Value = watchCreation.Elapsed.ToString()
                    //});

                    //if (parametros.Design)
                    //{
                    //    Stopwatch watchDesign = Stopwatch.StartNew();
                    //    excel = new NPOIDesign(excel, parametros.Resource).GetExcelExample();

                    //    watchDesign.Stop();
                    //    result.Tiempos.Add(new Tiempo
                    //    {
                    //        Descripcion = "Diseno",
                    //        Value = watchDesign.Elapsed.ToString()
                    //    });

                    //}
                    //Stopwatch watchFiletoDonwload = Stopwatch.StartNew();
                    //FileContentResult file = NPOIdownload(excel);
                    //watchFiletoDonwload.Stop();
                    //result.Tiempos.Add(new Tiempo
                    //{
                    //    Descripcion = "File to download",
                    //    Value = watchFiletoDonwload.Elapsed.ToString()
                    //});



                    //stopWatch.Stop();
                    //result.Tiempos.Add(new Tiempo
                    //{
                    //    Descripcion = "Total",
                    //    Value = stopWatch.Elapsed.ToString()
                    //});
                    //result.Intento = i;
                    //res.Resultados.Add(result);

                    ////return file;
                    //excel = null;
                    //file = null;
                }
            }

            
        }
        
       
    }
}