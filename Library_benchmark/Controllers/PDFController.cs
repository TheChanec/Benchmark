
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

        [HttpPost]
        public ActionResult GeneratePdf()
        {
            //var userTickets = JsonConvert.DeserializeObject<List<TicketsByNameDetails>>(jsonString);

            var doc = new Document(PageSize.A4, 10f, 10f, 120f, 100f);
            var strFilePath = Server.MapPath("~/PdfUploads/");

            var fileName = "Pdf_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".pdf";

            var titleFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.UNDERLINE, BaseColor.BLACK);
            var h1Font = FontFactory.GetFont(FontFactory.HELVETICA, 11, Font.NORMAL);
            var bodyFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.DARK_GRAY);


            var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(strFilePath + fileName, FileMode.Create));
            pdfWriter.PageEvent = new ITextEvents();
            doc.Open();

            var tblContainer = new PdfPTable(5) { TotalWidth = 520f, LockedWidth = true };
            float[] widths = { 90f, 150f, 120f, 95f, 65f };
            tblContainer.SetWidths(widths);
            var heading = new Phrase("LearnShareCorner demo code to Genereate Pdf using ITextSharp.", h1Font);
            var titleEmployee = new Phrase("Employee", titleFont);
            var titleName = new Phrase("Name", titleFont);
            var titleOccupation = new Phrase("Occupation", titleFont);
            var titleLa = new Phrase("Lapse Action", titleFont);
            var titleExpiryDate = new Phrase("Expiry Date", titleFont);
            var cellTicketName = new PdfPCell(heading) { Colspan = 5, Border = 0 };
            var cellTitleEmployee = new PdfPCell(titleEmployee);
            var cellTitleName = new PdfPCell(titleName);
            var cellTitleOccupation = new PdfPCell(titleOccupation);
            var cellTitleLa = new PdfPCell(titleLa);
            var cellTitleExpiryDate = new PdfPCell(titleExpiryDate);

            cellTitleEmployee.Border = 0;
            cellTitleName.Border = 0;
            cellTitleOccupation.Border = 0;
            cellTitleLa.Border = 0;
            cellTitleExpiryDate.Border = 0;

            tblContainer.AddCell(cellTicketName);

            tblContainer.AddCell(cellTitleEmployee);
            tblContainer.AddCell(cellTitleName);
            tblContainer.AddCell(cellTitleOccupation);
            tblContainer.AddCell(cellTitleLa);
            tblContainer.AddCell(cellTitleExpiryDate);

            doc.Add(tblContainer);

            var tblResult = new PdfPTable(5) { TotalWidth = 520f, LockedWidth = true };
            tblResult.SetWidths(widths);
            var employee = new Phrase("WebTechSys.in", bodyFont);
            var name = new Phrase("Mukesh Salaria", bodyFont);
            var occupation = new Phrase("Software Engineer", bodyFont);
            var la = new Phrase("None", bodyFont);

            var expiryDate = new Phrase("N/A", bodyFont);
            var cellEmployee = new PdfPCell(employee);
            var cellName = new PdfPCell(name);
            var cellOccupation = new PdfPCell(occupation);
            var cellLa = new PdfPCell(la);
            var cellExpiryDate = new PdfPCell(expiryDate);

            cellEmployee.Border = 0;
            cellName.Border = 0;
            cellOccupation.Border = 0;
            cellLa.Border = 0;
            cellExpiryDate.Border = 0;
            tblResult.AddCell(cellEmployee);
            tblResult.AddCell(cellName);
            tblResult.AddCell(cellOccupation);
            tblResult.AddCell(cellLa);
            tblResult.AddCell(cellExpiryDate);

            doc.Add(tblResult);

            doc.Close();
            //return Json(new { success = "true", link = strFilePath + fileName });
            byte[] contents = System.IO.File.ReadAllBytes(strFilePath + fileName);
            return File(contents, "application/pdf", fileName);


        }


    }
}