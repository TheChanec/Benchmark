
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

            var doc = new Document(PageSize.A4, 5f, 5f, 73.5f, 70f);
            var strFilePath = Server.MapPath("~/PdfUploads/");

            var fileName = "Pdf_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".pdf";

            var fontDoc = FontFactory.GetFont(FontFactory.HELVETICA, 12f, Font.BOLD, BaseColor.BLACK);
            var fontTitle = FontFactory.GetFont(FontFactory.HELVETICA, 18, Font.NORMAL, BaseColor.WHITE);
            var fontSubTitle = FontFactory.GetFont(FontFactory.HELVETICA, 11.3f, Font.NORMAL, BaseColor.WHITE);

            var bodyFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.DARK_GRAY);


            var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(strFilePath + fileName, FileMode.Create));
            pdfWriter.PageEvent = new ITextEvents();
            doc.Open();
            
            var tblContainer = new PdfPTable(4) { TotalWidth = 558f, LockedWidth = true };
            
            //float[] widths = { 90f, 150f, 120f, 95f, 65f };
            //tblContainer.SetWidths(widths);
            var title = new Phrase("DRIVER INSPECTION REPORT", fontTitle);
            var subtitle = new Phrase("AS REQUIRED BY DOT FEDERAL MOTOR CARRIER SAFETY REGULATIONS", fontSubTitle);

            var titleDate = new Phrase("DATE", fontDoc);
            var titleDriver = new Phrase("DRIVER", fontDoc);
            var titleTruck = new Phrase("TRUCK", fontDoc);
            var titleHour = new Phrase("HOUR", fontDoc);

            var cellTitle = new PdfPCell(title)
            {
                Colspan = 5,
                Border = 0,
                BackgroundColor = BaseColor.BLACK,
                HorizontalAlignment = PdfPCell.ALIGN_CENTER,
                VerticalAlignment = PdfPCell.ALIGN_BOTTOM,
                FixedHeight = 29f
            };
            var cellSubTitle = new PdfPCell(subtitle)
            {
                Colspan = 5,
                Border = 0,
                BackgroundColor = BaseColor.BLACK,
                HorizontalAlignment = PdfPCell.ALIGN_CENTER,
                VerticalAlignment = PdfPCell.ALIGN_TOP,
                FixedHeight = 29f
            };
            var cellDate = new PdfPCell(titleDate);
            var cellDrive = new PdfPCell(titleDriver);
            var cellTruck = new PdfPCell(titleTruck);
            var cellHour = new PdfPCell(titleHour);

            var fixedHeight = 30f;
            var borderWidth = 2f;
            var baseColorLines = new BaseColor(211, 211, 211);
            var baseColorBackground = new BaseColor(238, 239, 239);

            cellDate.Border = 0;
            cellDate.FixedHeight = fixedHeight;
            cellDate.BackgroundColor = baseColorBackground;
            cellDate.BorderWidthRight= borderWidth;
            cellDate.BorderColor = baseColorLines;

            cellDrive.Border = 0;
            cellDrive.FixedHeight = fixedHeight;
            cellDrive.BackgroundColor = baseColorBackground;
            cellDrive.BorderWidthRight = borderWidth;
            cellDrive.BorderColor = baseColorLines;

            cellTruck.Border = 0;
            cellTruck.FixedHeight = fixedHeight;
            cellTruck.BackgroundColor = baseColorBackground;
            cellTruck.BorderWidthRight = borderWidth;
            cellTruck.BorderColor = baseColorLines;

            cellHour.Border = 0;
            cellHour.BackgroundColor = baseColorBackground;
            cellHour.BorderWidthRight = 0f;

            tblContainer.AddCell(cellTitle);
            tblContainer.AddCell(cellSubTitle);

            tblContainer.AddCell(cellDate);
            tblContainer.AddCell(cellDrive);
            tblContainer.AddCell(cellTruck);
            tblContainer.AddCell(cellHour);

            doc.Add(tblContainer);

            var tblResult = new PdfPTable(4) { TotalWidth = 558f, LockedWidth = true };
            //tblResult.SetWidths(widths);
            var date = new Phrase("21 Apr, 2017", bodyFont);
            var driver = new Phrase("HUGO ISAAC RODRIGUEZ", bodyFont);
            var occupation = new Phrase("CR4150", bodyFont);
            var hour = new Phrase("08:54 AM", bodyFont);

            var cellEmployee = new PdfPCell(date);
            var cellName = new PdfPCell(driver);
            var cellOccupation = new PdfPCell(occupation);
            var cellExpiryDate = new PdfPCell(hour);

            cellEmployee.Border = 0;
            cellEmployee.FixedHeight = fixedHeight;
            cellEmployee.BackgroundColor = baseColorBackground;
            cellEmployee.BorderWidthRight = borderWidth;
            cellEmployee.BorderWidthBottom = borderWidth;
            cellEmployee.BorderColor = baseColorLines;

            cellName.Border = 0;
            cellName.FixedHeight = fixedHeight;
            cellName.BackgroundColor = baseColorBackground;
            cellName.BorderWidthRight = borderWidth;
            cellName.BorderWidthBottom = borderWidth;
            cellName.BorderColor = baseColorLines;

            cellOccupation.Border = 0;
            cellOccupation.FixedHeight = fixedHeight;
            cellOccupation.BackgroundColor = baseColorBackground;
            cellOccupation.BorderWidthRight = borderWidth;
            cellOccupation.BorderWidthBottom = borderWidth;
            cellOccupation.BorderColor = baseColorLines;

            cellExpiryDate.Border = 0;
            cellExpiryDate.BackgroundColor = baseColorBackground;
            cellExpiryDate.BorderWidthBottom = borderWidth;
            cellExpiryDate.BorderColor = baseColorLines;


            tblResult.AddCell(cellEmployee);
            tblResult.AddCell(cellName);
            tblResult.AddCell(cellOccupation);
            tblResult.AddCell(cellExpiryDate);

            doc.Add(tblResult);

            var tblGeneralInformation = new PdfPTable(3) { TotalWidth = 558f, LockedWidth = true };
            var informacionGeneral = new Phrase("GENERAL INFORMATION", fontDoc);
            var odometerStar = new Phrase("Odometer Start", fontDoc);
            var maxAirPressure = new Phrase("Max Air Pressure PSI", fontDoc);
            var lowAirWarningDevice = new Phrase("Low Air Warning Device PSI", fontDoc);

            var cellInformacionGeneral = new PdfPCell(informacionGeneral);
            var cellOdometerStar = new PdfPCell(odometerStar);
            var cellMaxAirPressure = new PdfPCell(maxAirPressure);
            var cellLowAirWarningDevice = new PdfPCell(lowAirWarningDevice);
            

            var valueOdometerStar = new Phrase("5", fontDoc);
            var valuemaxAirPressure = new Phrase("8", fontDoc);
            var valueLowAirWarningDevice = new Phrase("4", fontDoc);
            

            var cellValueOdometerStar = new PdfPCell(valueOdometerStar);
            var cellValueMaxAirPressure = new PdfPCell(valuemaxAirPressure);
            var cellValueLowAirWarningDevice = new PdfPCell(valueLowAirWarningDevice);

            var valueSecoundaryOdometerStar = new Phrase("--", fontDoc);
            var valueSecoundarymaxAirPressure = new Phrase("--", fontDoc);
            var valueSecoundaryLowAirWarningDevice = new Phrase("--", fontDoc);


            var cellSecoundaryValueOdometerStar = new PdfPCell(valueSecoundaryOdometerStar);
            var cellSecoundaryValueMaxAirPressure = new PdfPCell(valueSecoundarymaxAirPressure);
            var cellSecoundaryValueLowAirWarningDevice = new PdfPCell(valueSecoundaryLowAirWarningDevice);

            cellInformacionGeneral.Colspan = 3;
            cellInformacionGeneral.Border = 0;
            cellInformacionGeneral.BackgroundColor = baseColorBackground;
            cellInformacionGeneral.BorderWidthTop = borderWidth;
            cellInformacionGeneral.BorderWidthBottom = borderWidth;
            cellInformacionGeneral.BorderColor = baseColorLines;
            cellInformacionGeneral.FixedHeight = fixedHeight;

            
                
            cellOdometerStar.Border = 0;
            cellOdometerStar.FixedHeight = fixedHeight;

            cellMaxAirPressure.Border = 0;
            cellMaxAirPressure.FixedHeight = fixedHeight;

            cellLowAirWarningDevice.Border = 0;
            cellLowAirWarningDevice.FixedHeight = fixedHeight;

            cellValueOdometerStar.Border = 0;

            cellValueMaxAirPressure.Border = 0;

            cellValueLowAirWarningDevice.Border = 0;

            cellSecoundaryValueOdometerStar.Border = 0;
            cellSecoundaryValueOdometerStar.BorderWidthBottom = borderWidth;
            cellSecoundaryValueOdometerStar.BorderColor = baseColorLines;

            cellSecoundaryValueMaxAirPressure.Border = 0;
            cellSecoundaryValueMaxAirPressure.BorderWidthBottom = borderWidth;
            cellSecoundaryValueMaxAirPressure.BorderColor = baseColorLines;

            cellSecoundaryValueLowAirWarningDevice.Border = 0;
            cellSecoundaryValueLowAirWarningDevice.BorderWidthBottom = borderWidth;
            cellSecoundaryValueLowAirWarningDevice.BorderColor = baseColorLines;


            tblGeneralInformation.AddCell(cellInformacionGeneral);
            tblGeneralInformation.AddCell(cellOdometerStar);
            tblGeneralInformation.AddCell(cellMaxAirPressure);
            tblGeneralInformation.AddCell(cellLowAirWarningDevice);
            tblGeneralInformation.AddCell(cellValueOdometerStar);
            tblGeneralInformation.AddCell(cellValueMaxAirPressure);
            tblGeneralInformation.AddCell(cellValueLowAirWarningDevice);
            tblGeneralInformation.AddCell(cellSecoundaryValueOdometerStar);
            tblGeneralInformation.AddCell(cellSecoundaryValueMaxAirPressure);
            tblGeneralInformation.AddCell(cellSecoundaryValueLowAirWarningDevice);

            doc.Add(tblGeneralInformation);

            var tblCritical = new PdfPTable(2) { TotalWidth = 558f, LockedWidth = true };
            var critical = new Phrase("Critical", fontDoc);
            var water = new Phrase("Water", fontDoc);
            var valueWater = new Phrase("hsksnsk", fontDoc);

            var cellCritical = new PdfPCell(critical);
            var cellWater = new PdfPCell(water);
            var cellValueWater = new PdfPCell(valueWater);

            
            cellCritical.Colspan = 2;
            cellCritical.Border = 0;
            cellCritical.BackgroundColor = baseColorBackground;
            cellCritical.BorderWidthTop = borderWidth;
            cellCritical.BorderWidthBottom = borderWidth;
            cellCritical.BorderColor = baseColorLines;
            cellCritical.FixedHeight = fixedHeight;

            cellWater.Border = 0;
            cellCritical.Colspan = 2;
            cellWater.FixedHeight = fixedHeight;
            


            cellValueWater.Border = 0;
            cellCritical.Colspan = 2;
            cellValueWater.BorderWidthBottom = borderWidth;
            cellValueWater.BorderColor = baseColorLines;
            

            tblCritical.AddCell(cellCritical);
            tblCritical.AddCell(cellWater);
            tblCritical.AddCell(cellValueWater);

            doc.Add(tblCritical);

            doc.Close();
            //return Json(new { success = "true", link = strFilePath + fileName });
            byte[] contents = System.IO.File.ReadAllBytes(strFilePath + fileName);
            return File(contents, "application/pdf", fileName);


        }


    }
}