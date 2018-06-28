
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

                var pdf = new Document();

                if (informacion == null) continue;

                var watchCreation = Stopwatch.StartNew();

                MemoryStream workStream = new MemoryStream();
                PdfWriter.GetInstance(pdf, workStream).CloseStream = false;
                pdf = new ITextSharpServicio().GetPDFExample();

                byte[] byteInfo = workStream.ToArray();
                workStream.Write(byteInfo, 0, byteInfo.Length);
                workStream.Position = 0;
                Response.Buffer = true;
                Response.AddHeader("Content-Disposition", "attachment; filename= " + Server.HtmlEncode("abc.pdf"));
                Response.ContentType = "APPLICATION/pdf";
                Response.BinaryWrite(byteInfo);
            }


        }

        [HttpPost]
        public ActionResult GeneratePdf()
        {
            var doc = new Document(PageSize.A4, 5f, 5f, 73.5f, 70f);
            var strFilePath = Server.MapPath("~/PdfUploads/");

            var fileName = "Pdf_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".pdf";

            var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(strFilePath + fileName, FileMode.Create));
            pdfWriter.PageEvent = new ITextEvents();
            doc.Open();

            #region Fonts
            var fontDoceBold = FontFactory.GetFont(FontFactory.HELVETICA, 12f, Font.BOLD, BaseColor.BLACK);
            var fontDiesiocho = FontFactory.GetFont(FontFactory.HELVETICA, 18f, Font.NORMAL, BaseColor.WHITE);
            var fontOnceTres = FontFactory.GetFont(FontFactory.HELVETICA, 11.3f, Font.NORMAL, BaseColor.WHITE);
            var fontDoce = FontFactory.GetFont(FontFactory.HELVETICA, 12f, Font.NORMAL, BaseColor.BLACK);
            #endregion
            

            var tblContainer = new PdfPTable(4) { TotalWidth = 558f, LockedWidth = true };

            var title = new Phrase("DRIVER INSPECTION REPORT", fontDiesiocho);
            var subtitle = new Phrase("AS REQUIRED BY DOT FEDERAL MOTOR CARRIER SAFETY REGULATIONS", fontOnceTres);

            var titleDate = new Phrase("DATE", fontDoceBold);
            var titleDriver = new Phrase("DRIVER", fontDoceBold);
            var titleTruck = new Phrase("TRUCK", fontDoceBold);
            var titleHour = new Phrase("HOUR", fontDoceBold);

            var cellTitle = new PdfPCell(title)
            {
                Colspan = 5,
                Border = 0,
                BackgroundColor = BaseColor.BLACK,
                HorizontalAlignment = Element.ALIGN_CENTER,
                VerticalAlignment = Element.ALIGN_BOTTOM,
                FixedHeight = 29f
            };
            var cellSubTitle = new PdfPCell(subtitle)
            {
                Colspan = 5,
                Border = 0,
                BackgroundColor = BaseColor.BLACK,
                HorizontalAlignment = Element.ALIGN_CENTER,
                VerticalAlignment = Element.ALIGN_TOP,
                FixedHeight = 29f
            };
            var cellDate = new PdfPCell(titleDate);
            var cellDrive = new PdfPCell(titleDriver);
            var cellTruck = new PdfPCell(titleTruck);
            var cellHour = new PdfPCell(titleHour);

            const float fixedHeight = 30f;
            const float fixedHeightSeparacion = 10f;
            const float borderWidth = 2f;
            var baseColorLines = new BaseColor(211, 211, 211);
            var baseColorBackground = new BaseColor(238, 239, 239);
            var baseColorSeparacion = new BaseColor(223, 223, 223);

            cellDate.Border = 0;
            cellDate.FixedHeight = fixedHeight;
            cellDate.BackgroundColor = baseColorBackground;
            cellDate.BorderWidthRight = borderWidth;
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

            var date = new Phrase("21 Apr, 2017", fontDoce);
            var driver = new Phrase("HUGO ISAAC RODRIGUEZ", fontDoce);
            var occupation = new Phrase("CR4150", fontDoce);
            var hour = new Phrase("08:54 AM", fontDoce);

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

            var cellfinal = new PdfPCell();
            cellfinal.Border = 0;
            cellfinal.Colspan = 4;
            cellfinal.BackgroundColor = baseColorSeparacion;
            cellfinal.FixedHeight = fixedHeightSeparacion;


            tblResult.AddCell(cellEmployee);
            tblResult.AddCell(cellName);
            tblResult.AddCell(cellOccupation);
            tblResult.AddCell(cellExpiryDate);

            tblResult.AddCell(cellfinal);

            doc.Add(tblResult);

            var tblGeneralInformation = new PdfPTable(3) { TotalWidth = 558f, LockedWidth = true };
            var informacionGeneral = new Phrase("GENERAL INFORMATION", fontDoceBold);
            var odometerStar = new Phrase("Odometer Start", fontDoceBold);
            var maxAirPressure = new Phrase("Max Air Pressure PSI", fontDoceBold);
            var lowAirWarningDevice = new Phrase("Low Air Warning Device PSI", fontDoceBold);

            var cellInformacionGeneral = new PdfPCell(informacionGeneral);
            var cellOdometerStar = new PdfPCell(odometerStar);
            var cellMaxAirPressure = new PdfPCell(maxAirPressure);
            var cellLowAirWarningDevice = new PdfPCell(lowAirWarningDevice);


            var valueOdometerStar = new Phrase("5", fontDoceBold);
            var valuemaxAirPressure = new Phrase("8", fontDoceBold);
            var valueLowAirWarningDevice = new Phrase("4", fontDoceBold);


            var cellValueOdometerStar = new PdfPCell(valueOdometerStar);
            var cellValueMaxAirPressure = new PdfPCell(valuemaxAirPressure);
            var cellValueLowAirWarningDevice = new PdfPCell(valueLowAirWarningDevice);

            var valueSecoundaryOdometerStar = new Phrase("--", fontDoceBold);
            var valueSecoundarymaxAirPressure = new Phrase("--", fontDoceBold);
            var valueSecoundaryLowAirWarningDevice = new Phrase("--", fontDoceBold);


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

            cellfinal = new PdfPCell();
            cellfinal.Border = 0;
            cellfinal.Colspan = 3;
            cellfinal.BackgroundColor = baseColorSeparacion;
            cellfinal.FixedHeight = fixedHeightSeparacion;


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
            tblGeneralInformation.AddCell(cellfinal);

            doc.Add(tblGeneralInformation);

            var tblCritical = new PdfPTable(2) { TotalWidth = 558f, LockedWidth = true };
            var critical = new Phrase("Critical", fontDoceBold);
            var water = new Phrase("Water", fontDoceBold);
            var valueWater = new Phrase("hsksnsk", fontDoceBold);

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
            cellWater.Colspan = 2;
            cellWater.FixedHeight = fixedHeight;



            cellValueWater.Border = 0;
            cellValueWater.Colspan = 2;
            cellValueWater.BorderWidthBottom = borderWidth;
            cellValueWater.BorderColor = baseColorLines;

            cellfinal = new PdfPCell();
            cellfinal.Border = 0;
            cellfinal.Colspan = 3;
            cellfinal.BackgroundColor = baseColorSeparacion;
            cellfinal.FixedHeight = fixedHeightSeparacion;

            tblCritical.AddCell(cellCritical);
            tblCritical.AddCell(cellWater);
            tblCritical.AddCell(cellValueWater);
            tblCritical.AddCell(cellfinal);

            doc.Add(tblCritical);



            var tblMechanical = new PdfPTable(2) { TotalWidth = 558f, LockedWidth = true };
            var mechanical = new Phrase("Mechanical", fontDoceBold);
            var leaks = new Phrase("Leaks: Water, Oil, Fuel, Grease", fontDoceBold);
            var valueLeaks = new Phrase("hsksnsk", fontDoceBold);

            var cellMechanical = new PdfPCell(mechanical);
            var cellLeaks = new PdfPCell(leaks);
            var cellValueLeaks = new PdfPCell(valueLeaks);


            cellMechanical.Colspan = 2;
            cellMechanical.Border = 0;
            cellMechanical.BackgroundColor = baseColorBackground;
            cellMechanical.BorderWidthTop = borderWidth;
            cellMechanical.BorderWidthBottom = borderWidth;
            cellMechanical.BorderColor = baseColorLines;
            cellMechanical.FixedHeight = fixedHeight;

            cellLeaks.Border = 0;
            cellLeaks.Colspan = 2;
            cellLeaks.FixedHeight = fixedHeight;



            cellValueLeaks.Border = 0;
            cellValueLeaks.Colspan = 2;
            cellValueLeaks.BorderWidthBottom = borderWidth;
            cellValueLeaks.BorderColor = baseColorLines;


            tblMechanical.AddCell(cellMechanical);
            tblMechanical.AddCell(cellLeaks);
            tblMechanical.AddCell(cellValueLeaks);

            doc.Add(tblMechanical);



            var tblSupervisor = new PdfPTable(2) { TotalWidth = 558f, LockedWidth = true };
            var supervisor = new Phrase("Supervisor", fontDoceBold);
            var valueDateSupervisor = new Phrase("21 Apr, 2017", fontDoceBold);

            var cellSupervisor = new PdfPCell(supervisor);
            var cellDateSupervisor = new PdfPCell(date);
            var cellValueDateSupervisor = new PdfPCell(valueDateSupervisor);
            var firma = Image.GetInstance("C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/firma.PNG");
            firma.ScalePercent(2);
            var cellFirma = new PdfPCell(firma, true);


            cellSupervisor.Rowspan = 2;
            cellSupervisor.Border = 0;
            cellSupervisor.BackgroundColor = baseColorBackground;
            cellSupervisor.BorderWidthTop = borderWidth;
            cellSupervisor.BorderColor = baseColorLines;
            cellSupervisor.FixedHeight = fixedHeight;

            cellDateSupervisor.Border = 0;
            cellDateSupervisor.BackgroundColor = baseColorBackground;
            cellDateSupervisor.BorderWidthTop = borderWidth;
            cellDateSupervisor.BorderColor = baseColorLines;
            cellDateSupervisor.FixedHeight = fixedHeight;

            cellValueDateSupervisor.Border = 0;
            cellValueDateSupervisor.BackgroundColor = baseColorBackground;
            cellValueDateSupervisor.BorderColor = baseColorLines;
            cellValueDateSupervisor.FixedHeight = fixedHeight;

            cellFirma.Colspan = 2;
            cellFirma.Border = 0;
            cellFirma.BackgroundColor = baseColorBackground;
            cellFirma.BorderWidthBottom = borderWidth;
            cellFirma.BorderColor = baseColorLines;


            tblSupervisor.AddCell(cellSupervisor);
            tblSupervisor.AddCell(cellDateSupervisor);
            tblSupervisor.AddCell(cellValueDateSupervisor);
            tblSupervisor.AddCell(cellFirma);

            doc.Add(tblSupervisor);


            doc.Close();
            //return Json(new { success = "true", link = strFilePath + fileName });
            byte[] contents = System.IO.File.ReadAllBytes(strFilePath + fileName);
            return File(contents, "application/pdf", fileName);


        }


    }
}