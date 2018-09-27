using System;
using System.Globalization;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Library_benchmark.Helpers.ITextSharp;
using Library_benchmark.Models;
using NPOI.HSSF.Record;
using Font = iTextSharp.text.Font;
using Image = iTextSharp.text.Image;

namespace Library_benchmark.Helpers.IText
{
    /// <summary>
    /// Servicio Encargado de la Generacion de PDF's
    /// </summary>
    public class TextSharpServicio
    {
        private readonly Document _doc;
        private readonly MemoryStream _workStream;
        private readonly PdfWriter _pdfWriter;
        private readonly int _hojas;
        private readonly PdfDummy _informacion;
        private readonly string _filePath;
        private readonly string _fileName;
        private static Font _fontDoceBold;
        private static Font _fontDiesiocho;
        private static Font _fontOnceTres;
        private static Font _fontDoce;
        private static float _fixedHeight;
        private static float _fixedHeightSeparacion;
        private static float _borderWidth;
        private static BaseColor _baseColorLines;
        private static BaseColor _baseColorBackground;
        private static BaseColor _baseColorSeparacion;

        public TextSharpServicio(PdfDummy informacion, bool template, int hojas)
        {
            _doc = new Document(PageSize.A4, 5F, 5F, 73.5F, 70f);

            _workStream = new MemoryStream();
            _pdfWriter = PdfWriter.GetInstance(_doc, _workStream);
            _pdfWriter.CloseStream = false;
            _hojas = hojas;
           



            _informacion = informacion;


            if (!template)
                GenerarPdf();
            else
                UtilizarTemplatePdf();
        }

        private void GenerarPdf()
        {

            _pdfWriter.PageEvent = new TextEvents();
            _doc.Open();

            InicializarFonts();
            InicializarEstilos();

            for (int i = 0; i < _hojas; i++)
            {
                if (i != 0)
                {
                    _doc.NewPage();
                }
                
                CrearHeader();
                CrearInformacionGeneral();
                CrearCritical();
                CrearMechanical();
                CrearSupervisor();
            }

            

            _doc.Close();

        }

        private void UtilizarTemplatePdf()
        {

            const string fileNameExisting = @"C:\Users\mario.chan\Documents\GitHub\Benchmark\Library_benchmark\Content\Resource\DummyTemplate.pdf";
            var fileNameNew = @"C:\Users\mario.chan\Documents\GitHub\Benchmark\Library_benchmark\PdfUploads\" + "prueba.pdf";

            using (var existingFileStream = new FileStream(fileNameExisting, FileMode.Open))
            using (var newFileStream = new FileStream(fileNameNew, FileMode.Create))
            {

                var pdfReader = new PdfReader(existingFileStream);
                var stamper = new PdfStamper(pdfReader, newFileStream);

                var form = stamper.AcroFields;
                var fieldKeys = form.Fields.Keys;

                foreach (var fieldKey in fieldKeys)
                {
                    switch (fieldKey)
                    {
                        case "Fields.truck.FieldValue":
                            form.SetField(fieldKey, _informacion.Truck);
                            break;
                        case "body":

                            break;
                        case "Labels.pageNumberAndCount":
                            form.SetField(fieldKey, "2");
                            break;
                    }
                }




                stamper.FormFlattening = true;

                // You can also specify fields to be flattened, which
                // leaves the rest of the form still be editable/usable
                stamper.PartialFormFlattening("field1");

                stamper.Close();
                pdfReader.Close();
            }
        }



        #region Creacion del pdf

        /// <summary>
        /// Funcion que Inicializa valores como el background y los espacion al generar el pdf
        /// </summary>
        private static void InicializarEstilos()
        {
            _fixedHeight = 30f;
            _fixedHeightSeparacion = 10f;
            _borderWidth = 2f;
            _baseColorLines = new BaseColor(211, 211, 211);
            _baseColorBackground = new BaseColor(238, 239, 239);
            _baseColorSeparacion = new BaseColor(223, 223, 223);
        }
        /// <summary>
        /// Inicializa los tipos de letra que se van a utilizar al generar el pdf
        /// </summary>
        private static void InicializarFonts()
        {
            _fontDoceBold = FontFactory.GetFont(FontFactory.HELVETICA, 12f, Font.BOLD, BaseColor.BLACK);
            _fontDiesiocho = FontFactory.GetFont(FontFactory.HELVETICA, 18f, Font.NORMAL, BaseColor.WHITE);
            _fontOnceTres = FontFactory.GetFont(FontFactory.HELVETICA, 11.3f, Font.NORMAL, BaseColor.WHITE);
            _fontDoce = FontFactory.GetFont(FontFactory.HELVETICA, 12f, Font.NORMAL, BaseColor.BLACK);
        }
        /// <summary>
        /// Crea la Table para poder llenar Titulo y la primera seccion de el PDF
        /// </summary>
        private void CrearHeader()
        {
            

            var tblContainer = new PdfPTable(4) { TotalWidth = 558f, LockedWidth = true };

            var title = new Phrase(_informacion.Title, _fontDiesiocho);
            var subtitle = new Phrase(_informacion.SubTitle, _fontOnceTres);

            var titleDate = new Phrase("DATE", _fontDoceBold);
            var titleDriver = new Phrase("DRIVER", _fontDoceBold);
            var titleTruck = new Phrase("TRUCK", _fontDoceBold);
            var titleHour = new Phrase("HOUR", _fontDoceBold);

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



            cellDate.Border = 0;
            cellDate.FixedHeight = _fixedHeight;
            cellDate.BackgroundColor = _baseColorBackground;
            cellDate.BorderWidthRight = _borderWidth;
            cellDate.BorderColor = _baseColorLines;

            cellDrive.Border = 0;
            cellDrive.FixedHeight = _fixedHeight;
            cellDrive.BackgroundColor = _baseColorBackground;
            cellDrive.BorderWidthRight = _borderWidth;
            cellDrive.BorderColor = _baseColorLines;

            cellTruck.Border = 0;
            cellTruck.FixedHeight = _fixedHeight;
            cellTruck.BackgroundColor = _baseColorBackground;
            cellTruck.BorderWidthRight = _borderWidth;
            cellTruck.BorderColor = _baseColorLines;

            cellHour.Border = 0;
            cellHour.BackgroundColor = _baseColorBackground;
            cellHour.BorderWidthRight = 0f;

            tblContainer.AddCell(cellTitle);
            tblContainer.AddCell(cellSubTitle);

            tblContainer.AddCell(cellDate);
            tblContainer.AddCell(cellDrive);
            tblContainer.AddCell(cellTruck);
            tblContainer.AddCell(cellHour);





            _doc.Add(tblContainer);

            var tblResult = new PdfPTable(4) { TotalWidth = 558f, LockedWidth = true };

            var date = new Phrase(_informacion.Date.ToString(CultureInfo.InvariantCulture)/*"21 Apr, 2017"*/, _fontDoce);
            var driver = new Phrase(_informacion.Driver, _fontDoce);
            var occupation = new Phrase(_informacion.Truck, _fontDoce);
            var hour = new Phrase(_informacion.Date.ToString(CultureInfo.InvariantCulture)/*"08:54 AM"*/, _fontDoce);

            var cellEmployee = new PdfPCell(date);
            var cellName = new PdfPCell(driver);
            var cellOccupation = new PdfPCell(occupation);
            var cellExpiryDate = new PdfPCell(hour);

            cellEmployee.Border = 0;
            cellEmployee.FixedHeight = _fixedHeight;
            cellEmployee.BackgroundColor = _baseColorBackground;
            cellEmployee.BorderWidthRight = _borderWidth;
            cellEmployee.BorderWidthBottom = _borderWidth;
            cellEmployee.BorderColor = _baseColorLines;

            cellName.Border = 0;
            cellName.FixedHeight = _fixedHeight;
            cellName.BackgroundColor = _baseColorBackground;
            cellName.BorderWidthRight = _borderWidth;
            cellName.BorderWidthBottom = _borderWidth;
            cellName.BorderColor = _baseColorLines;

            cellOccupation.Border = 0;
            cellOccupation.FixedHeight = _fixedHeight;
            cellOccupation.BackgroundColor = _baseColorBackground;
            cellOccupation.BorderWidthRight = _borderWidth;
            cellOccupation.BorderWidthBottom = _borderWidth;
            cellOccupation.BorderColor = _baseColorLines;

            cellExpiryDate.Border = 0;
            cellExpiryDate.BackgroundColor = _baseColorBackground;
            cellExpiryDate.BorderWidthBottom = _borderWidth;
            cellExpiryDate.BorderColor = _baseColorLines;

            var cellfinal = new PdfPCell
            {
                Border = 0,
                Colspan = 4,
                BackgroundColor = _baseColorSeparacion,
                FixedHeight = _fixedHeightSeparacion
            };


            tblResult.AddCell(cellEmployee);
            tblResult.AddCell(cellName);
            tblResult.AddCell(cellOccupation);
            tblResult.AddCell(cellExpiryDate);

            tblResult.AddCell(cellfinal);

            _doc.Add(tblResult);
        }
        /// <summary>
        /// Crea la seccion donde vine la informacion mas relevante del PDF
        /// </summary>
        private void CrearInformacionGeneral()
        {

            var tblGeneralInformation = new PdfPTable(3) { TotalWidth = 558f, LockedWidth = true };
            var informacionGeneral = new Phrase("GENERAL INFORMATION", _fontDoceBold);
            var odometerStar = new Phrase("Odometer Start", _fontDoceBold);
            var maxAirPressure = new Phrase("Max Air Pressure PSI", _fontDoceBold);
            var lowAirWarningDevice = new Phrase("Low Air Warning Device PSI", _fontDoceBold);

            var cellInformacionGeneral = new PdfPCell(informacionGeneral);
            var cellOdometerStar = new PdfPCell(odometerStar);
            var cellMaxAirPressure = new PdfPCell(maxAirPressure);
            var cellLowAirWarningDevice = new PdfPCell(lowAirWarningDevice);


            var valueOdometerStar = new Phrase(_informacion.OdometerStatr.ToString(), _fontDoceBold);
            var valuemaxAirPressure = new Phrase(_informacion.PressurePsi.ToString(), _fontDoceBold);
            var valueLowAirWarningDevice = new Phrase(_informacion.DevicePsi.ToString(), _fontDoceBold);


            var cellValueOdometerStar = new PdfPCell(valueOdometerStar);
            var cellValueMaxAirPressure = new PdfPCell(valuemaxAirPressure);
            var cellValueLowAirWarningDevice = new PdfPCell(valueLowAirWarningDevice);

            var valueSecoundaryOdometerStar = new Phrase("--", _fontDoceBold);
            var valueSecoundarymaxAirPressure = new Phrase("--", _fontDoceBold);
            var valueSecoundaryLowAirWarningDevice = new Phrase("--", _fontDoceBold);


            var cellSecoundaryValueOdometerStar = new PdfPCell(valueSecoundaryOdometerStar);
            var cellSecoundaryValueMaxAirPressure = new PdfPCell(valueSecoundarymaxAirPressure);
            var cellSecoundaryValueLowAirWarningDevice = new PdfPCell(valueSecoundaryLowAirWarningDevice);


            cellInformacionGeneral.Colspan = 3;
            cellInformacionGeneral.Border = 0;
            cellInformacionGeneral.BackgroundColor = _baseColorBackground;
            cellInformacionGeneral.BorderWidthTop = _borderWidth;
            cellInformacionGeneral.BorderWidthBottom = _borderWidth;
            cellInformacionGeneral.BorderColor = _baseColorLines;
            cellInformacionGeneral.FixedHeight = _fixedHeight;



            cellOdometerStar.Border = 0;
            cellOdometerStar.FixedHeight = _fixedHeight;

            cellMaxAirPressure.Border = 0;
            cellMaxAirPressure.FixedHeight = _fixedHeight;

            cellLowAirWarningDevice.Border = 0;
            cellLowAirWarningDevice.FixedHeight = _fixedHeight;

            cellValueOdometerStar.Border = 0;

            cellValueMaxAirPressure.Border = 0;

            cellValueLowAirWarningDevice.Border = 0;

            cellSecoundaryValueOdometerStar.Border = 0;
            cellSecoundaryValueOdometerStar.BorderWidthBottom = _borderWidth;
            cellSecoundaryValueOdometerStar.BorderColor = _baseColorLines;

            cellSecoundaryValueMaxAirPressure.Border = 0;
            cellSecoundaryValueMaxAirPressure.BorderWidthBottom = _borderWidth;
            cellSecoundaryValueMaxAirPressure.BorderColor = _baseColorLines;

            cellSecoundaryValueLowAirWarningDevice.Border = 0;
            cellSecoundaryValueLowAirWarningDevice.BorderWidthBottom = _borderWidth;
            cellSecoundaryValueLowAirWarningDevice.BorderColor = _baseColorLines;

            //cellfinal = new PdfPCell
            //{
            //    Border = 0,
            //    Colspan = 3,
            //    BackgroundColor = _baseColorSeparacion,
            //    FixedHeight = _fixedHeightSeparacion
            //};


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
            //tblGeneralInformation.AddCell(cellfinal);

            _doc.Add(tblGeneralInformation);
        }
        /// <summary>
        /// CRea la seccion Citical
        /// </summary>
        private void CrearCritical()
        {

            var tblCritical = new PdfPTable(2) { TotalWidth = 558f, LockedWidth = true };
            var critical = new Phrase("Critical", _fontDoceBold);
            var water = new Phrase("Water", _fontDoceBold);
            var valueWater = new Phrase(_informacion.Water, _fontDoceBold);

            var cellCritical = new PdfPCell(critical);
            var cellWater = new PdfPCell(water);
            var cellValueWater = new PdfPCell(valueWater);


            cellCritical.Colspan = 2;
            cellCritical.Border = 0;
            cellCritical.BackgroundColor = _baseColorBackground;
            cellCritical.BorderWidthTop = _borderWidth;
            cellCritical.BorderWidthBottom = _borderWidth;
            cellCritical.BorderColor = _baseColorLines;
            cellCritical.FixedHeight = _fixedHeight;

            cellWater.Border = 0;
            cellWater.Colspan = 2;
            cellWater.FixedHeight = _fixedHeight;



            cellValueWater.Border = 0;
            cellValueWater.Colspan = 2;
            cellValueWater.BorderWidthBottom = _borderWidth;
            cellValueWater.BorderColor = _baseColorLines;

            //cellfinal = new PdfPCell
            //{
            //    Border = 0,
            //    Colspan = 3,
            //    BackgroundColor = _baseColorSeparacion,
            //    FixedHeight = _fixedHeightSeparacion
            //};

            tblCritical.AddCell(cellCritical);
            tblCritical.AddCell(cellWater);
            tblCritical.AddCell(cellValueWater);
            //tblCritical.AddCell(cellfinal);

            _doc.Add(tblCritical);
        }
        /// <summary>
        /// Crae la seccion donde vienen los comentarios del Mecanico
        /// </summary>
        private void CrearMechanical()
        {

            var tblMechanical = new PdfPTable(2) { TotalWidth = 558f, LockedWidth = true };
            var mechanical = new Phrase("Mechanical", _fontDoceBold);
            var leaks = new Phrase("Leaks: Water, Oil, Fuel, Grease", _fontDoceBold);
            var valueLeaks = new Phrase(_informacion.MechanicalComments, _fontDoceBold);

            var cellMechanical = new PdfPCell(mechanical);
            var cellLeaks = new PdfPCell(leaks);
            var cellValueLeaks = new PdfPCell(valueLeaks);


            cellMechanical.Colspan = 2;
            cellMechanical.Border = 0;
            cellMechanical.BackgroundColor = _baseColorBackground;
            cellMechanical.BorderWidthTop = _borderWidth;
            cellMechanical.BorderWidthBottom = _borderWidth;
            cellMechanical.BorderColor = _baseColorLines;
            cellMechanical.FixedHeight = _fixedHeight;

            cellLeaks.Border = 0;
            cellLeaks.Colspan = 2;
            cellLeaks.FixedHeight = _fixedHeight;



            cellValueLeaks.Border = 0;
            cellValueLeaks.Colspan = 2;
            cellValueLeaks.BorderWidthBottom = _borderWidth;
            cellValueLeaks.BorderColor = _baseColorLines;


            tblMechanical.AddCell(cellMechanical);
            tblMechanical.AddCell(cellLeaks);
            tblMechanical.AddCell(cellValueLeaks);

            _doc.Add(tblMechanical);
        }
        /// <summary>
        /// Crea la seccion donde vienen los comentarios del Supervisor
        /// </summary>
        private void CrearSupervisor()
        {

            var tblSupervisor = new PdfPTable(2) { TotalWidth = 558f, LockedWidth = true };
            var supervisor = new Phrase("Supervisor", _fontDoceBold);
            var valueDateSupervisor = new Phrase("21 Apr, 2017", _fontDoceBold);
            var date = new Phrase(_informacion.Date.ToString(CultureInfo.CurrentCulture)/*"21 Apr, 2017"*/, _fontDoce);

            var cellSupervisor = new PdfPCell(supervisor);
            var cellDateSupervisor = new PdfPCell(date);
            var cellValueDateSupervisor = new PdfPCell(valueDateSupervisor);
            var firma = Image.GetInstance("C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/firma.PNG");
            firma.ScalePercent(2);
            var cellFirma = new PdfPCell(firma, true);


            cellSupervisor.Rowspan = 2;
            cellSupervisor.Border = 0;
            cellSupervisor.BackgroundColor = _baseColorBackground;
            cellSupervisor.BorderWidthTop = _borderWidth;
            cellSupervisor.BorderColor = _baseColorLines;
            cellSupervisor.FixedHeight = _fixedHeight;

            cellDateSupervisor.Border = 0;
            cellDateSupervisor.BackgroundColor = _baseColorBackground;
            cellDateSupervisor.BorderWidthTop = _borderWidth;
            cellDateSupervisor.BorderColor = _baseColorLines;
            cellDateSupervisor.FixedHeight = _fixedHeight;

            cellValueDateSupervisor.Border = 0;
            cellValueDateSupervisor.BackgroundColor = _baseColorBackground;
            cellValueDateSupervisor.BorderColor = _baseColorLines;
            cellValueDateSupervisor.FixedHeight = _fixedHeight;

            cellFirma.Colspan = 2;
            cellFirma.Border = 0;
            cellFirma.BackgroundColor = _baseColorBackground;
            cellFirma.BorderWidthBottom = _borderWidth;
            cellFirma.BorderColor = _baseColorLines;


            tblSupervisor.AddCell(cellSupervisor);
            tblSupervisor.AddCell(cellDateSupervisor);
            tblSupervisor.AddCell(cellValueDateSupervisor);
            tblSupervisor.AddCell(cellFirma);

            _doc.Add(tblSupervisor);
        }

        #endregion


        internal byte[] GetFile()
        {
            return this._workStream.ToArray();
        }
    }
}