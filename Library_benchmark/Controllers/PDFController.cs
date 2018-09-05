using iTextSharp.text;
using iTextSharp.text.pdf;
using Library_benchmark.Helpers;
using Library_benchmark.Models;
using System;
using System.IO;
using System.Web.Mvc;
using Library_benchmark.Helpers.IText;
using FastReport;
using Newtonsoft.Json;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using FastReport.Utils;
using System.Drawing;

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
                    //return ITextSharp(parametros);
                case 2:
                    //return FastReport(parametros);
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
            Thread thread = new Thread(new ThreadStart(CreateReport));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();


            //return null;
        }


        private void Report()
        {
            // create report instance
            Report report = new Report();
            report.Load("C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/Resource/report.frx");

            //Si el formato es JSON, entonces se deben generar clases de C#, a partir del JSON para esas clases registrarlas como Business Objects
            //Como la generación debe ser dinámica, hasta el momento he visto que se tienen que repetir el proceso en cada ejecucion del reporte
            //Falta determinar si el Assembly generado en memoria se libera correctamente una vez que se generó el reporte
            report = CreateDynamicFormObject(report);



            // register the business object
            //report.RegisterData(, "Categories");

            // design the report
            report.Show(); //report.Design();

            // free resources used by report
            //report.Dispose();
        }

        private void CreateReport()
        {
            Report report = new Report();
            ReportPage page1 = new ReportPage()
            {
                Name = "Page1"
            };
            report.Pages.Add(page1);

            // create ReportTitle band
            page1.ReportTitle = new ReportTitleBand()
            {
                Name = "ReportTitle1",
                Width = 718.2F,
                Height = 75.6F
            };
            page1.ReportTitle.Objects.Add(
                new TextObject()
                {
                    Name = "Truck",
                    Text = "que onda",
                    Left = 548.1F,
                    Top = 37.8F,
                    Width = 141.75F,
                    Height = 18.8F,
                    FillColor = Color.White,
                    HorzAlign = HorzAlign.Right,
                    VertAlign = VertAlign.Bottom,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportTitle.Objects.Add(
                new PictureObject()
                {
                    Name = "Picture1",
                    Top = 9.45F,
                    Width = 160.65F,
                    Height = 56.7F,
                    FillColor = Color.White,
                    Image = System.Drawing.Image.FromFile(@"C:\Users\mario.chan\Documents\GitHub\Benchmark\Library_benchmark\Content\images\CemexPDF.PNG")
                });

            // create Page Header
            page1.PageHeader = new PageHeaderBand()
            {
                Name = "PageHeader1",
                Top = 79.6F,
                Width = 718.2F,
                Height = 85.05F
            };
            page1.PageHeader.Objects.Add(
                new TextObject()
                {
                    Name = "Title",
                    Width = 718.2F,
                    Height = 47.25F,
                    FillColor = Color.Black,
                    Text = "Informacion.Title",
                    HorzAlign = HorzAlign.Center,
                    VertAlign = VertAlign.Bottom,
                    Font = new System.Drawing.Font("Arial", 18, FontStyle.Regular),
                    TextColor = Color.White
                });
            page1.PageHeader.Objects.Add(
                new TextObject()
                {
                    Name = "SubTitle",
                    Top = 47.25F,
                    Width = 718.2F,
                    Height = 37.8F,
                    FillColor = Color.Black,
                    Text = "Informacion.SubTitle",
                    HorzAlign = HorzAlign.Center,
                    Font = new System.Drawing.Font("Arial", 11.5F, FontStyle.Regular),
                    TextColor = Color.White
                });

            // create Report Summary
            page1.ReportSummary = new ReportSummaryBand()
            {
                Name = "ReportSummary1",
                Top = 168.65F,
                Width = 718.2F,
                Height = 831.6F
            };
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape1",
                    Width = 189F,
                    Height = 94.5F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape2",
                    Left = 179.55F,
                    Width = 189F,
                    Height = 94.5F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape3",
                    Left = 359.1F,
                    Width = 189,
                    Height = 94.5F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape4",
                    Left =  538.65F ,
                    Width =  179.55F ,
                    Height =  94.5F ,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Date1",
                    Left = 18.9F,
                    Top = 47.25F,
                    Width = 141.75F,
                    Height = 18.9F,
                    Text = "Informacion.Date",
                    HorzAlign = HorzAlign.Right,
                    VertAlign = VertAlign.Bottom,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text5",
                    Left = 14.45F,
                    Top = 25.9F,
                    Width = 151.2F,
                    Height = 18.9F,
                    Text = "DATE",
                    VertAlign = VertAlign.Bottom,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text6",
                    Left = 198.45F,
                    Top = 18.9F,
                    Width = 151.2F,
                    Height = 18.9F,
                    Text = "DRIVER",
                    VertAlign = VertAlign.Bottom,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text7",
                    Left = 375,
                    Top = 30.35F,
                    Width = 151.2F,
                    Height = 18.9F,
                    Text = "TRUCK",
                    VertAlign = VertAlign.Bottom,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text8",
                    Left = 548.1F,
                    Top = 28.35F,
                    Width = 151.2F,
                    Height = 18.9F,
                    Text = "HOURS",
                    VertAlign = VertAlign.Bottom,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Driver",
                    Left = 198.45F,
                    Top = 37.8F,
                    Width = 151.2F,
                    Height = 37.8F,
                    Text = "Informacion.Driver",
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Truck",
                    Left = 378,
                    Top = 47.25F,
                    Width = 151.2F,
                    Height = 18.9F,
                    Text = "Informacion.Truck",
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Date",
                    Left = 548.1F,
                    Top = 47.25F,
                    Width = 151.2F,
                    Height = 18.9F,
                    Text = "Informacion.Date",
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape5",
                    Top =  85.05f ,
                    Width =  718.2f ,
                    Height =  37.8f ,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(223, 223, 223)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text12",
                    Top = 122.85F,
                    Width = 718.2F,
                    Height = 37.8F,
                    Border = new Border()
                    {
                        Lines = BorderLines.Left | BorderLines.Right,
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239),
                    Text = "      GENERAL INFORMATION",
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape6",
                    Top = 160.65F,
                    Width = 718.2F,
                    Height = 103.95F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.White

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text13",
                    Left = 27.65F,
                    Top = 187.35F,
                    Width = 141.75F,
                    Height = 18.9F,
                    FillColor = Color.White,
                    Text = "Odometer Start ",
                    HorzAlign = HorzAlign.Center,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text14",
                    Left = 207.9F,
                    Top = 189,
                    Width = 179.55F,
                    Height = 18.9F,
                    FillColor = Color.White,
                    Text = "Max Air Pressure PSI",
                    HorzAlign = HorzAlign.Center,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text15",
                    Left = 434.7F,
                    Top = 189,
                    Width = 264.6F,
                    Height = 18.9F,
                    FillColor = Color.White,
                    Text = "Low Air Warning Device PSI",
                    HorzAlign = HorzAlign.Center,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "OdometerStatr",
                    Left = 47.25F,
                    Top = 207.9F,
                    Width = 103.95F,
                    Height = 18.9F,
                    Text = "Informacion.OdometerStatr",
                    FillColor = Color.White,
                    HorzAlign = HorzAlign.Center,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "PressurePsi",
                    Left = 207.9F,
                    Top = 207.9F,
                    Width = 179.55F,
                    Height = 18.9F,
                    Text = "Informacion.PressurePsi",
                    FillColor = Color.White,
                    HorzAlign = HorzAlign.Center,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "DevicePsi",
                    Left = 453.6F,
                    Top = 207.9F,
                    Width = 236.25F,
                    Height = 18.9F,
                    Text = "Informacion.DevicePsi",
                    FillColor = Color.White,
                    HorzAlign = HorzAlign.Center,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape7",
                    Top = 255.15F,
                    Width = 718.2F,
                    Height = 37.8F,
                    FillColor = Color.FromArgb(223, 223, 223),
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    }

                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text19",
                    Top = 292.95F,
                    Width = 718.2F,
                    Height = 37.8F,
                    Border = new Border()
                    {
                        Lines = BorderLines.Left | BorderLines.Right,
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239),
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape8",
                    Top = 330.75F,
                    Width = 718.2F,
                    Height = 94.5F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.White

                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape9",
                    Top = 415.8F,
                    Width = 718.2F,
                    Height = 37.8F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(223, 223, 223)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text20",
                    Top = 453.6F,
                    Width = 718.2F,
                    Height = 37.8F,
                    Text = "      Mechanical",
                    Border = new Border()
                    {
                        Lines = BorderLines.Left | BorderLines.Right,
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239),
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape10",
                    Top = 491.4F,
                    Width = 718.2F,
                    Height = 103.95F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.White
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text21",
                    Left = 18.9F,
                    Top = 349.65F,
                    Width = 94.5F,
                    Height = 18.9F,
                    Text = "Water",
                    FillColor = Color.White,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Water",
                    Left = 18.9F,
                    Top = 368.55F,
                    Width = 274.05F,
                    Height = 18.9F,
                    Text = "Informacion.Water",
                    FillColor = Color.White,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text23",
                    Left = 18.9F,
                    Top = 510.3F,
                    Width = 132.3F,
                    Height = 47.25F,
                    Text = "Leaks: Water, Oil, Fuel, Grease",
                    FillColor = Color.White,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "MechanicalComments",
                    Left = 18.9F,
                    Top = 557.55F,
                    Width = 264.6F,
                    Height = 18.9F,
                    Text = "Informacion.MechanicalComments",
                    FillColor = Color.White,
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape11",
                    Top = 585.9F,
                    Width = 718.2F,
                    Height = 37.8F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(223, 223, 223)
                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape12",
                    Top = 614.25F,
                    Width = 718.2F,
                    Height = 198.45F,
                    Border = new Border()
                    {
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text25",
                    Left = 9.45F,
                    Top = 633.15F,
                    Width = 94.5F,
                    Height = 18.9F,
                    Text = "Supervisor",
                    FillColor = Color.FromArgb(238, 239, 239),
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Text26",
                    Left = 604.8F,
                    Top = 633.15F,
                    Width = 94.5F,
                    Height = 18.9F,
                    Text = "DATE",
                    HorzAlign = HorzAlign.Right,
                    FillColor = Color.FromArgb(238, 239, 239),
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportSummary.Objects.Add(
                new TextObject()
                {
                    Name = "Date",
                    Left = 463.05F,
                    Top = 652.05F,
                    Width = 236.25F,
                    Height = 18.9F,
                    Text = "Informacion.Date",
                    HorzAlign = HorzAlign.Right,
                    FillColor = Color.FromArgb(238, 239, 239),
                    VertAlign = VertAlign.Center,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold)
                });
            page1.ReportTitle.Objects.Add(
                new PictureObject()
                {
                    Name = "Picture2",
                    Left = 189,
                    Top = 699.3F,
                    Width = 330.75F,
                    Height = 75.6F,
                    Border = new Border()
                    {
                        Lines = BorderLines.All,
                        Color = Color.LightGray,
                        Width = 3
                    },
                    FillColor = Color.FromArgb(238, 239, 239)
                });

            //// create group header
            //GroupHeaderBand group1 = new GroupHeaderBand();
            //group1.Name = "Cities Data";
            //group1.Height = Units.Centimeters * 1;
            //// set group condition
            ////group1.Condition = "[Cities.CityName]";//[Cities.CityName].Substring(0, 1)
            //// add group to the page.Bands collection
            //page1.Bands.Add(group1);
            //// create group footer
            //group1.GroupFooter = new GroupFooterBand();
            //group1.GroupFooter.Name = "GroupFooter1";
            //group1.GroupFooter.Height = Units.Centimeters * 1;
            // create DataBand
            //DataBand data1 = new DataBand();
            //data1.Name = "Data1";
            //data1.Height = Units.Centimeters * 0.5f;
            // set data source
            //data1.DataSource = report.GetDataSource("Cities");
            // connect databand to a group
            //group1.Data = data1;
            // create "Text" objects
            // report title
            //TextObject text1 = new TextObject();
            //text1.Name = "Text1";
            //// set bounds
            //text1.Bounds = new System.Drawing.RectangleF(0, 0,
            //Units.Centimeters * 19, Units.Centimeters * 1);
            //// set text
            //text1.Text = "CitiesData";
            //// set appearance
            //text1.HorzAlign = HorzAlign.Center;
            //text1.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Bold);
            //// add it to ReportTitle
            //page1.ReportTitle.Objects.Add(text1);
            //// group
            //TextObject text2 = new TextObject();
            //text2.Name = "Text2";
            //text2.Bounds = new RectangleF(0, 0,
            //Units.Centimeters * 2, Units.Centimeters * 1);
            //// text2.Text = "[Cities.CityName]";//[Cities.CityName].Substring(0, 1)
            //text2.Font = new System.Drawing.Font("Tahoma", 10, FontStyle.Bold);
            // add it to GroupHeader
            //group1.Objects.Add(text2);
            report.Show();
        }

        private Report CreateDynamicFormObject(Report report)
        {

            var informacion = new Consultas().GetPdfInformacion();

            PropertyInfo[] properties = informacion.GetType().GetProperties();
            foreach (var prop in properties)
            {
                //report.RegisterData((string)prop.GetValue(informacion, null).ToString(), prop.Name);
                //report.GetDataSource(prop.Name).Enabled = true;
                //report.GetDataSource(prop.Name).InitSchema();

                FastReport.TextObject myText = (FastReport.TextObject)report.FindObject(prop.Name);
                myText.Text = (string)prop.GetValue(informacion, null).ToString();
            }




            return report;
        }
    }
}