using FastReport;
using Library_benchmark.Models;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;

namespace Library_benchmark.Helpers.Fast
{
    public class FastReportServicio
    {
        private readonly PdfDummy informacion;
        private Report report;
        private Thread thread;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="informacion"></param>
        public FastReportServicio(PdfDummy informacion)
        {
            Report report = new Report();
            this.informacion = informacion;


            thread = new Thread(new ThreadStart(Reporte));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pdfDummy"></param>
        /// <param name="informacion"></param>
        public FastReportServicio(byte[] pdfDummy, PdfDummy informacion)
        {
            var fs = new MemoryStream(pdfDummy);
            report = Report.FromStream(fs);

            this.informacion = informacion;

            thread = new Thread(new ThreadStart(Template));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

        }

        /// <summary>
        /// 
        /// </summary>
        private void Reporte()
        {
            ReportPage page1 = CreateReportPage();

            ReportTile(page1);
            PageHeader(page1);

            ReportSummary(page1);

            report.Show();
            //thread.Suspend();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private ReportPage CreateReportPage()
        {
            ReportPage page1 = new ReportPage()
            {
                Name = "Page1"
            };
            report.Pages.Add(page1);
            return page1;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="page1"></param>
        private void ReportSummary(ReportPage page1)
        {
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
                    Left = 538.65F,
                    Width = 179.55F,
                    Height = 94.5F,
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
                    Text = informacion.Date,
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
                    Text = informacion.Driver,
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
                    Text = informacion.Truck,
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
                    Text = informacion.Date,
                    Font = new System.Drawing.Font("Calibri", 11, FontStyle.Bold),

                });
            page1.ReportSummary.Objects.Add(
                new ShapeObject()
                {
                    Name = "Shape5",
                    Top = 85.05f,
                    Width = 718.2f,
                    Height = 37.8f,
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
                    Text = informacion.OdometerStatr,
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
                    Text = informacion.PressurePsi,
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
                    Text = informacion.DevicePsi,
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
                    Text = informacion.Water,
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
                    Text = informacion.MechanicalComments,
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
                    Text = informacion.Date,
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
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="page1"></param>
        private void PageHeader(ReportPage page1)
        {
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
                    Text = informacion.Title,
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
                    Text = informacion.SubTitle,
                    HorzAlign = HorzAlign.Center,
                    Font = new System.Drawing.Font("Arial", 11.5F, FontStyle.Regular),
                    TextColor = Color.White
                });
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="page1"></param>
        private void ReportTile(ReportPage page1)
        {
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
                    Text = informacion.Truck,
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
        }

        /// <summary>
        /// 
        /// </summary>
        private void Template()
        {
            PropertyInfo[] properties = informacion.GetType().GetProperties();
            foreach (var prop in properties)
            {
                //report.RegisterData((string)prop.GetValue(informacion, null).ToString(), prop.Name);
                //report.GetDataSource(prop.Name).Enabled = true;
                //report.GetDataSource(prop.Name).InitSchema();

                FastReport.TextObject myText = (FastReport.TextObject)report.FindObject(prop.Name);
                myText.Text = (string)prop.GetValue(informacion, null).ToString();
            }

            report.Show();

            thread.Suspend();

        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        internal Report GetExcelExample()
        {

            return report;
        }
    }
}