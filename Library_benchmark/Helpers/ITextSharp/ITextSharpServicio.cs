using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;

namespace Library_benchmark.Helpers.ITextSharp
{
    public class ITextSharpServicio
    {
        Document doc;
        PdfPTable tableLayout;
        IList<Models.Dummy> informacion;
        int sheets;
        private IEnumerable<PropertyInfo> cabeceras;
        MemoryStream workStream;
        private PdfWriter writer;

        public ITextSharpServicio(IList<Models.Dummy> informacion, MemoryStream workStream, int sheets)
        {
            this.informacion = informacion;
            this.sheets = sheets;
            this.workStream = workStream;

            //CreateDoc();
            //AbrirDocumento();
            //prueba();
            //CerrarDocumento();

            CreateDoc();
            CreateTableLayout();
            AbrirDocumento();
            CrearTitulo();
            CrearCabeceras();
            CrearContenido();
            ADDContenido();
            CerrarDocumento();
        }

        private void prueba()
        {
            
            iTextSharp.text.Font mainFont = FontFactory.GetFont("Segoe UI", 22, new iTextSharp.text.BaseColor(System.Drawing.ColorTranslator.FromHtml("#999")));
            iTextSharp.text.Font infoFont1 = FontFactory.GetFont("Kalinga", 10, new iTextSharp.text.BaseColor(System.Drawing.ColorTranslator.FromHtml("#666")));
            iTextSharp.text.Font expHeadFond = FontFactory.GetFont("Calibri (Body)", 12, new iTextSharp.text.BaseColor(System.Drawing.ColorTranslator.FromHtml("#666")));
            PdfContentByte contentByte = writer.DirectContent;
            
            ColumnText ct = new ColumnText(contentByte);
            //Create the font for show the name of user  
            
            PdfPTable modelInfoTable = new PdfPTable(1);
            modelInfoTable.TotalWidth = 100f;
            modelInfoTable.HorizontalAlignment = Element.ALIGN_LEFT;
            PdfPCell modelInfoCell1 = new PdfPCell()
            {
                BorderWidthBottom = 0f,
                BorderWidthTop = 0f,
                BorderWidthLeft = 0f,
                BorderWidthRight = 0f
            };
            //Set right hand the first heading  
            Phrase mainPharse = new Phrase();
            Chunk mChunk = new Chunk("Mario Enrique Chan Fernandez", mainFont);
            mainPharse.Add(mChunk);
            mainPharse.Add(new Chunk(Environment.NewLine));
            //Set the user role  
            Chunk infoChunk1 = new Chunk("Profile - Admin", infoFont1);
            mainPharse.Add(infoChunk1);
            mainPharse.Add(new Chunk(Environment.NewLine));
            //Set the user Gender  
            Chunk infoChunk21 = new Chunk("Gender - MAsculino", infoFont1);
            mainPharse.Add(infoChunk21);
            mainPharse.Add(new Chunk(Environment.NewLine));
            //Set the user age  
            Chunk infoChunk22 = new Chunk("Age - 19", infoFont1);
            mainPharse.Add(infoChunk22);
            mainPharse.Add(new Chunk(Environment.NewLine));
            iTextSharp.text.Font infoFont2 = FontFactory.GetFont("Kalinga", 10, new iTextSharp.text.BaseColor(System.Drawing.ColorTranslator.FromHtml("#848282")));
            string Location = "MTY, Mexico";
            Chunk infoChunk2 = new Chunk("Address -" + Location, infoFont2);
            mainPharse.Add(infoChunk2);
            modelInfoCell1.AddElement(mainPharse);
            //Set the mobile image and number  
            Phrase mobPhrase = new Phrase();
            // iTextSharp.text.Image mobileImage = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/goodmorning.jpg"));  
            // mobileImage.ScaleToFit(10, 10);  
            //Chunk cmobImg = new Chunk(mobileImage, 0, -2);  
            Chunk cmob = new Chunk("Contact - 3311295428", infoFont2);
            //mobPhrase.Add(cmobImg);  
            mobPhrase.Add(cmob);
            modelInfoCell1.AddElement(mobPhrase);
            //Set the message image and email id  
            Phrase msgPhrase = new Phrase();
            // iTextSharp.text.Image msgImage = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/goodmorning.jpg"));  
            //msgImage.ScaleToFit(10, 10);  
            //Chunk msgImg = new Chunk(msgImage, 0, -2);  
            iTextSharp.text.Font msgFont = FontFactory.GetFont("Kalinga", 10, new iTextSharp.text.BaseColor(System.Drawing.Color.Pink));
            Chunk cmsg = new Chunk("EMail - enrique.nec@gmail.com", msgFont);
            //msgPhrase.Add(msgImg);  
            msgPhrase.Add(cmsg);
            //Set the line after the user small information  
            iTextSharp.text.Font lineFont = FontFactory.GetFont("Kalinga", 10, new iTextSharp.text.BaseColor(System.Drawing.ColorTranslator.FromHtml("#e8e8e8")));
            Chunk lineChunk = new Chunk("____________________________________________________________________", lineFont);
            msgPhrase.Add(new Chunk(Environment.NewLine));
            msgPhrase.Add(lineChunk);
            modelInfoCell1.AddElement(msgPhrase);
            modelInfoTable.AddCell(modelInfoCell1);
            //Set the biography  
            PdfPCell cell1 = new PdfPCell()
            {
                BorderWidthBottom = 0f,
                BorderWidthTop = 0f,
                BorderWidthLeft = 0f,
                BorderWidthRight = 0f
            };
            cell1.PaddingTop = 5f;
            Phrase bioPhrase = new Phrase();
            Chunk bioChunk = new Chunk("Biography", mainFont);
            bioPhrase.Add(bioChunk);
            bioPhrase.Add(new Chunk(Environment.NewLine));
            Chunk bioInfoChunk = new Chunk("Algo muy extraño va aqui", infoFont1);
            bioPhrase.Add(bioInfoChunk);
            bioPhrase.Add(new Chunk(Environment.NewLine));
            bioPhrase.Add(lineChunk);
            cell1.AddElement(bioPhrase);
            modelInfoTable.AddCell(cell1);
            PdfPCell cellExp = new PdfPCell()
            {
                BorderWidthBottom = 0f,
                BorderWidthTop = 0f,
                BorderWidthLeft = 0f,
                BorderWidthRight = 0f
            };
            cellExp.PaddingTop = 5f;
            Phrase ExperiencePhrase = new Phrase();
            Chunk ExperienceChunk = new Chunk("Experience", mainFont);
            ExperiencePhrase.Add(ExperienceChunk);
            cellExp.AddElement(ExperiencePhrase);
            modelInfoTable.AddCell(cellExp);

            for (int i = 0; i < 50; i++)
            {
                //Set the experience  
                PdfPCell expcell = new PdfPCell()
                {
                    BorderWidthBottom = 0f,
                    BorderWidthTop = 0f,
                    BorderWidthLeft = 0f,
                    BorderWidthRight = 0f
                };
                expcell.PaddingTop = 5f;
                Phrase expPhrase = new Phrase();
                StringBuilder expStringBuilder = new StringBuilder();
                StringBuilder expStringBuilder1 = new StringBuilder();
                //Set the experience details  
                expStringBuilder.Append("Title - " + i + Environment.NewLine);
                expStringBuilder.Append("CompanyName - " + i + Environment.NewLine);
                expStringBuilder.Append("ComanyAddress - " + i + Environment.NewLine);
                expStringBuilder1.Append("From " + i + " To " + i + Environment.NewLine);
                // expPhrase.Add(new Chunk(Environment.NewLine));  
                Chunk expDetailChunk = new Chunk(expStringBuilder.ToString(), expHeadFond);
                expPhrase.Add(expDetailChunk);
                expPhrase.Add(new Chunk(expStringBuilder1.ToString(), infoFont2));
                expcell.AddElement(expPhrase);
                modelInfoTable.AddCell(expcell);
                string description = "que show men esto es una descripcion bien perruquis de la vida moderna";
                if (description.Length > 600)
                {
                    PdfPCell pCell1 = new PdfPCell()
                    {
                        BorderWidth = 0f
                    };
                    PdfPCell pCell2 = new PdfPCell()
                    {
                        BorderWidth = 0f
                    };
                    Phrase ph1 = new Phrase();
                    Phrase ph2 = new Phrase();
                    string experience1 = description.Substring(0, 599);
                    string experience2 = description.Substring(599, description.Length - 600);
                    ph1.Add(new Chunk(experience1, infoFont1));
                    ph2.Add(new Chunk(experience2, infoFont1));
                    pCell1.AddElement(ph1);
                    pCell2.AddElement(ph2);
                    modelInfoTable.AddCell(pCell1);
                    modelInfoTable.AddCell(pCell2);
                }
                else
                {
                    PdfPCell pCell1 = new PdfPCell()
                    {
                        BorderWidth = 0f
                    };
                    Phrase ph1 = new Phrase();
                    string experience1 = description;
                    ph1.Add(new Chunk(experience1, infoFont1));
                    pCell1.AddElement(ph1);
                    modelInfoTable.AddCell(pCell1);
                }
            }

            doc.Add(modelInfoTable);
            //Set the footer  
            PdfPTable footerTable = new PdfPTable(1);
            footerTable.TotalWidth = 644f;
            footerTable.LockedWidth = true;
            PdfPCell footerCell = new PdfPCell(new Phrase("Resume"));
            footerCell.BackgroundColor = new iTextSharp.text.BaseColor(System.Drawing.Color.Black);
            iTextSharp.text.Image footerImage = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/Content/images/goodmorning.jpg"));
            footerImage.SpacingBefore = 5f;
            footerImage.SpacingAfter = 5f;
            footerImage.ScaleToFit(100f, 22f);
            footerCell.AddElement(footerImage);
            footerCell.MinimumHeight = 30f;
            iTextSharp.text.Font newFont = FontFactory.GetFont("Segoe UI, Lucida Grande, Lucida Grande", 8, new iTextSharp.text.BaseColor(System.Drawing.Color.White));
            Paragraph rightReservedLabel = new Paragraph("© " + DateTime.Now.Year + " Resume. All rights reserved.", newFont);
            footerCell.AddElement(rightReservedLabel);
            footerCell.PaddingLeft = 430f;
            footerTable.AddCell(footerCell);
            footerTable.WriteSelectedRows(0, -1, 0, doc.PageSize.Height - 795, writer.DirectContent);
        }

        private void AbrirDocumento()
        {
            
            writer = PdfWriter.GetInstance(doc, workStream );
            writer.CloseStream = false;
            doc.Open();

            //Add Content to PDF   

        }

        private void CrearTitulo()
        {

            PdfPTable table = new PdfPTable(1);
            table.WidthPercentage = 100;
            PdfPTable table2 = new PdfPTable(2);

            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance("C:/Users/mario.chan/Documents/GitHub/Library_benchmark/Library_benchmark/Content/images/net.png");
            image.ScalePercent(7f);

            image.SetAbsolutePosition(doc.PageSize.Width - 36f - 72f,
                  doc.PageSize.Height - 36f - 216.6f);
            PdfPCell cell2 = new PdfPCell(image);
            cell2.Colspan = 2;
            cell2.Border = 0;
            table2.AddCell(cell2);

            cell2 = new PdfPCell(new Phrase("\nTITLE TEXT", new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD | Font.UNDERLINE)));
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.Colspan = 2;
            table2.AddCell(cell2);

            PdfPCell cell = new PdfPCell(table2);
            cell.Border = 0;
            table.HeaderRows = 1;
            table.AddCell(cell);
            table.AddCell(new PdfPCell(new Phrase("")));

            doc.Add(table);

        }

        internal Document GetPDFExample()
        {
            return doc;
        }

        private void CrearCabeceras()
        {
            //float[] headers = { 50, 24, 45, 35, 50 }; //Header Widths  
            //tableLayout.SetWidths(headers); //Set the pdf headers  
            tableLayout.WidthPercentage = 100; //Set the PDF File witdh percentage  
            tableLayout.HeaderRows = 1;



            var item = informacion.FirstOrDefault();
            cabeceras = item.GetType().GetProperties().Where(p => !p.GetGetMethod().GetParameters().Any());
            foreach (var prop in cabeceras)
            {
                //tableLayout.AddCell(new PdfPCell(new Phrase(prop.Name.ToString(), new Font(Font.FontFamily.HELVETICA, 8, 1, iTextSharp.text.BaseColor.YELLOW)))
                //{
                //    HorizontalAlignment = Element.ALIGN_LEFT,
                //    Padding = 5,
                //    BackgroundColor = new iTextSharp.text.BaseColor(128, 0, 0)
                //});

            }

        }

        private void CrearContenido()
        {
            foreach (var item in informacion)
            {

                foreach (var prop in cabeceras)
                {
                    //tableLayout.AddCell(new PdfPCell(new Phrase(prop.GetValue(item, null).ToString(), new Font(Font.FontFamily.HELVETICA, 8, 1, iTextSharp.text.BaseColor.BLACK)))
                    //{
                    //    HorizontalAlignment = Element.ALIGN_LEFT,
                    //    Padding = 5,
                    //    BackgroundColor = new iTextSharp.text.BaseColor(255, 255, 255)
                    //});
                }

            }

        }

        private void ADDContenido()
        {
            doc.Add(tableLayout);
        }

        private void CerrarDocumento()
        {

            doc.Close();
        }

        private void CreateDoc()
        {
            doc = new Document(PageSize.A4, 0f, 0f, 0f, 0f);
        }

        private void CreateTableLayout()
        {
            var item = informacion.FirstOrDefault();
            int columnas = item.GetType().GetProperties().Where(p => !p.GetGetMethod().GetParameters().Any()).Count();
            if (columnas > 0)
            {
                tableLayout = new PdfPTable(columnas);
            }


        }


    }
}