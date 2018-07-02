﻿using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Web;

namespace Library_benchmark.Helpers.ITextSharp
{
    public class TextEvents : PdfPageEventHelper
    {
        private PdfContentByte _cb;
        private PdfTemplate _headerTemplate, _footerTemplate;
        private BaseFont _bf;

        public string Header { get; set; }
        
        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            //try
            //{
            _bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            _cb = writer.DirectContent;
            _headerTemplate = _cb.CreateTemplate(560, 100);
            _footerTemplate = _cb.CreateTemplate(50, 50);
            //}
            //catch (DocumentException de)
            //{

            //}
            //catch (System.IO.IOException ioe)
            //{

            //}
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);

            
            var fontFolio = FontFactory.GetFont(FontFactory.TIMES, 14, Font.NORMAL, BaseColor.BLACK);
            var time = FontFactory.GetFont(FontFactory.HELVETICA, 11f, Font.NORMAL);
            var logo = Image.GetInstance(HttpContext.Current.Server.MapPath("~/Content/images/CemexPDF.png"));
            //logo.ScaleToFit(200, 57);
            logo.ScalePercent(60);
            //document.Add(new Phrase(Environment.NewLine));

            //Create PdfTable object
            var pdfTab = new PdfPTable(3);
            //float[] width = { 100f, 320f, 100f };
            //pdfTab.SetWidths(width);
            //pdfTab.TotalWidth = 720f;
            //pdfTab.LockedWidth = true;
            //We will have to create separate cells to include image logo and 2 separate strings
            //Row 1
            var cellLogo = new PdfPCell(logo);
            var pdfCell2 = new PdfPCell();
            var text = "Page " + writer.PageNumber + " of ";
            //string


            //Add paging to header
            {
                _cb.BeginText();
                _cb.SetFontAndSize(_bf, 12);
                _cb.SetTextMatrix(document.PageSize.GetRight(0), document.PageSize.GetTop(0));
                _cb.ShowText(text);
                _cb.EndText();

                float len = _bf.GetWidthPoint(text, 12);
                //Adds "12" in Page 1 of 12
                _cb.AddTemplate(_headerTemplate, document.PageSize.GetRight(0) + len, document.PageSize.GetTop(0));
            }
            //Add paging to footer
            {
                //var leftCol = new Paragraph("Mukesh Salaria\n" + "Software Engineer", time);
                //var rightCol = new Paragraph("LearnShareCorner.com\n" + "Techical Blog", time);
                //var phone = new Paragraph("Phone +91-9814268272", time);
                float len = _bf.GetWidthPoint(text, 12);
                //Adds "12" in Page 1 of 12
                var algo = document.PageSize.GetRight(100) + len;
                var otracosa = document.PageSize.GetTop(45);

                var address = new Paragraph(text + " " + algo + " " + otracosa, time);
                //var fax = new Paragraph("mukeshsalaria01@gmail.com", time);

                //leftCol.Alignment = Element.ALIGN_LEFT;
                //rightCol.Alignment = Element.ALIGN_RIGHT;
                //fax.Alignment = Element.ALIGN_RIGHT;
                //phone.Alignment = Element.ALIGN_LEFT;
                address.Alignment = Element.ALIGN_CENTER;

                var footerTbl = new PdfPTable(3) { TotalWidth = 520f, HorizontalAlignment = Element.ALIGN_CENTER, LockedWidth = true };
                float[] widths = { 150f, 220f, 150f };
                footerTbl.SetWidths(widths);
                var footerCell1 = new PdfPCell(/*leftCol*/);
                var footerCell2 = new PdfPCell();
                var footerCell3 = new PdfPCell(/*rightCol*/);
                var sep = new PdfPCell();
                var footerCell4 = new PdfPCell(/*phone*/);
                var footerCell5 = new PdfPCell(address);
                var footerCell6 = new PdfPCell(/*fax*/);


                footerCell1.Border = 0;
                footerCell2.Border = 0;
                footerCell3.Border = 0;
                footerCell4.Border = 0;
                footerCell5.Border = 0;
                footerCell6.Border = 0;
                footerCell6.HorizontalAlignment = Element.ALIGN_RIGHT;
                sep.Border = 0;
                sep.FixedHeight = 10f;
                footerCell3.HorizontalAlignment = Element.ALIGN_RIGHT;
                footerCell6.PaddingLeft = 0;
                sep.Colspan = 3;

                footerTbl.AddCell(footerCell1);
                footerTbl.AddCell(footerCell2);
                footerTbl.AddCell(footerCell3);
                footerTbl.AddCell(sep);
                footerTbl.AddCell(footerCell4);
                footerTbl.AddCell(footerCell5);
                footerTbl.AddCell(footerCell6);
                footerTbl.WriteSelectedRows(0, -1, 40, 80, writer.DirectContent);
            }
            //Row 2
            // PdfPCell pdfCell4 = new PdfPCell(new Phrase("No job is so urgent that it cannot be done safely", baseFontNormal));
            //Row 3

            var cellFolio = new PdfPCell(new Phrase("CR4150", fontFolio));
            var pdfCell4 = new PdfPCell();
            //var pdfCell5 = new PdfPCell(new Phrase("TIME:" + string.Format("{0:t}", DateTime.Now), baseFontBig));
            var pdfCell5 = new PdfPCell(new Phrase(""));

            //set the alignment of all three cells and set border to 0
            cellLogo.HorizontalAlignment = Element.ALIGN_LEFT;
            cellFolio.HorizontalAlignment = Element.ALIGN_RIGHT;
            pdfCell5.HorizontalAlignment = Element.ALIGN_RIGHT;

            //pdfCell2.VerticalAlignment = Element.ALIGN_BOTTOM;

            //pdfCell1.Colspan = 3;
            //pdfCell2.Colspan = 3;
            pdfCell2.PaddingTop = 9f;
            cellFolio.PaddingTop = 8f;
            cellFolio.PaddingRight = 10f;

            pdfCell5.PaddingTop = 9f;

            cellLogo.Border = 0;
            pdfCell2.Border = 0;
            cellFolio.Border = 0;
            pdfCell4.Border = 0;
            pdfCell5.Border = 0;

            //add all three cells into PdfTable
            pdfTab.AddCell(cellLogo);
            pdfTab.AddCell(pdfCell2);
            pdfTab.AddCell(cellFolio);
            pdfTab.AddCell(pdfCell4);
            pdfTab.AddCell(pdfCell5);

            pdfTab.TotalWidth = 520f;
            pdfTab.LockedWidth = true;
            //pdfTab.TotalWidth = document.PageSize.Width;
            //pdfTab.WidthPercentage = 100;

            //call WriteSelectedRows of PdfTable. This writes rows from PdfWriter in PdfTable
            //first param is start row. -1 indicates there is no end row and all the rows to be included to write
            //Third and fourth param is x and y position to start writing
            pdfTab.WriteSelectedRows(0, -1, 40, document.PageSize.Height - 30, writer.DirectContent);
            //set pdfContent value

            //Move the pointer and draw line to separate header section from rest of page
            //cb.MoveTo(40, document.PageSize.Height - 100);
            //cb.LineTo(document.PageSize.Width - 40, document.PageSize.Height - 100);
            //cb.Stroke();

            //Move the pointer and draw line to separate footer section from rest of page
            _cb.MoveTo(40, document.PageSize.GetBottom(50));
            _cb.LineTo(document.PageSize.Width - 40, document.PageSize.GetBottom(50));
            _cb.Stroke();
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            _headerTemplate.BeginText();
            _headerTemplate.SetFontAndSize(_bf, 12);
            _headerTemplate.SetTextMatrix(0, 0);
            _headerTemplate.ShowText((writer.PageNumber - 1).ToString());
            _headerTemplate.EndText();

            _footerTemplate.BeginText();
            _footerTemplate.SetFontAndSize(_bf, 12);
            _footerTemplate.SetTextMatrix(0, 0);
            _footerTemplate.ShowText((writer.PageNumber - 1).ToString());
            _footerTemplate.EndText();
        }
    }
}