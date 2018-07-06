using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Web;

namespace Library_benchmark.Helpers.ITextSharp
{
    /// <inheritdoc />
    /// <summary>
    /// Clase encargada de la generacion de Headers y Footers para el PDF
    /// </summary>
    public class TextEvents : PdfPageEventHelper
    {
        private PdfContentByte _contentByte;
        private PdfTemplate _headerTemplate, _footerTemplate;
        private BaseFont _baseFont;

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {

            _baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            _contentByte = writer.DirectContent;
            _headerTemplate = _contentByte.CreateTemplate(560, 100);
            _footerTemplate = _contentByte.CreateTemplate(50, 50);

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
            var text = "Pagina " + writer.PageNumber + " de ";
            //string


            //Add paging to header
            {
                _contentByte.BeginText();
                _contentByte.SetFontAndSize(_baseFont, 12);
                _contentByte.SetTextMatrix(document.PageSize.GetRight(0), document.PageSize.GetTop(0));
                _contentByte.ShowText(text);
                _contentByte.EndText();

                var len = _baseFont.GetWidthPoint(text, 12);
                //Adds "12" in Page 1 of 12
                _contentByte.AddTemplate(_headerTemplate, document.PageSize.GetRight(0) + len, document.PageSize.GetTop(0));
            }

            {
                var numberOfPages = new Paragraph(text + " ", time) { Alignment = Element.ALIGN_CENTER };
                var footerTbl = new PdfPTable(1) { TotalWidth = 520f, HorizontalAlignment = Element.ALIGN_CENTER, LockedWidth = true };

                var sep = new PdfPCell
                {
                    Border = 0,
                    FixedHeight = 10f,
                    Colspan = 1
                };
                var footerCell1 = new PdfPCell(numberOfPages) { Border = 0 };

                footerTbl.AddCell(sep);
                footerTbl.AddCell(footerCell1);
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
            _contentByte.MoveTo(40, document.PageSize.GetBottom(50));
            _contentByte.LineTo(document.PageSize.Width - 40, document.PageSize.GetBottom(50));
            _contentByte.Stroke();
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            _headerTemplate.BeginText();
            _headerTemplate.SetFontAndSize(_baseFont, 12);
            _headerTemplate.SetTextMatrix(0, 0);
            _headerTemplate.ShowText((writer.PageNumber - 1).ToString());
            _headerTemplate.EndText();

            _footerTemplate.BeginText();
            _footerTemplate.SetFontAndSize(_baseFont, 12);
            _footerTemplate.SetTextMatrix(0, 0);
            _footerTemplate.ShowText((writer.PageNumber - 1).ToString());
            _footerTemplate.EndText();
        }
    }
}