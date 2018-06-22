using System;
using System.Drawing;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Library_benchmark.Helpers
{
    internal class NPOIDesign
    {
        private XSSFWorkbook excel;
        private bool resource;
        private int rowInicial;

        private ICellStyle headerStyle;
        private ICellStyle normalStyle;
        private ICellStyle dateStyle;
        private ICellStyle cabeceraStyle;



        /// <summary>
        /// 
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="resource"></param>
        public NPOIDesign(XSSFWorkbook excel, bool resource)
        {
            this.excel = excel;
            this.resource = resource;
            rowInicial = 4;

            DarFormato();
        }


        /// <summary>
        /// 
        /// </summary>
        private void DarFormato()
        {
            if (excel != null)
            {
                if (!resource)
                    PutImagenTitulo();

                //PutTypeAndSizeText();
                PutCabeceras();
                //PutCeldasNormales();
                PutFitInCells();

            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void PutImagenTitulo()
        {
            for (int i = 0; i < excel.NumberOfSheets; i++)
            {
                DiseñoCabeceras(i);
            }


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        private void DiseñoCabeceras(int sheet)
        {
            var pestana = excel.GetSheetAt(sheet);

            Image image = Image.FromFile(@"C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/Cemex.png");
            for (int i = 0; i < 2; i++)
            {
                IRow row;
                row = pestana.GetRow(i);
                if (row == null)
                {
                    row = pestana.CreateRow(i);
                }

                row.HeightInPoints = 51f;


                for (int j = 7; j < 13; j++)
                {
                    if (row.GetCell(j) == null)
                    {
                        row.CreateCell(j);
                    }

                }
            }

            var cra = new CellRangeAddress(0, 1, 8, 12);


            pestana.AddMergedRegion(cra);

            if (cabeceraStyle == null)
                cabeceraStyle = GetCabeceraCellStyle();
            var celda = pestana.GetRow(0).GetCell(8);
            celda.SetCellValue("NPOI");
            celda.CellStyle = cabeceraStyle;


            ////row0.HeightInPoints = (float)image.Height;
            var converter = new ImageConverter();
            var data = (byte[])converter.ConvertTo(image, typeof(byte[]));
            //var pictureIndex = excel.AddPicture(data, PictureType.PNG);
            //var helper = excel.GetCreationHelper();
            //var drawing = pestana.CreateDrawingPatriarch();
            //var anchor = new XSSFClientAnchor(900, 0, 0, 0, 1, 1, 7, 2);


            //var picture = drawing.CreatePicture(anchor, pictureIndex);

            //picture.Resize(1);

            int myPictureId = excel.AddPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG);

            IDrawing drawing = pestana.CreateDrawingPatriarch();
            XSSFClientAnchor myAnchor = new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 7, 2);



            IPicture myPicture = drawing.CreatePicture(myAnchor, myPictureId);
            myPicture.Resize();

        }

        /// <summary>
        /// 
        /// </summary>
        private void PutCabeceras()
        {

            if (headerStyle == null)
                headerStyle = GetHeaderCellStyle();


            for (int i = 0; i < excel.NumberOfSheets; i++)
            {
                if (excel.GetSheetAt(i).GetRow(rowInicial - 1) != null)
                {
                    foreach (var item in excel.GetSheetAt(i).GetRow(rowInicial - 1).Cells)
                    {
                        item.CellStyle = headerStyle;

                    }
                }

            }

        }

        

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private ICellStyle GetHeaderCellStyle()
        {
            var style = excel.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.DarkBlue.Index; ;
            style.FillPattern = FillPattern.SolidForeground;

            IFont hfont = excel.CreateFont();
            hfont.FontHeightInPoints = 12;
            hfont.Color = IndexedColors.White.Index;
            hfont.FontName = "Century Gothic";
            style.SetFont(hfont);

            return style;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private ICellStyle GetCabeceraCellStyle()
        {
            ICellStyle style = excel.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.DarkBlue.Index; ;
            style.FillPattern = FillPattern.SolidForeground;

            style.FillBackgroundColor = IndexedColors.DarkBlue.Index;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            IFont hfont = excel.CreateFont();
            hfont.FontHeightInPoints = 36;
            hfont.Color = IndexedColors.White.Index;
            hfont.FontName = "Century Gothic";
            hfont.IsBold = true;
            style.SetFont(hfont);

            return style;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private ICellStyle GetNormalCellStyle()
        {
            ICellStyle style = excel.CreateCellStyle();
            var hfont = excel.CreateFont();
            hfont.FontHeightInPoints = 12;
            hfont.Color = IndexedColors.Black.Index;
            hfont.FontName = "Century Gothic";


            style.SetFont(hfont);
            return style;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private ICellStyle GetDateCellStyle()
        {
            var style = excel.CreateCellStyle();

            var hfont = excel.CreateFont();
            hfont.FontHeightInPoints = 13;
            hfont.Color = IndexedColors.Black.Index;
            hfont.IsBold = true;
            style.SetFont(hfont);
            style.DataFormat = excel.CreateDataFormat().GetFormat("");

            return style;
        }

        /// <summary>
        /// 
        /// </summary>
        private void PutFitInCells()
        {


            for (int i = 0; i < excel.NumberOfSheets; i++)
            {
                if (excel.GetSheetAt(i).GetRow(rowInicial) != null)
                {
                    int noOfColumns = excel.GetSheetAt(i).GetRow(rowInicial).LastCellNum;
                    for (int j = 0; j < noOfColumns; j++)
                    {
                        excel.GetSheetAt(i).AutoSizeColumn(j);
                    }
                }


            }

        }

        /// <summary>
        /// 
        /// </summary>
        private void PutTypeAndSizeText()
        {
            if (normalStyle == null)
                normalStyle = GetNormalCellStyle();

            for (int i = 0; i < excel.NumberOfSheets; i++)
            {
                if (excel.GetSheetAt(i).GetRow(rowInicial) != null)
                {
                    for (int j = 0; j < excel.GetSheetAt(i).GetRow(rowInicial).Cells.Count; j++)
                    {

                        excel.GetSheetAt(i).SetDefaultColumnStyle(j, normalStyle);
                    }
                }


            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        internal XSSFWorkbook GetExcelExample()
        {
            return excel;
        }
    }
}