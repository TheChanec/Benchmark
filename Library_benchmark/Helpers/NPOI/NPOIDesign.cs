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

        ICellStyle headerStyle;
        ICellStyle normalStyle;
        ICellStyle dateStyle;
        ICellStyle cabeceraStyle;




        public NPOIDesign(XSSFWorkbook excel, bool resource)
        {
            this.excel = excel;
            this.resource = resource;
            rowInicial = 4;

            DarFormato();
        }












        private void DarFormato()
        {
            if (excel != null)
            {
                if (!resource)
                    PutImagenTitulo();

                PutTypeAndSizeText();
                PutCabeceras();
                PutCeldasNormales();
                PutFitInCells();

            }
        }

        private void PutImagenTitulo()
        {
            for (int i = 0; i < excel.NumberOfSheets; i++)
            {
                DiseñoCabeceras(i);
            }


        }

        private void DiseñoCabeceras(int sheet)
        {
            var pestana = excel.GetSheetAt(sheet);

            Image image = Image.FromFile(@"C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/net.png");
            for (int i = 0; i < 6; i++)
            {
                IRow row;
                row = pestana.GetRow(i);
                if (row == null)
                {
                    row = pestana.CreateRow(i);
                }




                for (int j = 0; j < 16; j++)
                {
                    if (row.GetCell(j) == null)
                    {
                        row.CreateCell(j);
                    }

                }
            }

            var cra = new CellRangeAddress(0, 5, 2, 15);


            pestana.AddMergedRegion(cra);

            if (cabeceraStyle == null)
                cabeceraStyle = GetCabeceraCellStyle();
            var celda = pestana.GetRow(0).GetCell(2);
            celda.SetCellValue("NPOI");
            celda.CellStyle = cabeceraStyle;


            //row0.HeightInPoints = (float)image.Height;
            var converter = new ImageConverter();
            var data = (byte[])converter.ConvertTo(image, typeof(byte[]));
            var pictureIndex = excel.AddPicture(data, PictureType.PNG);
            var helper = excel.GetCreationHelper();
            var drawing = pestana.CreateDrawingPatriarch();
            var anchor = new XSSFClientAnchor(0, 0, 600, 0, 0, 0, 1, 6);


            var picture = drawing.CreatePicture(anchor, pictureIndex);

            picture.Resize(1);



        }

        internal void PutCabeceras()
        {

            if (headerStyle == null)
                headerStyle = GetHeaderCellStyle();


            for (int i = 0; i < excel.NumberOfSheets; i++)
            {
                if (excel.GetSheetAt(i).GetRow(rowInicial - 1)!= null)
                {
                    foreach (var item in excel.GetSheetAt(i).GetRow(rowInicial - 1).Cells)
                    {
                        item.CellStyle = headerStyle;

                    }
                }
                
            }

        }

        internal void PutCeldasNormales()
        {

            if (normalStyle == null)
                normalStyle = GetNormalCellStyle();


            for (int i = 0; i < excel.NumberOfSheets; i++)
            {


                for (int j = rowInicial; j <= excel.GetSheetAt(i).LastRowNum; j++)
                {
                    foreach (var item in excel.GetSheetAt(i).GetRow(j).Cells)
                    {
                        var style = item.CellStyle;
                        item.CellStyle = normalStyle;

                    }
                }


            }

        }

        private ICellStyle GetHeaderCellStyle()
        {
            var style = excel.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.DarkBlue.Index; ;
            style.FillPattern = FillPattern.SolidForeground;
            style.FillBackgroundColor = IndexedColors.DarkBlue.Index;

            var hfont = excel.CreateFont();
            hfont.FontHeightInPoints = 13;
            hfont.Color = IndexedColors.White.Index;
            style.SetFont(hfont);

            return style;
        }

        private ICellStyle GetCabeceraCellStyle()
        {
            ICellStyle style = excel.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.DarkBlue.Index; ;
            style.FillPattern = FillPattern.SolidForeground;
            
            style.FillBackgroundColor = IndexedColors.DarkBlue.Index;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            IFont hfont = excel.CreateFont();
            hfont.FontHeightInPoints = 72;
            hfont.Color = IndexedColors.White.Index;
            hfont.FontName = "Arial";
            style.SetFont(hfont);

            return style;
        }

        private ICellStyle GetNormalCellStyle()
        {
            ICellStyle style = excel.CreateCellStyle();
            var hfont = excel.CreateFont();
            hfont.FontHeightInPoints = 12;
            hfont.Color = IndexedColors.Black.Index;
            hfont.FontName = "Arial";


            style.SetFont(hfont);
            return style;
        }

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

        internal void PutFitInCells()
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

        internal void PutTypeAndSizeText()
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

        internal XSSFWorkbook GetExcelExample()
        {
            return excel;
        }
    }
}