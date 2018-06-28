using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Drawing;

namespace Library_benchmark.Helpers.NPOI
{
    internal class NpoiDesign
    {
        private XSSFWorkbook _excel;
        private bool _resource;
        private int _rowInicial;

        private ICellStyle _headerStyle;
        private ICellStyle _normalStyle;
        private ICellStyle _cabeceraStyle;



        /// <summary>
        /// 
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="resource"></param>
        public NpoiDesign(XSSFWorkbook excel, bool resource)
        {
            _excel = excel;
            _resource = resource;
            _rowInicial = 4;

            DarFormato();
        }


        /// <summary>
        /// 
        /// </summary>
        private void DarFormato()
        {
            if (_excel == null) return;
            if (!_resource)
                PutImagenTitulo();

            //PutTypeAndSizeText();
            PutCabeceras();
            //PutCeldasNormales();
            PutFitInCells();
        }

        /// <summary>
        /// 
        /// </summary>
        private void PutImagenTitulo()
        {
            for (var i = 0; i < _excel.NumberOfSheets; i++)
                DiseñoCabeceras(i);

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        private void DiseñoCabeceras(int sheet)
        {
            if (sheet <= 0) throw new ArgumentOutOfRangeException(nameof(sheet));

            var pestana = _excel.GetSheetAt(sheet);

            var image = Image.FromFile(@"C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/Cemex.png");
            for (var i = 0; i < 2; i++)
            {
                var row = pestana.GetRow(i) ?? pestana.CreateRow(i);
                row.HeightInPoints = 51f;
                for (var j = 7; j < 13; j++)
                {
                    if (row.GetCell(j) == null)
                        row.CreateCell(j);

                }
            }

            var cra = new CellRangeAddress(0, 1, 8, 12);


            pestana.AddMergedRegion(cra);

            if (_cabeceraStyle == null)
                _cabeceraStyle = GetCabeceraCellStyle();
            var celda = pestana.GetRow(0).GetCell(8);
            celda.SetCellValue("NPOI");
            celda.CellStyle = _cabeceraStyle;


            ////row0.HeightInPoints = (float)image.Height;
            var converter = new ImageConverter();
            var data = (byte[])converter.ConvertTo(image, typeof(byte[]));
            //var pictureIndex = excel.AddPicture(data, PictureType.PNG);
            //var helper = excel.GetCreationHelper();
            //var drawing = pestana.CreateDrawingPatriarch();
            //var anchor = new XSSFClientAnchor(900, 0, 0, 0, 1, 1, 7, 2);


            //var picture = drawing.CreatePicture(anchor, pictureIndex);

            //picture.Resize(1);

            var myPictureId = _excel.AddPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG);

            var drawing = pestana.CreateDrawingPatriarch();
            var myAnchor = new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 7, 2);



            var myPicture = drawing.CreatePicture(myAnchor, myPictureId);
            myPicture.Resize();

        }

        /// <summary>
        /// 
        /// </summary>
        private void PutCabeceras()
        {

            if (_headerStyle == null)
                _headerStyle = GetHeaderCellStyle();


            for (var i = 0; i < _excel.NumberOfSheets; i++)
            {
                if (_excel.GetSheetAt(i).GetRow(_rowInicial - 1) == null) continue;
                foreach (var item in _excel.GetSheetAt(i).GetRow(_rowInicial - 1).Cells)
                {
                    item.CellStyle = _headerStyle;

                }

            }

        }



        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private ICellStyle GetHeaderCellStyle()
        {
            var style = _excel.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.DarkBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;

            var hfont = _excel.CreateFont();
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
            var style = _excel.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.DarkBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;

            style.FillBackgroundColor = IndexedColors.DarkBlue.Index;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            var hfont = _excel.CreateFont();
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
            ICellStyle style = _excel.CreateCellStyle();
            var hfont = _excel.CreateFont();
            hfont.FontHeightInPoints = 12;
            hfont.Color = IndexedColors.Black.Index;
            hfont.FontName = "Century Gothic";


            style.SetFont(hfont);
            return style;
        }

        /// <summary>
        /// 
        /// </summary>
        private void PutFitInCells()
        {


            for (int i = 0; i < _excel.NumberOfSheets; i++)
            {
                if (_excel.GetSheetAt(i).GetRow(_rowInicial) != null)
                {
                    int noOfColumns = _excel.GetSheetAt(i).GetRow(_rowInicial).LastCellNum;
                    for (int j = 0; j < noOfColumns; j++)
                    {
                        _excel.GetSheetAt(i).AutoSizeColumn(j);
                    }
                }


            }

        }

        /// <summary>
        /// 
        /// </summary>
        private void PutTypeAndSizeText()
        {
            if (_normalStyle == null)
                _normalStyle = GetNormalCellStyle();

            for (var i = 0; i < _excel.NumberOfSheets; i++)
            {
                if (_excel.GetSheetAt(i).GetRow(_rowInicial) == null) continue;
                for (var j = 0; j < _excel.GetSheetAt(i).GetRow(_rowInicial).Cells.Count; j++)
                {

                    _excel.GetSheetAt(i).SetDefaultColumnStyle(j, _normalStyle);
                }


            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        internal XSSFWorkbook GetExcelExample()
        {
            return _excel;
        }
    }
}