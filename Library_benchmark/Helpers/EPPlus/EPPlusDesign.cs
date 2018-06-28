using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Library_benchmark.Helpers.EPPlus
{
    public class EpplusDesign
    {
        private readonly ExcelPackage _excel;
        private readonly bool _resource;
        private int _rowInicial;
        private Image _logo;
        private readonly Color _colorPrimary = Color.DarkBlue;
        private readonly Color _colorPrimaryText = Color.White;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="resource"></param>
        /// <param name="logo"></param>
        public EpplusDesign(ExcelPackage excel, bool resource, Image logo)
        {
            _excel = excel;
            _resource = resource;
            _logo = logo;
            _rowInicial = 4;

            DarFormato();
        }

        /// <summary>
        /// 
        /// </summary>
        private void DarFormato()
        {
            if (_excel == null) return;
            foreach (var item in _excel.Workbook.Worksheets)
            {
                PutCabeceras(item);
                PutTypeAndSizeText(item);
                PutFitInCells(item);
                if (!_resource)
                    PutImagenTitulo(item);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        private void PutImagenTitulo(ExcelWorksheet item)
        {
            item.Row(1).Height = 51;
            item.Row(2).Height = 51;

            item.Cells["I1:M2"].Merge = true;
            item.Cells["I1:M2"].Value = "EPPLUS";

            item.Cells["I1:M2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            item.Cells["I1:M2"].Style.Fill.BackgroundColor.SetColor(_colorPrimary);
            item.Cells["I1:M2"].Style.Font.Color.SetColor(_colorPrimaryText);
            item.Cells["I1:M2"].Style.Font.Size = 36f;
            item.Cells["I1:M2"].Style.Font.Bold = true;
            item.Cells["I1:M2"].Style.Font.Name = "Century Gothic";
            item.Cells["I1:M2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            item.Cells["I1:M2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            Image logo = Image.FromFile("C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/Cemex.png");
            var picture = item.Drawings.AddPicture("DotNet", logo);
            picture.SetPosition(13, 11);
            picture.SetSize(442, 116);


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        private void PutCabeceras(ExcelWorksheet workSheet)
        {

            var allCells = workSheet.Cells[_rowInicial, 1, _rowInicial, workSheet.Dimension.End.Column];
            allCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allCells.Style.Fill.BackgroundColor.SetColor(_colorPrimary /*Color.FromArgb(0, 42, 89)*/);
            allCells.Style.Font.Color.SetColor(_colorPrimaryText);
            allCells.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);




        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        private static void PutFitInCells(ExcelWorksheet workSheet)
        {
            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        private void PutTypeAndSizeText(ExcelWorksheet workSheet)
        {

            var allCells = workSheet.Cells[_rowInicial, 1, workSheet.Dimension.End.Row, workSheet.Dimension.End.Column];
            allCells.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Century Gothic", 12));



        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        internal ExcelPackage GetExcelExample()
        {
            return _excel;
        }
    }
}