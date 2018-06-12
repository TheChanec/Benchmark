using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Library_benchmark.Helpers
{
    public class EPPlusDesign
    {
        private ExcelPackage excel;
        private bool resource;
        private int rowInicial;

        public EPPlusDesign(ExcelPackage excel, bool resource)
        {
            this.excel = excel;
            this.resource = resource;
            rowInicial = 8;

            DarFormato();
        }

        private void DarFormato()
        {
            if (excel != null)
            {
                foreach (var item in excel.Workbook.Worksheets)
                {
                    if (!resource)
                        PutImagenTitulo(item);

                    PutCabeceras(item);
                    PutTypeAndSizeText(item);
                    PutFitInCells(item);
                }

            }
        }

        private void PutImagenTitulo(ExcelWorksheet item)
        {
            Image logo = Image.FromFile("C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/net.png");
            var picture = item.Drawings.AddPicture("DotNet", logo);
            picture.SetPosition(0, 5, 0, 5);
            picture.SetSize(110, 110);

            item.Cells["C1:O6"].Merge = true;
            item.Cells["C1:O6"].Value = "EPPLUS";

            item.Cells["C1:O6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            item.Cells["C1:O6"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(45, 0, 42, 89));
            item.Cells["C1:O6"].Style.Font.Color.SetColor(Color.White);
            item.Cells["C1:O6"].Style.Font.Size = 72f;
            item.Cells["C1:O6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }


        internal void PutCabeceras(ExcelWorksheet workSheet)
        {

            var allCells = workSheet.Cells[rowInicial, 1, rowInicial, workSheet.Dimension.End.Column];
            allCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allCells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(45, 0, 42, 89));
            allCells.Style.Font.Color.SetColor(Color.White);
            allCells.Style.Font.Bold = true;




        }
        internal void PutFitInCells(ExcelWorksheet workSheet)
        {
            try
            {

                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
            }
            catch (Exception)
            {


            }

        }
        internal void PutTypeAndSizeText(ExcelWorksheet workSheet)
        {

            var allCells = workSheet.Cells[rowInicial, 1, workSheet.Dimension.End.Row, workSheet.Dimension.End.Column];
            allCells.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            var cellFont = allCells.Style.Font;
            cellFont.SetFromFont(new Font("Arial", 12));



        }

        internal ExcelPackage GetExcelExample()
        {
            
            return excel;
        }
    }
}