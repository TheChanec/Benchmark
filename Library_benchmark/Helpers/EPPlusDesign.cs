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
        private bool pintarCabeceras;

        public EPPlusDesign(ExcelPackage excel, bool pintarCabeceras)
        {
            this.excel = excel;
            this.pintarCabeceras = pintarCabeceras;

            DarFormato();
        }

        private void DarFormato()
        {
            if (excel != null)
            {
                foreach (var item in excel.Workbook.Worksheets)
                {
                    if (pintarCabeceras)
                    {
                        PutCabeceras(item);
                    }
                    PutTypeAndSizeText(item);
                    PutFitInCells(item);
                }

            }
        }

        internal void PutCabeceras(ExcelWorksheet workSheet)
        {
            try
            {
                var allCells = workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column];
                allCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                allCells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(45, 0, 42, 89));
                allCells.Style.Font.Color.SetColor(Color.White);
                allCells.Style.Font.Bold = true;
            }
            catch (Exception)
            {


            }


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
            try
            {
                var allCells = workSheet.Cells[1, 1, workSheet.Dimension.End.Row, workSheet.Dimension.End.Column];
                allCells.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                var cellFont = allCells.Style.Font;
                cellFont.SetFromFont(new Font("Arial", 12));
            }
            catch (Exception)
            {


            }

        }

        internal ExcelPackage GetExcelExample()
        {
            return excel;
        }
    }
}