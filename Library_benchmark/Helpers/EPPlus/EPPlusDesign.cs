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
        private Image logo;
        private Color colorPrimary = Color.DarkBlue;
        private Color colorPrimaryText= Color.White;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="resource"></param>
        /// <param name="logo"></param>
        public EPPlusDesign(ExcelPackage excel, bool resource, Image logo)
        {
            this.excel = excel;
            this.resource = resource;
            this.logo = logo;
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
                foreach (var item in excel.Workbook.Worksheets)
                {
                    

                    PutCabeceras(item);
                    PutTypeAndSizeText(item);
                    PutFitInCells(item);

                    if (!resource)
                        PutImagenTitulo(item);
                }

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
            item.Cells["I1:M2"].Style.Fill.BackgroundColor.SetColor(colorPrimary);
            item.Cells["I1:M2"].Style.Font.Color.SetColor(colorPrimaryText);
            item.Cells["I1:M2"].Style.Font.Size = 36f;
            item.Cells["I1:M2"].Style.Font.Name = "Century Gothic";
            item.Cells["I1:M2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


            Image logo = Image.FromFile("C:/Users/mario.chan/Documents/GitHub/Benchmark/Library_benchmark/Content/images/Cemex.png");
            var picture = item.Drawings.AddPicture("DotNet", logo);
            picture.SetPosition(13, 11);
            picture.SetSize(442,116);
            
            
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        private void PutCabeceras(ExcelWorksheet workSheet)
        {

            var allCells = workSheet.Cells[rowInicial, 1, rowInicial, workSheet.Dimension.End.Column];
            allCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allCells.Style.Fill.BackgroundColor.SetColor(colorPrimary /*Color.FromArgb(0, 42, 89)*/);
            allCells.Style.Font.Color.SetColor(colorPrimaryText);
            allCells.Style.Border.BorderAround( ExcelBorderStyle.Thin ,Color.Black);
            



        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        private void PutFitInCells(ExcelWorksheet workSheet)
        {


            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workSheet"></param>
        private void PutTypeAndSizeText(ExcelWorksheet workSheet)
        {

            var allCells = workSheet.Cells[rowInicial, 1, workSheet.Dimension.End.Row, workSheet.Dimension.End.Column];
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

            return excel;
        }
    }
}