using System;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace Library_benchmark.Helpers
{
    internal class NPOIDesign
    {
        private HSSFWorkbook excel;
        private bool pintarCabeceras;

        HSSFCellStyle headerStyle;
        HSSFCellStyle normalStyle;


        private int sheets;

        public NPOIDesign(HSSFWorkbook excel, bool pintarCabeceras)
        {
            this.excel = excel;
            this.pintarCabeceras = pintarCabeceras;

            DarFormato();
        }












        private void DarFormato()
        {
            if (excel != null)
            {
                PutTypeAndSizeText();
                PutCabeceras();
                PutCeldasNormales();
                PutFitInCells();

            }
        }

        internal void PutCabeceras()
        {
            try
            {
                if (headerStyle == null)
                    headerStyle = GetHeaderCellStyle();


                for (int i = 0; i < excel.Workbook.NumSheets; i++)
                {
                    try
                    {
                        excel.GetSheetAt(i).GetRow(0).RowStyle = headerStyle;


                        foreach (var item in excel.GetSheetAt(i).GetRow(0).Cells)
                        {
                            item.CellStyle = headerStyle;

                        }
                    }
                    catch (Exception)
                    {

                        //throw;
                    }


                    //PutTypeAndSizeText(sheet);
                    //PutFitInCells((HSSFSheet)sheet);
                }



                //int noOfColumns = sheet.GetRow(0).LastCellNum;
                //for (int i = 0; i < noOfColumns; i++)
                //{
                //    sheet.SetDefaultColumnStyle(i, headerStyle);
                //}



            }
            catch (Exception)
            {
                //throw;

            }


        }

        internal void PutCeldasNormales()
        {

            if (normalStyle == null)
                normalStyle = GetNormalCellStyle();


            for (int i = 0; i < excel.Workbook.NumSheets; i++)
            {

                //excel.GetSheetAt(i).GetRow(0).RowStyle = headerStyle;

                for (int j = 1; j <= excel.GetSheetAt(i).LastRowNum; j++)
                {
                    foreach (var item in excel.GetSheetAt(i).GetRow(j).Cells)
                    {
                        item.CellStyle = normalStyle;

                    }
                }

                
            }


            //PutTypeAndSizeText(sheet);
            //PutFitInCells((HSSFSheet)sheet);




            //int noOfColumns = sheet.GetRow(0).LastCellNum;
            //for (int i = 0; i < noOfColumns; i++)
            //{
            //    sheet.SetDefaultColumnStyle(i, headerStyle);
            //}






        }

        private HSSFCellStyle GetHeaderCellStyle()
        {
            HSSFCellStyle style = (HSSFCellStyle)excel.CreateCellStyle();
            style.FillForegroundColor = HSSFColor.Yellow.Index; ;
            style.FillPattern = FillPattern.SolidForeground;
            style.FillBackgroundColor = HSSFColor.Yellow.Index;


            var hfont = (HSSFFont)excel.CreateFont();
            hfont.FontHeightInPoints = 13;
            hfont.IsBold = true;
            hfont.Color = IndexedColors.White.Index;
            style.SetFont(hfont);

            return style;
        }

        private HSSFCellStyle GetNormalCellStyle()
        {
            var style = (HSSFCellStyle)excel.CreateCellStyle();
            var hfont = (HSSFFont)excel.CreateFont();
            hfont.FontHeightInPoints = 13;
            hfont.Color = IndexedColors.Black.Index;
            hfont.IsBold = true;
            style.SetFont(hfont);
            return style;
        }

        internal void PutFitInCells()
        {


            for (int i = 0; i < excel.Workbook.NumSheets; i++)
            {
                try
                {
                    int noOfColumns = excel.GetSheetAt(i).GetRow(0).LastCellNum;
                    for (int j = 0; j < noOfColumns; j++)
                    {
                        excel.GetSheetAt(i).AutoSizeColumn(j);
                    }
                }
                catch (Exception)
                {

                    throw;
                }

            }






        }
        internal void PutTypeAndSizeText()
        {
            try
            {
                if (headerStyle == null)
                    headerStyle = GetHeaderCellStyle();


                for (int i = 0; i < excel.Workbook.NumSheets; i++)
                {

                    excel.GetSheetAt(i).GetRow(0).RowStyle = headerStyle;

                    for (int j = 0; j < excel.GetSheetAt(i).GetRow(0).Cells.Count; j++)
                    {
                        excel.GetSheetAt(i).SetDefaultColumnStyle(j, headerStyle);
                    }




                    //PutTypeAndSizeText(sheet);
                    //PutFitInCells((HSSFSheet)sheet);
                }

            }
            catch (Exception)
            {
                //throw;

            }






        }



        internal HSSFWorkbook GetExcelExample()
        {
            return excel;
        }
    }
}