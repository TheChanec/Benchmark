using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;

namespace Library_benchmark.Helpers
{
    public class EPPLUSServicio
    {
        private ExcelPackage excel;
        private object cabeceras;
        private object informacion;
        private ExcelWorksheet currentsheet;

        public EPPLUSServicio(object cabeceras, object informacion)
        {
            this.cabeceras = cabeceras;
            this.informacion = informacion;
        }

        private void createWorkBook()
        {
            excel = new ExcelPackage();
            addSheet("Sheet1");
        }

        private void addSheet(string name)
        {
            currentsheet = excel.Workbook.Worksheets.Add(name);
        }

        private void addCol(int row, int col, string headerText, IEnumerable<string> info)
        {
            if (info.Select(x => x) != null)
            {
                currentsheet.Cells[row, col].Value = headerText;
                currentsheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                currentsheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(45, 0, 42, 89));
                currentsheet.Cells[row, col].Style.Font.Color.SetColor(Color.White);
                currentsheet.Cells[row, col].Style.Font.Bold = true;

                currentsheet.Cells[row, col].Value = headerText;
                try
                {
                    currentsheet.Cells[row + 1, col].LoadFromCollection(info, false);
                }
                catch (Exception)
                {

                    //throw;
                }
            }
            else {

            }
            
        }

        private void createExcelExample()
        {
        }

        enum Labels {
            Hola ="",
            Adios=""
        }
    }
}