using Library_benchmark.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace Library_benchmark.Helpers
{
    public class EPPLUSServicio
    {
        private ExcelPackage excel;
        private IList<Dummy> informacion;
        private ExcelWorksheet currentsheet;
        private ExcelWorksheet basesheet;
        private int InicialRow;
        private bool design;

        public EPPLUSServicio(IList<Dummy> informacion, bool design, int sheets)
        {
            this.informacion = informacion;
            this.design = design;
            if (design)
                this.InicialRow = 8;
            else
                this.InicialRow = 1;

            createWorkBook();
            createSheets(sheets);

        }

        public EPPLUSServicio(string path, IList<Dummy> informacion, int sheets)
        {
            this.informacion = informacion;
            this.InicialRow = 8;
            this.design = false;


            createWorkBook(path);
            createSheetBase();
            createSheets(sheets);

        }



        private void createWorkBook()
        {
            excel = new ExcelPackage();
        }
        private void createWorkBook(string path)
        {
            excel = new ExcelPackage(new FileInfo(path));

        }
        private void createSheetBase()
        {
            basesheet = excel.Workbook.Worksheets.FirstOrDefault();
        }
        private void createSheets(int sheets)
        {
            for (int i = 0; i < sheets; i++)
            {
                addSheet("Sheet" + i);
                addInformation();
            }
        }

        private void addSheet(string name)
        {
            if (!excel.Workbook.Worksheets.Where(x => x.Name == name).Any())
            {
                if (basesheet != null)
                    currentsheet = excel.Workbook.Worksheets.Add(name, basesheet);
                else
                    currentsheet = excel.Workbook.Worksheets.Add(name);
            }
            else
                currentsheet = excel.Workbook.Worksheets.Where(x => x.Name == name).FirstOrDefault();

            currentsheet.DefaultRowHeight = 15;
        }





        private void FormatExcel()
        {

        }

        private void addInformation()
        {

            currentsheet.Cells[InicialRow, 1].LoadFromCollection(informacion, true, TableStyles.None);

            Mascaras();
        }

        private void Mascaras()
        {
            int propiedad = 0;
            foreach (PropertyInfo prop in typeof(Dummy).GetProperties())
            {
                propiedad++;
                if (prop.PropertyType.Equals(typeof(DateTime)))
                {
                    currentsheet.Column(propiedad).Style.Numberformat.Format = "mm/dd/yyyy"; // hh:mm:ss AM/PM";
                }
                else if (prop.PropertyType.Equals(typeof(Decimal)))
                {
                    currentsheet.Column(propiedad).Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
                }


            }


        }

        internal ExcelPackage GetExcelExample()
        {

            // FormatExcel();

            return excel;
        }
    }
}