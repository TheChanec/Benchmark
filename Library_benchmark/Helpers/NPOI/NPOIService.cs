using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using Library_benchmark.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Library_benchmark.Helpers
{
    public class NPOIService
    {
        private IList<Dummy> informacion;
        private bool design;
        private HSSFWorkbook excel;
        private ISheet currentsheet;
        private int rowInicial;
        private ISheet basesheet;

        public NPOIService(IList<Dummy> informacion, bool design, int sheets)
        {
            this.informacion = informacion;
            this.design = design;
            if (design)
                this.rowInicial = 8;
            else
                this.rowInicial = 1;

            createWorkBook();
            createSheets(sheets);

        }

        public NPOIService(byte[] excelFile, IList<Dummy> informacion, int sheets)
        {
            this.informacion = informacion;
            this.design = false;
            this.rowInicial = 8;

            createWorkBook(excelFile);
            createSheetBase();
            createSheets(sheets);
        }

        private void createSheetBase()
        {
            basesheet = excel.GetSheetAt(0);
        }

        private void createWorkBook()
        {
            excel = new HSSFWorkbook();

        }
        private void createWorkBook(byte[] excelFile)
        {
            var fs = new MemoryStream(excelFile);
            excel = new HSSFWorkbook(fs);
        }


        private void createSheets(int sheets)
        {
            for (int i = 0; i < sheets; i++)
            {
                addSheet("Sheet" + i);
                addcabeceras();
                addInformation();
            }
        }

        private void addcabeceras()
        {
            IRow row = currentsheet.CreateRow(rowInicial - 1);
            int cell = 0;

            var item = informacion.FirstOrDefault();
            foreach (var prop in item.GetType().GetProperties().Where(p => !p.GetGetMethod().GetParameters().Any()))
            {
                var celda = row.CreateCell(cell++);
                celda.SetCellValue(prop.Name.ToString());
            }
        }

        private void addSheet(string name)
        {
            currentsheet = excel.GetSheet(name);
            if (currentsheet == null)
            {
                if (basesheet != null)
                    currentsheet = basesheet.CopySheet(name, true);
                else
                    currentsheet = excel.CreateSheet(name);

            }
            currentsheet.DefaultRowHeight = 300;
        }

        private void addInformation()
        {
            int cont = rowInicial;
            foreach (var item in informacion)
            {
                IRow row = currentsheet.CreateRow(cont);
                int cell = 0;

                foreach (var prop in item.GetType().GetProperties().Where(p => !p.GetGetMethod().GetParameters().Any()))
                {
                    var celda = row.CreateCell(cell++);

                    if (prop.PropertyType.Equals(typeof(DateTime)))
                    {
                        var date = (DateTime)prop.GetValue(item, null);
                        celda.SetCellValue(date.ToString("MM/dd/yyyy"));

                    }
                    else if (prop.PropertyType.Equals(typeof(Decimal)))
                    {
                        var money = (Decimal)prop.GetValue(item, null);
                        celda.SetCellValue(money.ToString("C"));

                    }
                    else
                        celda.SetCellValue(prop.GetValue(item, null).ToString());
                }
                cont++;
            }

        }

        internal HSSFWorkbook GetExcelExample()
        {
            return excel;
        }
    }
}