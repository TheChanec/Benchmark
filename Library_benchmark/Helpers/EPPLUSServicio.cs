using Library_benchmark.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;

namespace Library_benchmark.Helpers
{
    public class EPPLUSServicio
    {
        private ExcelPackage excel;
        private IList<Dummy> informacion;
        private ExcelWorksheet currentsheet;

        public EPPLUSServicio(IList<Dummy> informacion, int sheets)
        {
            this.informacion = informacion;

            createWorkBook();
            for (int i = 0; i <= sheets; i++)
            {

                if (i == 0)
                {
                    addSheet("Portada");
                    ImagePortada();
                }
                else
                {
                    addSheet("Sheet" + i);
                    addInformation();
                }
            }
            
        }

        public EPPLUSServicio(string path,IList<Dummy> informacion, int sheets)
        {
            this.informacion = informacion;

            excel = new ExcelPackage(new FileInfo(path));

            for (int i = 0; i <= sheets; i++)
            {

                if (i == 0)
                {
                    addSheet("Portada");
                    ImagePortada();
                }
                else
                {
                    addSheet("Sheet" + i);
                    addInformation();
                }
            }

        }

        public EPPLUSServicio()
        {
        }

        private void createWorkBook()
        {
            excel = new ExcelPackage();
        }

        private void addSheet(string name)
        {
            currentsheet = excel.Workbook.Worksheets.Add(name);
        }

        internal ExcelPackage GetExcelExample()
        {

            // FormatExcel();

            return excel;
        }

        private void ImagePortada()
        {
            Image logo = Image.FromFile("C:/Users/mario.chan/Documents/GitHub/Library_benchmark/Library_benchmark/Content/images/net.png");
            var picture = currentsheet.Drawings.AddPicture("32", logo);
            picture.SetPosition(2, 0, 2, 0);
        }

        private void FormatExcel()
        {

        }

        private void addInformation()
        {

            currentsheet.Cells[1, 1].LoadFromCollection(informacion, true);


        }

    }
}