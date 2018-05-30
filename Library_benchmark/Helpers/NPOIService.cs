using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Library_benchmark.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Library_benchmark.Helpers
{
    public class NPOIService
    {
        private IList<Dummy> informacion;
        private HSSFWorkbook excel;
        private ISheet currentsheet;

        public NPOIService(IList<Dummy> informacion, int sheets)
        {
            this.informacion = informacion;
            createWorkBook();

            for (int i = 0; i < sheets; i++)
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

        private void ImagePortada()
        {
            //throw new NotImplementedException();
        }

        private void createWorkBook()
        {
            excel = new HSSFWorkbook();

        }

        private void addSheet(string name)
        {
            currentsheet = excel.CreateSheet(name);
        }

        internal HSSFWorkbook GetExcelExample()
        {
            return excel;
        }

        private void addInformation()
        {
            Type type = typeof(Dummy);

            ICreationHelper cH = excel.GetCreationHelper();

            int cont = 0;
            foreach (var item in informacion)
            {
                IRow row = currentsheet.CreateRow(cont);
                int cell = 0;

                row.CreateCell(cell++).SetCellValue(item.Propiedad1.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad2.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad3.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad4.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad5.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad6.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad7.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad8.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad9.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad10.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad11.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad12.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad13.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad14.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad15.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad16.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad17.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad18.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad19.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad20.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad21.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad22.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad23.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad24.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad25.ToString());
                row.CreateCell(cell++).SetCellValue(item.Propiedad26.ToString());

                cont++;
            }
            
        }
    }
}