using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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
        private HSSFWorkbook excel;
        private ISheet currentsheet;


        public NPOIService(IList<Dummy> informacion, int sheets)
        {
            this.informacion = informacion;
            createWorkBook();
            createSheets(sheets, true);

        }

        public NPOIService(byte[] excelFile, IList<Dummy> informacion, int sheets)
        {
            this.informacion = informacion;

            var fs = new MemoryStream(excelFile);
            excel = new HSSFWorkbook(fs);
            createSheets(sheets, true);
        }

        public NPOIService(int sheets)
        {
            createWorkBook();
            createSheets(sheets);

        }

        private void createSheets(int sheets, bool addInfo = false)
        {
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
                    if (addInfo)
                        addInformation();

                }
            }
        }

        private void ImagePortada()
        {


            Image image = Image.FromFile("C:/Users/mario.chan/Documents/GitHub/Library_benchmark/Library_benchmark/Content/images/net.png");
            var row0 = currentsheet.CreateRow(0);
            row0.CreateCell(0);

            //row0.HeightInPoints = (float)image.Height;
            var converter = new ImageConverter();
            var data = (byte[])converter.ConvertTo(image, typeof(byte[]));

            var pictureIndex = excel.AddPicture(data, PictureType.PNG);
            var helper = excel.GetCreationHelper();
            var drawing = currentsheet.CreateDrawingPatriarch();
            var anchor = helper.CreateClientAnchor();
            anchor.Col1 = 2; //0 index based column
            anchor.Row1 = 2; //0 index based row
            var picture = drawing.CreatePicture(anchor, pictureIndex);
            picture.Resize();
        }

        private void createWorkBook()
        {
            excel = new HSSFWorkbook();

        }

        private void addSheet(string name)
        {
            currentsheet = excel.GetSheet(name);
            if (currentsheet == null)
                currentsheet = excel.CreateSheet(name);
        }

        internal HSSFWorkbook GetExcelExample()
        {
            return excel;
        }

        private void addInformation()
        {


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