using Library_benchmark.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Library_benchmark.Helpers.NPOI
{
    public class NpoiService
    {
        private readonly IList<ExcelDummy> _informacion;
        private XSSFWorkbook _excel;
        private XSSFSheet _currentsheet;
        private XSSFSheet _basesheet;
        private readonly int _rowInicial;
        private readonly bool _mascaras;


        /// <summary>
        /// Contructor base para NPOI
        /// </summary>
        /// <param name="informacion">Lista de registros que se incrustaran en las hojas</param>
        /// <param name="design">parametro bandera para definir si se pondra diseño a las hojas</param>
        /// <param name="mascaras">Parametro para anunciar si permite mascaras en el excel</param>
        /// <param name="sheets">numero de hojas que tendra el workbook</param>
        public NpoiService(IList<ExcelDummy> informacion, bool design, bool mascaras, int sheets)
        {
            _mascaras = mascaras;
            _informacion = informacion;

            _rowInicial = design ? 4 : 1;

            CreateWorkBook();
            CreateSheets(sheets);
        }

        /// <summary>
        /// Contructor para armar el NPOI en base a un template ya establecido
        /// </summary>
        /// <param name="excelFile">archivo template en arreglo de bytes</param>
        /// <param name="informacion">informacion que se incrustara en las hojas</param>
        /// <param name="mascaras">Parametro para anunciar si permite mascaras en el excel</param>
        /// <param name="sheets">numero de hojas que tendra el workbook</param>
        public NpoiService(byte[] excelFile, IList<ExcelDummy> informacion, bool mascaras, int sheets)
        {
            _informacion = informacion;
            _rowInicial = 4;
            _mascaras = mascaras;

            CreateWorkBook(excelFile);
            CreateSheetBase();
            CreateSheets(sheets);
        }


        /// <summary>
        /// obtiene la primera hoja de el template para utilizarla como hoja base para armar el workbook
        /// </summary>
        private void CreateSheetBase()
        {
            _basesheet = (XSSFSheet)_excel.GetSheetAt(0);
        }

        /// <summary>
        /// crea un workbook base 
        /// </summary>
        private void CreateWorkBook()
        {
            _excel = new XSSFWorkbook();

        }

        /// <summary>
        /// crea un workbook en base al templete establecido 
        /// </summary>
        /// <param name="excelFile">Templete</param>
        private void CreateWorkBook(byte[] excelFile)
        {

            var fs = new MemoryStream(excelFile);
            //excel = new HSSFWorkbook(fs);
            _excel = (XSSFWorkbook)WorkbookFactory.Create(fs);
        }

        /// <summary>
        /// Funcion que crea hojas en el workbook
        /// </summary>
        /// <param name="sheets">numero de hojas que se crearan</param>
        private void CreateSheets(int sheets)
        {
            for (var i = 0; i < sheets; i++)
            {
                AddSheet("Sheet" + i);
                Addcabeceras();

                AddInformation();

                PutFitInCells();
            }
        }

        /// <summary>
        /// Pone un autoFit en las columnas
        /// </summary>
        private void PutFitInCells()
        {
            int noOfColumns = _currentsheet.GetRow(_rowInicial - 1).LastCellNum;
            for (var j = 0; j < noOfColumns; j++)
            {
                _currentsheet.AutoSizeColumn(j, false);
            }
        }

        /// <summary>
        /// Se encarga de poner el titulo de las tablas y de definil el estilo que tendra cada columna por defecto
        /// </summary>
        private void Addcabeceras()
        {
            var row = _currentsheet.GetRow(_rowInicial - 1) ?? _currentsheet.CreateRow(_rowInicial - 1);
            var cell = 0;
            var item = _informacion.FirstOrDefault();

            if (item == null) return;
            foreach (var prop in item.GetType().GetProperties().Where(p => !p.GetGetMethod().GetParameters().Any()))
            {
                var celda = row.GetCell(cell) ?? row.CreateCell(cell);

                if (_mascaras)
                {
                    var hfont = _excel.CreateFont();
                    hfont.FontHeightInPoints = 12;
                    hfont.Color = IndexedColors.Black.Index;
                    hfont.FontName = "Century Gothic";

                    if (prop.PropertyType == typeof(DateTime))
                    {
                        var style = _excel.CreateCellStyle();
                        style.DataFormat = _excel.CreateDataFormat().GetFormat("m/d/yyyy");

                        style.SetFont(hfont);
                        _currentsheet.SetDefaultColumnStyle(cell, style);
                    }
                    else if (prop.PropertyType == typeof(decimal))
                    {
                        var style = _excel.CreateCellStyle();
                        style.DataFormat = _excel.CreateDataFormat().GetFormat("[$$-409]#,##0.00");

                        style.SetFont(hfont);
                        _currentsheet.SetDefaultColumnStyle(cell, style);
                        celda.SetCellType(CellType.Numeric);
                    }
                    else
                    {
                        var style = _excel.CreateCellStyle();

                        style.SetFont(hfont);
                        _currentsheet.SetDefaultColumnStyle(cell, style);
                        celda.SetCellType(CellType.Numeric);
                    }
                }

                cell++;
                celda.SetCellValue(prop.Name);
            }
        }

        /// <summary>
        /// Agrega Sheet a el excel en base a nombre
        /// </summary>
        /// <param name="name"></param>
        private void AddSheet(string name)
        {
            _currentsheet = (XSSFSheet)_excel.GetSheet(name);
            if (_currentsheet == null)
            {
                if (_basesheet != null)
                    _currentsheet = (XSSFSheet)_basesheet.CopySheet(name, true);
                else
                    _currentsheet = (XSSFSheet)_excel.CreateSheet(name);

            }
            _currentsheet.DefaultRowHeight = 300;
        }

        /// <summary>
        /// Agrega informacion a la Sheet que este en memoria 
        /// </summary>
        private void AddInformation()
        {
            var cont = _rowInicial;
            foreach (var item in _informacion)
            {
                var row = _currentsheet.GetRow(cont) ?? _currentsheet.CreateRow(cont);
                var cell = 0;

                foreach (var prop in item.GetType().GetProperties().Where(p => !p.GetGetMethod().GetParameters().Any()))
                {
                    var celda = row.GetCell(cell) ?? row.CreateCell(cell);

                    var style = _currentsheet.GetColumnStyle(cell);
                    if (prop.PropertyType == typeof(DateTime))
                    {
                        var date = (DateTime)prop.GetValue(item, null);
                        celda.SetCellValue(date.Date);
                        style.DataFormat = _excel.CreateDataFormat().GetFormat("MM/dd/yyyy");
                    }
                    else if (prop.PropertyType == typeof(decimal))
                    {
                        var money = (decimal)prop.GetValue(item, null);
                        celda.SetCellValue(Convert.ToDouble(money));
                        style.DataFormat = _excel.CreateDataFormat().GetFormat("[$$-409]#,##0.00");

                        celda.SetCellType(CellType.Numeric);

                    }
                    else
                        celda.SetCellValue(prop.GetValue(item, null).ToString());


                    celda.CellStyle = style;
                    cell++;
                }
                cont++;
            }

        }

        /// <summary>
        /// Obtiene el dato del Workbook 
        /// </summary>
        /// <returns>dato referente a el excel que se esta armando</returns>
        internal XSSFWorkbook GetExcelExample()
        {
            return _excel;
        }

    }
}