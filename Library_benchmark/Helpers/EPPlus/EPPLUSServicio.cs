using Library_benchmark.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Library_benchmark.Helpers.EPPlus
{
    public class EpplusServicio
    {
        private ExcelPackage _excel;
        private ExcelWorksheet _currentsheet;
        private ExcelWorksheet _basesheet;
        private readonly IList<Dummy> _informacion;
        private readonly bool _design;
        private readonly bool _mascaras;
        private int _inicialRow;

        

        public EpplusServicio(IList<Dummy> informacion, bool design, bool mascaras, int sheets)
        {
            _informacion = informacion;
            _design = design;
            _mascaras = mascaras;
            _inicialRow = design ? 4 : 1;

            CreateWorkBook();
            CreateSheets(sheets);

        }

        public EpplusServicio(byte[] documentDummy, IList<Dummy> informacion, bool mascaras, int sheets)
        {
            _informacion = informacion;
            _inicialRow = 4;
            _design = false;
            _mascaras = mascaras;

            CreateWorkBook(documentDummy);
            CreateSheetBase();
            //deleteWorkSheets();
            CreateSheets(sheets);
        }

        private void CreateWorkBook(byte[] documentDummy)
        {
            using (var memStream = new MemoryStream(documentDummy))
            {
                _excel = new ExcelPackage(memStream);
            }
        }
        private void CreateWorkBook()
        {
            _excel = new ExcelPackage();
        }
        private void CreateWorkBook(string path)
        {
            _excel = new ExcelPackage(new FileInfo(path));

        }
        private void CreateSheetBase()
        {
            if (_excel.Workbook.Worksheets.Any())
                _basesheet = _excel.Workbook.Worksheets.FirstOrDefault();
        }
        private void CreateSheets(int sheets)
        {
            for (var i = 0; i < sheets; i++)
            {
                AddSheet("Sheet" + i);
                AddInformation();
            }
        }
        private void AddSheet(string name)
        {
            if (_excel.Workbook.Worksheets.All(x => x.Name != name))
            {
                _currentsheet = _basesheet != null ? _excel.Workbook.Worksheets.Add(name, _basesheet) : _excel.Workbook.Worksheets.Add(name);
            }
            else
            {
                _currentsheet = _excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == name);
            }


            _currentsheet.DefaultRowHeight = 17.25;
        }
        private void AddInformation()
        {

            _currentsheet.Cells[_inicialRow, 1].LoadFromCollection(_informacion, true, TableStyles.None);
            _currentsheet.Cells[_currentsheet.Dimension.Address].AutoFitColumns();
            if (_mascaras)
            {
                Mascaras();
            }

        }
        private void Mascaras()
        {
            var propiedad = 0;
            foreach (var prop in typeof(Dummy).GetProperties())
            {
                propiedad++;
                if (prop.PropertyType.Equals(typeof(DateTime)))
                {
                    _currentsheet.Column(propiedad).Style.Numberformat.Format = "mm/dd/yyyy"; // hh:mm:ss AM/PM";
                }
                else if (prop.PropertyType.Equals(typeof(Decimal)))
                {
                    _currentsheet.Column(propiedad).Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
                }
            }
        }
        internal ExcelPackage GetExcelExample()
        {
            return _excel;
        }
    }
}