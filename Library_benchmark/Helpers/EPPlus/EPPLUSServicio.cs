using Library_benchmark.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Library_benchmark.Helpers.EPPlus
{
    /// <summary>
    /// 
    /// </summary>
    public class EpplusServicio
    {
        #region Variables Locales
        private ExcelPackage _excel;
        private ExcelWorksheet _currentsheet;
        private ExcelWorksheet _basesheet;
        private readonly IList<ExcelDummy> _informacion;
        private readonly bool _mascaras;
        private int _inicialRow; 
        #endregion

        #region Constuctores

        /// <summary>
        /// 
        /// </summary>
        /// <param name="informacion"></param>
        /// <param name="design"></param>
        /// <param name="mascaras"></param>
        /// <param name="sheets"></param>
        public EpplusServicio(IList<ExcelDummy> informacion, bool design, bool mascaras, int sheets)
        {
            _informacion = informacion;
            _mascaras = mascaras;
            _inicialRow = design ? 4 : 1;

            CreateWorkBook();
            CreateSheets(sheets);

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="documentDummy"></param>
        /// <param name="informacion"></param>
        /// <param name="mascaras"></param>
        /// <param name="sheets"></param>
        public EpplusServicio(byte[] documentDummy, IList<ExcelDummy> informacion, bool mascaras, int sheets)
        {
            _informacion = informacion;
            _inicialRow = 4;
            _mascaras = mascaras;

            CreateWorkBook(documentDummy);
            CreateSheetBase();
            //deleteWorkSheets();
            CreateSheets(sheets);
        }

        #endregion


        /// <summary>
        /// 
        /// </summary>
        /// <param name="documentDummy"></param>
        private void CreateWorkBook(byte[] documentDummy)
        {
            using (var memStream = new MemoryStream(documentDummy))
                _excel = new ExcelPackage(memStream);

        }
        /// <summary>
        /// 
        /// </summary>
        private void CreateWorkBook()
        {
            _excel = new ExcelPackage();
        }
        /// <summary>
        /// 
        /// </summary>
        private void CreateSheetBase()
        {
            if (_excel.Workbook.Worksheets.Any())
                _basesheet = _excel.Workbook.Worksheets.FirstOrDefault();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheets"></param>
        private void CreateSheets(int sheets)
        {
            for (var i = 0; i < sheets; i++)
            {
                AddSheet("Sheet" + i);
                AddInformation();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
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

            if (_currentsheet != null) _currentsheet.DefaultRowHeight = 17.25;
        }
        /// <summary>
        /// 
        /// </summary>
        private void AddInformation()
        {

            _currentsheet.Cells[_inicialRow, 1].LoadFromCollection(_informacion, true, TableStyles.None);
            _currentsheet.Cells[_currentsheet.Dimension.Address].AutoFitColumns();
            if (_mascaras)
            {
                UseMascaras();
            }

        }
        /// <summary>
        /// 
        /// </summary>
        private void UseMascaras()
        {
            var propiedad = 0;
            foreach (var prop in typeof(ExcelDummy).GetProperties())
            {
                propiedad++;
                if (prop.PropertyType == typeof(DateTime))
                {
                    _currentsheet.Column(propiedad).Style.Numberformat.Format = "mm/dd/yyyy"; // hh:mm:ss AM/PM";
                }
                else if (prop.PropertyType == typeof(decimal))
                {
                    _currentsheet.Column(propiedad).Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        internal ExcelPackage GetExcelExample()
        {
            return _excel;
        }
    }
}