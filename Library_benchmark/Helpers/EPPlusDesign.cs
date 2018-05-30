using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeOpenXml;

namespace Library_benchmark.Helpers
{
    public class EPPlusDesign
    {
        private ExcelPackage excel;
        private int rowInicial;
        private bool pintarCabeceras;
        private string typo;
        private int textSize;

        public EPPlusDesign(ExcelPackage excel, int rowInicial, bool pintarCabeceras, string typo, int textSize)
        {
            this.excel = excel;
            this.rowInicial = rowInicial;
            this.pintarCabeceras = pintarCabeceras;
            this.typo = typo;
            this.textSize = textSize;

            if (pintarCabeceras) {
                PutCabeceras();
            }

            PutTypeText();
            PutFitInCells();
            PutSizeText();

        }

        internal void PutCabeceras()
        {
            throw new NotImplementedException();
        }

        internal void PutFitInCells()
        {
            throw new NotImplementedException();
        }

        internal object PutTypeText()
        {
            throw new NotImplementedException();
        }

        internal void PutSizeText()
        {
            throw new NotImplementedException();
        }

        internal ExcelPackage GetExcelExample()
        {
            throw new NotImplementedException();
        }
    }
}