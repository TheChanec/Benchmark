using System;
using NPOI.HSSF.UserModel;

namespace Library_benchmark.Helpers
{
    internal class NPOIDesign
    {
        private HSSFWorkbook excel;

        public NPOIDesign(HSSFWorkbook excel)
        {
            this.excel = excel;
        }

        internal HSSFWorkbook GetExcelExample()
        {
            throw new NotImplementedException();
        }
    }
}