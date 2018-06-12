using Microsoft.VisualStudio.TestTools.UnitTesting;
using Library_benchmark.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Library_benchmark.Controllers.Tests
{
    [TestClass()]
    public class HomeControllerTests
    {
        [TestMethod()]
        public void EPPLUSResultTest()
        {
            ExcelController controller = new ExcelController();
            var parametros = new Models.ParametrosExcel() { Cols =5 , Rows = 3  } ;
            //var result = controller.EPPLUS(parametros);

            Assert.Fail();
        }
    }
}