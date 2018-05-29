using Microsoft.VisualStudio.TestTools.UnitTesting;
using Library_benchmark.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace Library_benchmark.Controllers.Tests
{
    [TestClass()]
    public class HomeControllerTests
    {
        [TestMethod()]
        public void EPPLUSResultTest()
        {
            HomeController controller = new HomeController();
            var parametros = new Models.Parametros() { Cols =5 , Rows = 3  } ;
            //var result = controller.EPPLUSResult(parametros);

            Assert.Fail();
        }
    }
}