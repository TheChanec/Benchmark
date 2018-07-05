using System.Web.Mvc;

namespace Library_benchmark.Controllers
{
    /// <inheritdoc />
    /// <summary>
    /// Controller principal
    /// </summary>
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }


    }
}