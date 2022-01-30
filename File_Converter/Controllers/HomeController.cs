using Microsoft.AspNetCore.Mvc;

namespace File_Converter.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
