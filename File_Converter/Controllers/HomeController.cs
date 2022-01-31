using Microsoft.AspNetCore.Mvc;

namespace File_Converter.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
    }
}
