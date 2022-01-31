using Microsoft.AspNetCore.Mvc;

namespace File_Converter.Controllers
{
    public class ConvertFromPdfController : Controller
    {
        /// <summary>
        /// This method allows to choose the kind of conversion.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .pdf to .docx;
        /// </summary>
        /// <returns></returns>
        public IActionResult Pdf_To_Word()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .pdf to .pptx;
        /// </summary>
        /// <returns></returns>
        [HttpGet] 
        public IActionResult Pdf_To_PowerPoint()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .pdf to .xlsx;
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Pdf_To_Excel()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .pdf to .jpg;
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Pdf_To_Jpg()
        {
            return View();
        }
    }
}
