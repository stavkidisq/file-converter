﻿using Microsoft.AspNetCore.Mvc;

namespace File_Converter.Controllers
{
    public class ConvertController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .docx, .doc to .pdf;
        /// </summary>
        /// <returns></returns>
        public IActionResult Word_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .pptx, .ppt to .pdf;
        /// </summary>
        /// <returns></returns>
        public IActionResult PowerPoint_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .xlsx, .xls to .pdf;
        /// </summary>
        /// <returns></returns>
        public IActionResult Excel_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .jpg to .pdf;
        /// </summary>
        /// <returns></returns>
        public IActionResult Jpg_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .html to .pdf;
        /// </summary>
        /// <returns></returns>
        public IActionResult Html_To_Pdf()
        {
            return View();
        }
    }
}
