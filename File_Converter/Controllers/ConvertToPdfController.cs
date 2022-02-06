using Aspose.Words;
using File_Converter.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.FileProviders;
using System.Net.Http.Headers;

namespace File_Converter.Controllers
{
    public class ConvertToPdfController : Controller
    {
        IWebHostEnvironment _appEnvironment;

        public ConvertToPdfController(IWebHostEnvironment appEnvironment)
        {
            _appEnvironment = appEnvironment;
        }

        /// <summary>
        /// This method with HttpGet attribute allows to choose the kind of conversion.
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .docx, .doc to .pdf;
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Word_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute convert files .docx, .doc to .pdf;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> Word_To_Pdf(IFormFile uploadedFile)
        {
            if (uploadedFile != null)
            {
                // Path to the folder Files
                string path = @"\files\" + uploadedFile.FileName;

                // Save files in Files in catalog wwwroot
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await uploadedFile.CopyToAsync(fileStream);
                }

                //Convert file from WORD to PDF
                string dirPath = _appEnvironment.WebRootPath + @"\files\";
                return File(await ConvertDOCXFile(dirPath), "application/pdf");
            }

            return View();
        }

        /// <summary>
        /// This method convert WORD files to PDF and return bytes array of current PDF file;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        public async Task<byte[]> ConvertDOCXFile(string dirPath)
        {
            //Load document
            Document document = new Document(dirPath + "test.docx");

            //Convert WORD to PDF
            document.Save(dirPath + "test.pdf");

            //Launch document
            string pdfPath = @"wwwroot\files\test.pdf";
            byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            return pdfBytes;
        }

        /// <summary>
        /// This method converts all files from .pptx, .ppt to .pdf;
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult PowerPoint_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .xlsx, .xls to .pdf;
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Excel_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .jpg to .pdf;
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Jpg_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method converts all files from .html to .pdf;
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public IActionResult Html_To_Pdf()
        {
            return View();
        }
    }
}
