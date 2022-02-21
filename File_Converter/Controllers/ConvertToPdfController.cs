using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using Aspose.Imaging;
using File_Converter.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.FileProviders;
using System.Net.Http.Headers;
using Aspose.Pdf;
using System.Net;
using File_Converter.Models.ValidationAttributes;
using System.ComponentModel.DataAnnotations;
using File_Converter.Models.BusinessModels;

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
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpGet attribute converts all files from .docx, .doc to .pdf;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Word_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute convert files .docx, .doc to .pdf;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns>If everything is ok, then current pdf file</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Word_To_Pdf(UploadWordModel uploadedFile)
        {
            if(ModelState.IsValid)
            {
                // Path to the folder Files
                string path = @"\files\" + uploadedFile.Information.FileName;

                // Save files in Files in catalog wwwroot
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await uploadedFile.Information.CopyToAsync(fileStream);
                }

                //Convert file from WORD to PDF
                string dirPath = @"wwwroot\files\";
                return File
                    (await ConvertDOCXFile(dirPath, uploadedFile.Name, uploadedFile.Type), "application/pdf", fileDownloadName: uploadedFile.Name + ".pdf");
            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert WORD files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns>Array consisting from bytes of current PDF file</returns>
        public async Task<byte[]> ConvertDOCXFile(string dirPath, string fileName, string fileType)
        {
            //Load document
            Aspose.Words.Document document = new Aspose.Words.Document(dirPath + fileName + fileType);

            //Convert WORD to PDF
            document.Save(dirPath + fileName + ".pdf");

            //Open document in browser
            string pdfPath = dirPath + fileName + ".pdf";
            byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            return pdfBytes;
        }

        /// <summary>
        /// This method with HttpGet attribute converts all files from .pptx, .ppt to .pdf;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult PowerPoint_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .pptx, .ppt to .pdf;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns>If everything is ok, then current pdf file</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> PowerPoint_To_Pdf(UploadPowerPointModel uploadedFile)
        {
            if(ModelState.IsValid)
            {
                // Path to the folder Files;
                string path = @"\files\" + uploadedFile.Information.FileName;

                // Save files in Files in catalog wwwroot;
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await uploadedFile.Information.CopyToAsync(fileStream);
                }

                //Convert file from POWERPOINT to PDF;
                string dirPath = @"wwwroot\files\";

                // TODO: may be allows uploadedFile.Name
                return File
                    (await ConvertPPTXFile(dirPath, uploadedFile.Name, uploadedFile.Type), "application/pdf", fileDownloadName: uploadedFile.Name + ".pdf");
            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert PowerPoint files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns>Array consisting from bytes of current PDF file</returns>
        public async Task<byte[]> ConvertPPTXFile(string dirPath, string fileName, string fileType)
        {
            // Instantiate a Presentation object that represents a PPTX file
            Presentation presentation = new Presentation(dirPath + fileName + fileType);

            // Save the presentation as PDF
            presentation.Save(dirPath + fileName + ".pdf", Aspose.Slides.Export.SaveFormat.Pdf);

            //Launch document
            string pdfPath = dirPath + fileName + ".pdf";
            byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            return pdfBytes;
        }

        /// <summary>
        /// This method with HttpGet attribute converts all files from .xlsx, .xls to .pdf;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Excel_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .xlsx, .xls to .pdf;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns>If everything is ok, then current pdf file</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Excel_To_Pdf(UploadExcelModel uploadedFile)
        {
            if(ModelState.IsValid)
            {
                // Path to the folder Files;
                string path = @"\files\" + uploadedFile.Information.FileName;

                // Save files in Files in catalog wwwroot;
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await uploadedFile.Information.CopyToAsync(fileStream);
                }

                //Convert file from EXCEL to PDF;
                string dirPath = @"wwwroot\files\";

                // TODO: may be allows uploadedFile.Name
                return File
                    (await ConvertEXCELFile(dirPath, uploadedFile.Name, uploadedFile.Type), "application/pdf",fileDownloadName: uploadedFile.Name + ".pdf");

            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert EXCEL files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns>Array consisting from bytes of current PDF file</returns>
        public async Task<byte[]> ConvertEXCELFile(string dirPath, string fileName, string fileType)
        {
            // Instantiate a Presentation object that represents a PPT file
            Workbook presentation = new Workbook(dirPath + fileName + fileType);

            // Save the presentation as PDF
            presentation.Save(dirPath + fileName + ".pdf");

            //Launch document
            string pdfPath = dirPath + fileName + ".pdf";
            byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            return pdfBytes;
        }

        /// <summary>
        /// This method converts all files from .jpg to .pdf;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Jpg_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .jpg to .pdf;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns>If everything is ok, then current pdf file</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Jpg_To_Pdf(UploadJpgModel uploadedFile)
        {
            if(ModelState.IsValid)
            {
                // Path to the folder Files;
                string path = @"\files\" + uploadedFile.Information.FileName;

                // Save files in Files in catalog wwwroot;
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await uploadedFile.Information.CopyToAsync(fileStream);
                }

                //Convert file from EXCEL to PDF;
                string dirPath = @"wwwroot\files\";

                // TODO: may be allows uploadedFile.Name
                return File
                    (await ConvertJPGFile(dirPath, uploadedFile.Name, uploadedFile.Type), "application/pdf", fileDownloadName: uploadedFile.Name + ".pdf");

            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert JPG files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns>Array consisting from bytes of current PDF file</returns>
        public async Task<byte[]> ConvertJPGFile(string dirPath, string fileName, string fileType)
        {
            // Initialize new PDF document
            Aspose.Pdf.Document doc = new Aspose.Pdf.Document();

            // Add empty page in empty document
            Page page = doc.Pages.Add();
            Aspose.Pdf.Image image = new Aspose.Pdf.Image();
            image.File = (dirPath + fileName + fileType);

            // Add image on a page
            page.Paragraphs.Add(image);

            // Save output PDF file
            doc.Save(dirPath + fileName + ".pdf");

            //Launch document
            string pdfPath = dirPath + fileName + ".pdf";
            byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            return pdfBytes;
        }

        /// <summary>
        /// This method with HttpGet attribute converts all files from .html to .pdf;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Html_To_Pdf()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .html to .pdf;
        /// </summary>
        /// <param name="url">Contains a link to the page which should be converted</param>
        /// <returns>If everything is ok, then current pdf file</returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Html_To_Pdf(string url)
        {
            // Set page size A3 and Landscape orientation;
            Aspose.Pdf.HtmlLoadOptions options = new Aspose.Pdf.HtmlLoadOptions(url)
            {
                PageInfo = { Width = 842, Height = 1191, IsLandscape = true }
            };

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(await GetContentFromUrlAsStream(url), options);
            pdfDocument.Save(@"wwwroot\files\html_test.PDF");

            return View();
        }

        /// <summary>
        /// This method get content from URL as stream;
        /// </summary>
        /// <param name="url">Contains a link to the page which should be converted</param>
        /// <param name="credentials">Contains credentials for Web client authentication</param>
        /// <returns>Stream URL of page</returns>
        private async Task<Stream> GetContentFromUrlAsStream(string url, ICredentials credentials = default)
        {
            using (var handler = new HttpClientHandler { Credentials = credentials })
            using (var httpClient = new HttpClient(handler))
            {
                return await httpClient.GetStreamAsync(url);
            }
        }
    }
}
