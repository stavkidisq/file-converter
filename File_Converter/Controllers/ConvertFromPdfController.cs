using File_Converter.Models.BusinessModels;
using File_Converter.Models.ValidationAttributes;
using Microsoft.AspNetCore.Mvc;
using System.IO.Compression;

namespace File_Converter.Controllers
{
    public class ConvertFromPdfController : Controller
    {
        IWebHostEnvironment _appEnvironment;

        public ConvertFromPdfController(IWebHostEnvironment appEnvironment)
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
        /// This method with HttpGet attribute converts all files from .pdf to .docx;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Pdf_To_Word()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .pdf to .jpg;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Pdf_To_Word(UploadPdfModel uploadedFile)
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
                    (await ConvertToDOCXFile(dirPath, uploadedFile.Information.FileName), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert WORD files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToDOCXFile(string dirPath, string fileName)
        {
            // Open the source PDF document
            Aspose.Pdf.Document document = new Aspose.Pdf.Document(dirPath + fileName);

            // Save the file into MS document format
            document.Save(dirPath + "convertedfile.docx", Aspose.Pdf.SaveFormat.DocX);

            //Open document in browser
            string docPath = @"wwwroot\files\convertedfile.docx";
            byte[] docBytes = await System.IO.File.ReadAllBytesAsync(docPath);
            await System.IO.File.WriteAllBytesAsync(docPath, docBytes);

            return docBytes;
        }

        /// <summary>
        /// This method with HttpGet attribute converts all files from .pdf to .pptx;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet] 
        public IActionResult Pdf_To_PowerPoint()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .pdf to .jpg;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Pdf_To_PowerPoint(UploadPdfModel uploadedFile)
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
                    (await ConvertToPPTXFile(dirPath, uploadedFile.Information.FileName), "application/vnd.openxmlformats-officedocument.presentationml.presentation");

            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert PowerPoint files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToPPTXFile(string dirPath, string fileName)
        {
            // Load PDF document
            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(dirPath + fileName);
            // Instantiate PptxSaveOptions instance
            Aspose.Pdf.PptxSaveOptions pptx_save = new Aspose.Pdf.PptxSaveOptions();
            // Save the output in PPTX format
            doc.Save(dirPath + "convertedfile.pptx", pptx_save);

            //Open document in browser
            string docPath = @"wwwroot\files\convertedfile.pptx";
            byte[] docBytes = await System.IO.File.ReadAllBytesAsync(docPath);
            await System.IO.File.WriteAllBytesAsync(docPath, docBytes);

            return docBytes;
        }
        /// <summary>
        /// This method with HttpGet attribute converts all files from .pdf to .xlsx;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Pdf_To_Excel()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .pdf to .jpg;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Pdf_To_Excel(UploadPdfModel uploadedFile)
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
                    (await ConvertToXLSXFile(dirPath, uploadedFile.Information.FileName), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert Excel files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToXLSXFile(string dirPath, string fileName)
        {
            // Load PDF document
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(dirPath + fileName);

            // Instantiate ExcelSave Option object
            Aspose.Pdf.ExcelSaveOptions excelsave = new Aspose.Pdf.ExcelSaveOptions();

            // Save the output in XLS format
            pdfDocument.Save(dirPath + "convertedfile.xlsx", excelsave);

            //Open document in browser
            string docPath = @"wwwroot\files\convertedfile.xlsx";
            byte[] docBytes = await System.IO.File.ReadAllBytesAsync(docPath);
            await System.IO.File.WriteAllBytesAsync(docPath, docBytes);

            return docBytes;
        }

        /// <summary>
        /// This method with HttpGet attribute converts all files from .pdf to .jpg;
        /// </summary>
        /// <returns>Razor view of page</returns>
        [HttpGet]
        public IActionResult Pdf_To_Jpg()
        {
            return View();
        }

        /// <summary>
        /// This method with HttpPost attribute converts all files from .pdf to .jpg;
        /// </summary>
        /// <param name="uploadedFile">Contains all information about file, which will be convert</param>
        /// <returns></returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Pdf_To_Jpg(UploadPdfModel uploadedFile)
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
                    (await ConvertToJPGFile(dirPath, uploadedFile.Information.FileName), "application/zip");
            }

            return View();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dirPath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToJPGFile(string dirPath, string fileName)
        {
            // Load PDF document
            Aspose.Words.Document doc = new Aspose.Words.Document(dirPath + fileName);

            string _dirPath = dirPath + @"Test_zip\";
            Directory.CreateDirectory(_dirPath);

            // Path of zip archive
            string zipPath = dirPath + "Test_zip.zip";

            //Read each page from PDF as PNG
            for (int i = 0; i < doc.PageCount; i++)
            {
                // Save one page from PDF as PNG
                var extractedPage = doc.ExtractPages(i, 1);
                extractedPage.Save(_dirPath + $"test_{i}.png");
            }

            ZipFile.CreateFromDirectory(_dirPath, zipPath);

            //Open document in browser
            string docPath = dirPath + "Test_zip.zip";
            byte[] docBytes = await System.IO.File.ReadAllBytesAsync(docPath);
            await System.IO.File.WriteAllBytesAsync(docPath, docBytes);

            return docBytes;
        }
    }
}
