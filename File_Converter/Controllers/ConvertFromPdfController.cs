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

                return File
                    (await ConvertToDOCXFile
                    (@"wwwroot\files\", uploadedFile.Name, uploadedFile.Type), 
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileDownloadName: uploadedFile.Name+ ".docx");
            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert WORD files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToDOCXFile(string dirPath, string fileName, string fileType)
        {
            // Open the source PDF document
            Aspose.Pdf.Document document = new Aspose.Pdf.Document(dirPath + fileName+ fileType);

            // Save the file into MS document format
            document.Save(dirPath + fileName + ".docx", Aspose.Pdf.SaveFormat.DocX);

            //Open document in browser
            string docPath = dirPath + fileName + ".docx";
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

                return File
                    (await ConvertToPPTXFile
                    (@"wwwroot\files\", uploadedFile.Name, uploadedFile.Type),
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation", fileDownloadName: uploadedFile.Name + ".pptx");
            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert PowerPoint files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToPPTXFile(string dirPath, string fileName, string fileType)
        {
            // Load PDF document
            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(dirPath + fileName + fileType);
            // Instantiate PptxSaveOptions instance
            Aspose.Pdf.PptxSaveOptions pptx_save = new Aspose.Pdf.PptxSaveOptions();
            // Save the output in PPTX format
            doc.Save(dirPath + fileName + ".pptx", pptx_save);

            //Open document in browser
            string docPath = dirPath + fileName + ".pptx";
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

                return File
                    (await ConvertToXLSXFile
                    (@"wwwroot\files\", uploadedFile.Name, uploadedFile.Type),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileDownloadName: uploadedFile.Name + ".xlsx");
            }

            return View();
        }

        /// <summary>
        /// This asynchronous method convert Excel files to PDF;
        /// </summary>
        /// <param name="dirPath">Path to the current file directory</param>
        /// <param name="fileName">File name which will be converted</param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToXLSXFile(string dirPath, string fileName, string fileType)
        {
            // Load PDF document
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(dirPath + fileName + fileType);

            // Instantiate ExcelSave Option object
            Aspose.Pdf.ExcelSaveOptions excelsave = new Aspose.Pdf.ExcelSaveOptions();

            // Save the output in XLS format
            pdfDocument.Save(dirPath + fileName + ".xlsx", excelsave);

            //Open document in browser
            string docPath = dirPath + fileName + ".xlsx";
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

                return File
                    (await ConvertToJPGFile
                    (@"wwwroot\files\", uploadedFile.Name, uploadedFile.Type),
                    "application/zip", fileDownloadName: uploadedFile.Name + ".zip");
            }

            return View();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dirPath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToJPGFile(string dirPath, string fileName, string fileType)
        {
            // Load PDF document
            Aspose.Words.Document doc = new Aspose.Words.Document(dirPath + fileName + fileType);

            string _dirPath = dirPath + fileName + @"\";

            if(Directory.Exists(_dirPath))
            {
                Directory.Delete(_dirPath, true);
            }

            Directory.CreateDirectory(_dirPath);

            // Path of zip archive
            string zipPath = dirPath + fileName + ".zip";

            //Read each page from PDF as PNG
            for (int i = 0; i < doc.PageCount; i++)
            {
                // Save one page from PDF as PNG
                var extractedPage = doc.ExtractPages(i, 1);
                extractedPage.Save(_dirPath + fileName + $"{i}.png");
            }

            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }

            ZipFile.CreateFromDirectory(_dirPath, zipPath);

            //Open document in browser
            string docPath = dirPath + fileName + ".zip";
            byte[] docBytes = await System.IO.File.ReadAllBytesAsync(docPath);
            await System.IO.File.WriteAllBytesAsync(docPath, docBytes);

            return docBytes;
        }
    }
}
