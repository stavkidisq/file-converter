using File_Converter.Models.ValidationAttributes;
using System.ComponentModel.DataAnnotations;

namespace File_Converter.Models.BusinessModels
{
    public class UploadPdfModel
    {
        public string Name
            => Information.FileName.Substring(0, Information.FileName.IndexOf('.'));
        public string Type
            => Information.FileName.Substring(Information.FileName.IndexOf('.'), Information.FileName.Length - Information.FileName.IndexOf('.'));

        [Required(ErrorMessage = "Please select a file.")]
        [DataType(DataType.Upload)]
        [MaxFileSize(10 * 1024 * 1024)]
        [AllowedExtensions(new string[] { ".pdf" })]
        public IFormFile Information { get; set; }
    }
}
