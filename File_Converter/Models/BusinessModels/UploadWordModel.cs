using File_Converter.Models.ValidationAttributes;
using System.ComponentModel.DataAnnotations;

namespace File_Converter.Models.BusinessModels
{
    public class UploadWordModel
    {
        public string Name 
            => Information.FileName.Substring(0, Information.FileName.IndexOf('.'));
        public string Type 
            => Information.FileName.Substring(Information.FileName.IndexOf('.'), Information.FileName.Length - Information.FileName.IndexOf('.'));

        [Required]
        [DataType(DataType.Upload)]
        [MaxFileSize(4 * 1024 * 1024)]
        [AllowedExtensions(new string[] { ".docx", ".doc" })]
        public IFormFile Information { get; set; }
    }
}
