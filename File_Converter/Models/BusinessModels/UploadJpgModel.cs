﻿using File_Converter.Models.ValidationAttributes;
using System.ComponentModel.DataAnnotations;

namespace File_Converter.Models.BusinessModels
{
    public class UploadJpgModel
    {
        [Required]
        [DataType(DataType.Upload)]
        [MaxFileSize(4 * 1024 * 1024)]
        [AllowedExtensions(new string[] { ".png" })]
        public IFormFile Information { get; set; }
    }
}
