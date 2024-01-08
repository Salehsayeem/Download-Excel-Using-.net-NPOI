using Microsoft.AspNetCore.Mvc;

namespace DownloadExcel.Models
{
    public class ApiResponse
    {
        public int Status { get; set; }
        public FileContentResult File { get; set; }
    }
}
