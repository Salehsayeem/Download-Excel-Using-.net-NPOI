using Microsoft.AspNetCore.Mvc;

namespace DownloadExcel.Models
{
    public class ApiResponse
    {
        public int Status { get; set; }
        public FileContentResult File { get; set; }
    }
    public class ListWrapper<T>
    {
        public List<T> Items { get; set; }
    }
}
