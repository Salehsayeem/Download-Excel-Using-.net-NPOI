using DownloadExcel.Helpers;
using DownloadExcel.Models;
using Microsoft.AspNetCore.Mvc;

namespace DownloadExcel.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
                {
                    Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                    TemperatureC = Random.Shared.Next(-20, 55),
                    Summary = Summaries[Random.Shared.Next(Summaries.Length)]
                })
                .ToArray();
        }

        [HttpGet("DownloadExcelApi")]
        public ApiResponse Download(CancellationToken ct)
        {

            var products = new List<Product>
            {
                new Product()
                {
                    Id = Guid.NewGuid(),
                    Name = "Product 1",
                    Quantity = 10,
                    Price = 100,
                    IsActive = true,
                    ExpiryDate = DateTime.Now.AddDays(10)

                },
                new Product()
                {
                    Id = Guid.NewGuid(),
                    Name = "Product 2",
                    Quantity = 20,
                    Price = 200,
                    IsActive = true,
                    ExpiryDate = DateTime.Now.AddDays(20)
                }
            };
            var file = ExcelHelper.CreateFile(products);
            var f = File(file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "products.xlsx");

            return new ApiResponse()
            {
                Status = 200,
                File = f,
            };

        }
    }
}