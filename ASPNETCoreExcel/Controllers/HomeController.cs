using ASPNETCoreExcel.Models;
using ASPNETCoreExcel.Services;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ASPNETCoreExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ExcelImportService _excelService;
        public HomeController(ILogger<HomeController> logger, ExcelImportService excelService)
        {
            _logger = logger;
            _excelService = excelService;
        }

        public IActionResult Index()
        {
            return View(new List<Person>());
        }
        [HttpPost]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ViewBag.Error = "Lütfen bir Excel dosyası seçin.";
                return View("Index", new List<Person>());
            }

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Position = 0;

                var data = _excelService.ImportFromExcel(stream);
                return View("Index", data);
            }
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
