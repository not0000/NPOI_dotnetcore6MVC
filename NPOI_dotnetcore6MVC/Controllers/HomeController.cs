using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using NPOI_dotnetcore6MVC.Models;
using System;
using System.Diagnostics;

namespace NPOI_dotnetcore6MVC.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IHostEnvironment _hostingEnvironment; // 用 DI 加入 Server wwwroot的 根目錄路徑

        public HomeController(ILogger<HomeController> logger, IHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment; // 用 DI 加入 Server wwwroot的 根目錄路徑
        }

        public IActionResult Index()
        {
            return View();
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

        public string HowToUseSession()
        {
            /// 使用 Session
            var personObj = new MySession()
            {
                Name = "New Person In Session"
            };

            HttpContext.Session.SetComplexObjectSession("John Doe", personObj); // 物件寫入 Session 
            var objFromSession = HttpContext.Session.GetComplexObjectSession<MySession>("John Doe"); // 讀取Session的物件

            return objFromSession.Name;
        }

        public class MySession
        {
            public MySession()
            {
                Id = 0;
                Name = "";
            }

            public int Id { get; set; }
            public string Name { get; set; }

        }
    }
}