using Excel_NPOI;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using System.Data;

namespace NPOI_dotnetcore6MVC.Controllers
{
    public class NPOIController : Controller
    {
        private readonly IHostEnvironment _hostingEnvironment; // 用 DI 加入 Server wwwroot的 根目錄路徑
        public NPOIController(IHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment; // 用 DI 加入 Server wwwroot的 根目錄路徑
        }

        public IActionResult Index()
        {
            // 需要從 nuget 安裝 NPOI， dotnet core 安裝的是這個 https://www.nuget.org/packages/DotNetCore.NPOI
            MyNPOI writer = new MyNPOI(); // 自己組裝方便存取 Excel NPOI 的元件，位置在 NOPOClass.cs
            var destPath = _hostingEnvironment.ContentRootPath + "/wwwroot/Excel/"; // Excel 檔案的根目錄路徑
            var sourcefilename = "活頁簿1.xlsx"; // 宣告來源檔名
            writer.open(destPath + sourcefilename); // 開啟 excel

            var cellPosition = writer.ExcelCoordinateToCellPosition("A1"); // 選擇要寫入哪一格

            // 輸入資料，前面兩個參數 cellPosition.Item1 和 cellPosition.Item2 對應 excel 座標軸，如果不從上面取A1，也可以自己用數字代入
            writer.SetCell(cellPosition.Item1, cellPosition.Item2, "寫入A1的值", CellType.String);


            var cellPosition2 = writer.ExcelCoordinateToCellPosition("D7");
            writer.SetCell(cellPosition2.Item1, cellPosition2.Item2, "D7被後端程式寫入了資料", CellType.String);
            writer.SetCell(8, 5, "這格是F9，SetCell 起點是0，座標8,5，先數縱向座標軸，再數橫向座標軸", CellType.String);

            var savefilename = "活頁簿1_儲存.xlsx"; // 存檔的檔案
            writer.SaveClose(destPath + savefilename); // 儲存

            return View();

        }
    }
}
