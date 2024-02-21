using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel_NPOI
{

    /// <summary>
    /// C# WriteExcel & ReadExcel (NPOI/OLE)
    /// 參考來源 https://ithelp.ithome.com.tw/articles/10212249
    /// </summary>
    /// <returns></returns>
    public class MyNPOI
    {
        public ISheet sheet;
        public FileStream fileStream;
        public IWorkbook workbook = null; //新建IWorkbook對象 

        /// <summary>
        /// 開啟檔案
        /// </summary>
        /// <param name="fileName"></param>
        public void open(String fileName)
        {
            try
            {
                fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本 
                {
                    workbook = new XSSFWorkbook(fileStream); //xlsx數據讀入workbook 
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本 
                {
                    workbook = new HSSFWorkbook(fileStream); //xls數據讀入workbook 
                }
                sheet = workbook.GetSheetAt(0); //獲取第一個工作表 

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 寫入文字資料到excel儲存格
        /// </summary>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <param name="value"></param>
        /// <param name="_celltype"></param>
        public void SetCell(int iRow, int iCol, string value, CellType _celltype)
        {
            IRow row;
            ICell cell = null;
            if (sheet.GetRow(iRow) != null)
                row = sheet.GetRow(iRow);
            else
            {
                //int ostatniWiersz = sheet.LastRowNum;
                //row = (HSSFRow)sheet.CreateRow(ostatniWiersz + 1);//這樣會有問題
                row = sheet.CreateRow(iRow);//add row
            }

            if (row != null)
            {
                cell = row.GetCell(iCol);
                if (cell == null)
                {
                    cell = row.CreateCell(iCol, _celltype);//add cell
                }
                if (cell != null)
                {
                    //cell.SetCellType ( _celltype);//reset type不用reset也可以
                    if (_celltype == NPOI.SS.UserModel.CellType.Numeric)
                        cell.SetCellValue(double.Parse(value));
                    else if (_celltype == NPOI.SS.UserModel.CellType.Formula)
                        cell.SetCellFormula(value);
                    else
                        cell.SetCellValue(value);

                }
            }
        }

        public void Clear(int ifromRow)
        {
            for (int i = (sheet.FirstRowNum + 0); i <= sheet.LastRowNum; i++)   //-- 每一列做迴圈
            {
                HSSFRow row = (HSSFRow)sheet.GetRow(i);  //--不包含 Excel表頭列的 "其他資料列"

                if (row != null)
                {
                    if (i >= ifromRow)
                    {
                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)   //-- 每一個欄位做迴圈
                        {
                            SetCell(i, j, "", NPOI.SS.UserModel.CellType.Blank);
                            //CellType.Blank);不會清空格式化的cell
                            //CellType.Formula);清空格式化的cell,也清不是格式化的
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 儲存與關閉檔案
        /// </summary>
        /// <param name="path"></param>
        public void SaveClose(string path)
        {
            FileStream fs = null;
            try
            {
                sheet.ForceFormulaRecalculation = true;//更新公式的值 
                fs = new FileStream(path, FileMode.Create);
                workbook.Write(fs);
                fs.Close();

            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                throw ex;
            }
            finally
            {
                fileStream.Close();
            }
        }

        /// <summary>
        /// 讀Excel 參考來源 https://ithelp.ithome.com.tw/articles/10212249 https://www.796t.com/content/1537631886.html
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public DataTable getexcel(String fileName)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook = null; //新建IWorkbook對象 
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本 
            {
                workbook = new XSSFWorkbook(fileStream); //xlsx數據讀入workbook 
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本 
            {
                workbook = new HSSFWorkbook(fileStream); //xls數據讀入workbook 
            }
            ISheet sheet = workbook.GetSheetAt(0); //獲取第一個工作表 
            IRow row;// = sheet.GetRow(0); //新建當前工作表行數據 
                     // MessageBox.Show(sheet.LastRowNum.ToString());
            row = sheet.GetRow(0); //row讀入頭部
            if (row != null)
            {
                for (int m = 0; m < row.LastCellNum; m++) //表頭 
                {
                    string cellValue = row.GetCell(m).ToString(); //獲取i行j列數據 
                    Console.WriteLine(cellValue);
                    dt.Columns.Add(cellValue);
                }
            }
            for (int i = 1; i <= sheet.LastRowNum; i++) //對工作表每一行 
            {
                System.Data.DataRow dr = dt.NewRow();
                row = sheet.GetRow(i); //row讀入第i行數據 
                if (row != null)
                {
                    for (int j = 0; j < row.LastCellNum; j++) //對工作表每一列 
                    {
                        string cellValue = row.GetCell(j).ToString(); //獲取i行j列數據 
                        Console.WriteLine(cellValue);
                        dr[j] = cellValue;
                    }
                }
                dt.Rows.Add(dr);
            }
            //Console.ReadLine();//這個有問題,讀不出來,反正它只是debug用的,所以取消它
            fileStream.Close();
            return dt;
        }

        /// <summary>
        /// 將datatable對象保存為Excel文件
        /// 提供Excel保存路徑及datatable數據對象，成功返回真，失敗返回假。
        /// 參考來源 https://ithelp.ithome.com.tw/articles/10212249 https://www.796t.com/content/1537631886.html
        /// </summary>
        /// <param name="path"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool DataTableToExcel(String path, DataTable dt)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new HSSFWorkbook();
                    sheet = workbook.CreateSheet("Sheet0");//創建一個名稱為Sheet0的表 
                    int rowCount = dt.Rows.Count;//行數 
                    int columnCount = dt.Columns.Count;//列數

                    //設置列頭 
                    row = sheet.CreateRow(0);//excel第一行設為列頭 
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //設置每行每列的單元格, 
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行開始寫入數據 
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    using (fs = System.IO.File.OpenWrite(path))
                    {
                        workbook.Write(fs);//向打開的這個xls文件中寫入數據 
                        result = true;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                // todo: 記錄例外情形 ex.ToString()
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }
        }


        public Tuple<int, int> ExcelCoordinateToCellPosition(string excelCoordinate)
        {
            // 使用正則表達式將字母和數字分離
            var match = Regex.Match(excelCoordinate, @"([A-Za-z]+)(\d+)");

            if (match.Success)
            {
                // 取得字母部分（列）
                string columnStr = match.Groups[1].Value.ToUpper();
                int column = GetExcelColumnNumber(columnStr);

                // 取得數字部分（行）
                int row = int.Parse(match.Groups[2].Value);

                return new Tuple<int, int>(row - 1, column - 1);
            }

            // 如果座標格式不正確，返回 (-1, -1) 或者拋出異常，取決於你的需求
            return new Tuple<int, int>(-1, -1);
        }

        /// <summary>
        /// 取得Excel欄位數字
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public int GetExcelColumnNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");
            columnName = columnName.ToUpperInvariant();
            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

        /// <summary>
        /// 取得Excel欄位名稱
        /// </summary>
        /// <param name="columnNumber">第幾欄，例如 4回傳D </param>
        /// <returns></returns>
        public string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }
    }

}
