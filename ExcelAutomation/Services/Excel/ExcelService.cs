using ExcelAutomation.Models.Entities;
using ExcelAutomation.ViewModels;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace ExcelAutomation.Services.Excel
{
    public class ExcelService
    {
        public ExcelService()
        {
        }

        /// <summary>
        /// 月次集計データをExcelファイルとして出力
        /// </summary>
        public void CreateMonthlyReport(string savePath, List<MonthlySummaryModel> data, DateTime targetMonth)
        {
            InteropExcel.Application? excelApp = null;
            InteropExcel.Workbook? workbook = null;
            InteropExcel.Worksheet? sheet = null;
            InteropExcel.Range? titleRange = null;
            InteropExcel.Range? headerRange = null;
            InteropExcel.Range? dataRange = null;
            InteropExcel.Range? finalRange = null;

            try
            {
                excelApp = new InteropExcel.Application { Visible = false, DisplayAlerts = false };
                workbook = excelApp.Workbooks.Add();
                sheet = (InteropExcel.Worksheet)workbook.Sheets[1];
                sheet.Name = $"{targetMonth.Month}月集計";

                object[,] values = new object[data.Count, 3];

                for (int i = 0; i < data.Count; i++)
                {
                    values[i, 0] = data[i].ProductCode;
                    values[i, 1] = data[i].TotalQuantity;
                    values[i, 2] = data[i].TotalPrice;
                }

                sheet.Cells[1, 1] = $"{targetMonth:yyyy年MM月} 月次売上集計表";
                titleRange = sheet.Range["A1", "C1"];
                titleRange.Merge();
                titleRange.Font.Size = 14;
                titleRange.Font.Bold = true;
                titleRange.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;

                string[] headers = { "商品コード", "合計数量", "合計金額" };
                for (int col = 0; col < headers.Length; col++)
                {
                    sheet.Cells[2, col + 1] = headers[col];
                }

                headerRange = sheet.Range["A2", "C2"];
                headerRange.Interior.Color = System.Drawing.Color.LightGray;
                headerRange.Font.Bold = true;
                headerRange.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;

                if (data.Count > 0)
                {
                    InteropExcel.Range startCell = sheet.Cells[3, 1];
                    InteropExcel.Range endCell = sheet.Cells[3 + data.Count - 1, 3];
                    dataRange = sheet.Range[startCell, endCell];
                    dataRange.Value2 = values;

                    finalRange = sheet.Range["A2", endCell];
                    finalRange.Borders.LineStyle = InteropExcel.XlLineStyle.xlContinuous;

                    InteropExcel.Range qtyCol = sheet.Range[sheet.Cells[3, 2], endCell];
                    qtyCol.NumberFormat = "#,##0";

                    InteropExcel.Range priceCol = sheet.Range[sheet.Cells[3, 3], endCell];
                    priceCol.NumberFormat = "#,##0";
                }

                sheet.Columns.AutoFit();

                if (File.Exists(savePath))
                {
                    File.Delete(savePath);
                }

                workbook.SaveAs(savePath);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (finalRange != null) Marshal.ReleaseComObject(finalRange);
                if (dataRange != null) Marshal.ReleaseComObject(dataRange);
                if (headerRange != null) Marshal.ReleaseComObject(headerRange);
                if (titleRange != null) Marshal.ReleaseComObject(titleRange);
                if (sheet != null) Marshal.ReleaseComObject(sheet);

                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Excelファイルから売上データを読み込み
        /// </summary>
        public List<SalesHistory> ReadSalesFile(string filePath, string fileName)
        {
            // ファイルロックチェック
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                }
            }
            catch (IOException)
            {
                throw new Exception($"ファイル '{fileName}' が開かれています。\nExcelを閉じてから再試行してください。");
            }

            var list = new List<SalesHistory>();
            InteropExcel.Application? excelApp = null;
            InteropExcel.Workbook? workbook = null;
            InteropExcel.Worksheet? sheet = null;
            DateTime now = DateTime.Now;

            try
            {
                excelApp = new InteropExcel.Application { Visible = false };
                workbook = excelApp.Workbooks.Open(filePath);
                sheet = (InteropExcel.Worksheet)workbook.Sheets[1];

                InteropExcel.Range usedRange = sheet.UsedRange;
                object[,] values = (object[,])usedRange.Value2;

                for (int i = 2; i <= values.GetLength(0); i++)
                {
                    if (values[i, 1] == null) continue;

                    string productCode = values[i, 2]?.ToString() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(productCode))
                    {
                        throw new Exception($"ファイル '{fileName}' の {i} 行目: 商品コード(2列目)が取得できません。");
                    }

                    object rawQty = values[i, 3];
                    if (rawQty == null || !int.TryParse(rawQty.ToString(), out int qty))
                    {
                        throw new Exception($"ファイル '{Path.GetFileName(filePath)}' の {i} 行目: 数量(3列目)が取得できません。");
                    }

                    object rawUnitPrice = values[i, 4];
                    if (rawUnitPrice == null || !int.TryParse(rawUnitPrice.ToString(), out int unitPrice))
                    {
                        throw new Exception($"ファイル '{Path.GetFileName(filePath)}' の {i} 行目: 単価(4列目)が取得できません。");
                    }

                    object rawTotalPrice = values[i, 5];
                    if (rawTotalPrice == null || !int.TryParse(rawTotalPrice.ToString(), out int totalPrice))
                    {
                        throw new Exception($"ファイル '{Path.GetFileName(filePath)}' の {i} 行目: 合計金額(5列目)が取得できません。");
                    }

                    var data = new SalesHistory();
                    // ファイル名(yyyyMMdd)から日付を設定
                    data.SalesDate = DateTime.ParseExact(Path.GetFileNameWithoutExtension(fileName), "yyyyMMdd", CultureInfo.InvariantCulture);
                    data.ProductCode = productCode;
                    data.Quantity = qty;
                    data.UnitPrice = unitPrice;
                    data.TotalPrice = totalPrice;
                    data.CreatedBy = "ExcelAutomation";
                    data.CreatedAt = now;
                    data.UpdatedBy = "ExcelAutomation";
                    data.UpdatedAt = now;

                    list.Add(data);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return list;
        }
    }
}