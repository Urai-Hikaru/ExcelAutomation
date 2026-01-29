using System.Runtime.InteropServices;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace ExcelAutomation.Services.Pdf
{
    public class PdfService
    {
        public void ExportToPdf(string inputExcelPath, string outputPdfPath)
        {
            InteropExcel.Application? excelApp = null;
            InteropExcel.Workbook? workbook = null;

            try
            {
                // Excelアプリケーションを起動（画面には表示しない）
                excelApp = new InteropExcel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                // ブックを開く
                workbook = excelApp.Workbooks.Open(inputExcelPath);

                // PDFとしてエクスポート
                workbook.ExportAsFixedFormat(
                    Type: InteropExcel.XlFixedFormatType.xlTypePDF,
                    Filename: outputPdfPath,
                    Quality: InteropExcel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    IgnorePrintAreas: false,
                    OpenAfterPublish: false
                );
            }
            catch (Exception ex)
            {
                throw new Exception($"PDF変換に失敗しました: {ex.Message}", ex);
            }
            finally
            {
                // ブックを閉じる（保存しない）
                if (workbook != null)
                {
                    workbook.Close(SaveChanges: false);
                    Marshal.ReleaseComObject(workbook);
                }

                // Excelアプリを終了
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}