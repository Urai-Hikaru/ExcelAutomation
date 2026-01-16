using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAutomation.Services
{
    public class PdfService
    {
        public void ExportToPdf(string inputExcelPath, string outputPdfPath)
        {
            Excel.Application? excelApp = null;
            Excel.Workbook? workbook = null;

            try
            {
                // Excelアプリケーションを起動（画面には表示しない）
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false // 保存時の確認メッセージなどを抑制
                };

                // ブックを開く
                workbook = excelApp.Workbooks.Open(inputExcelPath);

                // PDFとしてエクスポート
                workbook.ExportAsFixedFormat(
                    Type: Excel.XlFixedFormatType.xlTypePDF,
                    Filename: outputPdfPath,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
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

                // ガベージコレクションを強制実行してCOMオブジェクトを確実に解放
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}