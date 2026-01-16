using ExcelAutomation.Commands;
using ExcelAutomation.Models;
using ExcelAutomation.Services;
using MaterialDesignThemes.Wpf;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace ExcelAutomation.ViewModels
{
    public class BatchProcessingViewModel : ViewModelBase
    {
        // ==========================================================
        // プロパティ定義
        // ==========================================================

        public SnackbarMessageQueue MessageQueue { get; } = new SnackbarMessageQueue();

        // 処理ロジックの選択状態
        private string _selectedLogicTag = "Sales";
        public string SelectedLogicTag
        {
            get => _selectedLogicTag;
            set => SetProperty(ref _selectedLogicTag, value);
        }
        public ObservableCollection<FileItemModel> FileList { get; set; } = new ObservableCollection<FileItemModel>();

        // UI制御用
        public bool IsBusy => StatusService.Instance.IsBusy;

        public ICommand ExecuteCommand { get; }

        public ICommand ClearCommand { get; }

        // ==========================================================
        // コンストラクタ
        // ==========================================================
        public BatchProcessingViewModel()
        {
            ExecuteCommand = new RelayCommand(async () => await ExecuteProcess(), () => !IsBusy);

            ClearCommand = new RelayCommand(() =>
            {
                FileList.Clear();
                return Task.CompletedTask;
            });

            StatusService.Instance.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(StatusService.IsBusy))
                {
                    OnPropertyChanged(nameof(IsBusy));
                    (ExecuteCommand as RelayCommand)?.RaiseCanExecuteChanged();
                }
            };
        }

        // ==========================================================
        // 処理実行ロジック
        // ==========================================================
        private async Task ExecuteProcess()
        {
            try
            {
                var targetFiles = FileList.Where(x => x.IsSelected).ToList();

                if (!targetFiles.Any())
                {
                    // ログ記録
                    SystemLogService.Instance.LogInfo("バッチ処理中断: 処理対象ファイルが選択されていません");
                    MessageQueue.Enqueue("処理対象のファイルが選択されていません。");
                    return;
                }

                // ログ記録: 開始
                SystemLogService.Instance.LogInfo($"バッチ処理を開始します。ロジック: {SelectedLogicTag}, 対象数: {targetFiles.Count}");

                StatusService.Instance.Start("バッチ処理を開始します...");

                var excelService = new ExcelService();
                var salesService = new SalesService();
                var pdfService = new PdfService();

                int totalCount = targetFiles.Count;
                int currentCount = 0;

                switch (SelectedLogicTag)
                {
                    case "Sales":

                        var invalidFiles = new List<string>();

                        int successCountSales = 0;

                        foreach (var file in targetFiles)
                        {
                            currentCount++;
                            int percent = (int)((double)currentCount / totalCount * 100);
                            StatusService.Instance.UpdateProgress(percent, $"売上登録中 ({currentCount}/{totalCount}): {file.FileName}");

                            // 拡張子を除いたファイル名を取得
                            string fileNameOnly = Path.GetFileNameWithoutExtension(file.FileName);

                            if (!DateTime.TryParseExact(fileNameOnly, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                            {
                                // 日付エラーはWarningとして記録
                                SystemLogService.Instance.LogWarning($"ファイル名日付形式エラー: {file.FileName}");
                                invalidFiles.Add(file.FileName);
                                continue;
                            }

                            try
                            {
                                await Task.Run(() =>
                                {
                                    // 1. ExcelServiceを使ってデータ読み込み
                                    var dataList = excelService.ReadSalesFile(file.FilePath, file.FileName);

                                    // 2. データがあればDBに一括登録
                                    if (dataList.Count > 0)
                                    {
                                        salesService.RegisterSalesBatch(dataList);
                                    }
                                });
                                successCountSales++;
                            }
                            catch (Exception ex)
                            {
                                // 個別ファイルのエラーログ
                                SystemLogService.Instance.LogError(ex, $"売上登録処理エラー: {file.FileName}");

                                MessageQueue.Enqueue($"エラー: {file.FileName} - {ex.Message}", "OK", () => { });
                            }
                        }

                        if (invalidFiles.Count > 0)
                        {
                            string msg = $"日付形式エラーのファイルが {invalidFiles.Count} 件ありました。";
                            MessageQueue.Enqueue(msg);
                        }

                        // 完了ログ
                        SystemLogService.Instance.LogInfo($"売上登録バッチ完了。成功: {successCountSales}件, 除外: {invalidFiles.Count}件");

                        MessageQueue.Enqueue($"売上登録が終了しました。(成功: {successCountSales}件)");
                        break;

                    case "Invoice":

                        string baseFolder = ConfigurationManager.AppSettings["InvoiceOutputFolder"]
                                            ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                        string outputFolder = Path.Combine(baseFolder, DateTime.Now.ToString("yyyyMMdd"));

                        if (!Directory.Exists(outputFolder))
                        {
                            Directory.CreateDirectory(outputFolder);
                        }

                        int successCountInvoice = 0;

                        foreach (var file in targetFiles)
                        {
                            currentCount++;

                            int percent = (int)((double)currentCount / totalCount * 100);
                            StatusService.Instance.UpdateProgress(percent, $"PDF作成中 ({currentCount}/{totalCount}): {file.FileName}");

                            string fileNameOnly = Path.GetFileNameWithoutExtension(file.FilePath);

                            string pdfPath = Path.Combine(outputFolder, fileNameOnly + ".pdf");

                            try
                            {
                                await Task.Run(() =>
                                {
                                    pdfService.ExportToPdf(file.FilePath, pdfPath);
                                });
                                successCountInvoice++;
                            }
                            catch (Exception ex)
                            {
                                // 個別ファイルのエラーログ
                                SystemLogService.Instance.LogError(ex, $"PDF作成処理エラー: {file.FileName}");

                                MessageQueue.Enqueue($"PDF作成失敗: {file.FileName}", "詳細", () => MessageBox.Show(ex.Message));
                            }
                        }

                        // 完了ログ
                        SystemLogService.Instance.LogInfo($"請求書PDF作成バッチ完了。成功: {successCountInvoice}件, 出力先: {outputFolder}");

                        MessageQueue.Enqueue(
                            $"PDF出力が完了しました。(成功: {successCountInvoice}件)",
                            "フォルダを開く",
                            () => Process.Start("explorer.exe", outputFolder)
                        );
                        break;

                    default:
                        SystemLogService.Instance.LogWarning($"不明な処理ロジックが選択されました: {SelectedLogicTag}");
                        MessageQueue.Enqueue("不明な処理ロジックです。");
                        break;
                }
            }
            catch (Exception ex)
            {
                // 致命的なエラーログ
                SystemLogService.Instance.LogError(ex, "バッチ処理実行中に致命的なエラーが発生しました");

                MessageBox.Show($"致命的なエラーが発生しました:\n{ex.Message}", "System Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                StatusService.Instance.End("待機中");
            }
        }
    }
}