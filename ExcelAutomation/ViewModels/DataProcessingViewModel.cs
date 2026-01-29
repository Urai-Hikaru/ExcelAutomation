using ExcelAutomation.Commands;
using ExcelAutomation.Models.Entities;
using ExcelAutomation.Services;
using ExcelAutomation.Services.Common;
using ExcelAutomation.Services.Excel;
using MaterialDesignThemes.Wpf;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace ExcelAutomation.ViewModels
{
    public class DataProcessingViewModel : BaseViewModel
    {
        private readonly SalesService _salesService = new SalesService();

        // ==========================================================
        // プロパティ定義
        // ==========================================================
        public SnackbarMessageQueue MessageQueue { get; } = new SnackbarMessageQueue();

        private DateTime _targetDate;
        public DateTime TargetDate
        {
            get => _targetDate;
            set => SetProperty(ref _targetDate, value);
        }

        private ObservableCollection<SalesHistoryGridModel> _dailySalesList = new ObservableCollection<SalesHistoryGridModel>();
        public ObservableCollection<SalesHistoryGridModel> DailySalesList
        {
            get => _dailySalesList;
            set => SetProperty(ref _dailySalesList, value);
        }

        // UI制御用
        public bool IsBusy => StatusService.Instance.IsBusy;

        public ICommand SearchCommand { get; }
        public ICommand ExportCommand { get; }

        // ==========================================================
        // コンストラクタ
        // ==========================================================
        public DataProcessingViewModel()
        {
            TargetDate = DateTime.Today;
            SearchCommand = new RelayCommand(async () => await SearchProcess(), () => !IsBusy);
            ExportCommand = new RelayCommand(async () => await ExportProcess(), () => !IsBusy);

            StatusService.Instance.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(StatusService.IsBusy))
                {
                    OnPropertyChanged(nameof(IsBusy)); // Viewに通知
                    (SearchCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
                }
            };
        }

        // ==========================================================
        // 処理実行ロジック
        // ==========================================================
        private async Task SearchProcess()
        {
            try
            {
                SystemLogService.Instance.LogInfo($"データ検索を開始します。対象日: {TargetDate:yyyy/MM/dd}");

                StatusService.Instance.Start("データを検索中...");

                StatusService.Instance.UpdateProgress(20, "データベースに問い合わせ中...");

                var result = await Task.Run(() => _salesService.GetDailySales(TargetDate));

                StatusService.Instance.UpdateProgress(60, "画面に表示しています...");

                // UIコレクションに追加
                DailySalesList = new ObservableCollection<SalesHistoryGridModel>(result);

                if (result.Count == 0)
                {
                    SystemLogService.Instance.LogInfo("検索結果: 0件");
                    MessageQueue.Enqueue("該当するデータはありませんでした。");
                }
                else
                {
                    SystemLogService.Instance.LogInfo($"検索完了: {result.Count} 件取得");
                    MessageQueue.Enqueue($"検索完了: {result.Count} 件のデータが見つかりました。");
                }
            }
            catch (Exception ex)
            {
                SystemLogService.Instance.LogError(ex, "データ検索処理でエラーが発生しました");

                MessageQueue.Enqueue("検索中にエラーが発生しました", "詳細", () => MessageBox.Show(ex.Message));
            }
            finally
            {
                StatusService.Instance.End("待機中");
            }
        }

        private async Task ExportProcess()
        {
            try
            {
                SystemLogService.Instance.LogInfo($"Excel出力を開始します。対象月: {TargetDate:yyyy/MM}");

                StatusService.Instance.Start("Excel出力中...");

                // フォルダパスの準備
                string baseFolder = ConfigurationManager.AppSettings["SalesOutputFolder"]
                                    ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string outputFolder = Path.Combine(baseFolder, "MonthlySummary");

                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                string fileName = $"Summary_{TargetDate:yyyyMM}.xlsx";
                string savePath = Path.Combine(outputFolder, fileName);

                // 期間計算
                DateTime startDate = new DateTime(TargetDate.Year, TargetDate.Month, 1);
                DateTime endDate = startDate.AddMonths(1).AddDays(-1);

                await Task.Run(() =>
                {
                    StatusService.Instance.UpdateProgress(30, "データベースから集計データを取得中...");

                    var summaryData = _salesService.GetMonthlySummary(startDate, endDate);

                    if (summaryData.Count == 0)
                    {
                        MessageQueue.Enqueue($"{TargetDate:yyyy年MM月} のデータが存在しません。");
                        return;
                    }

                    StatusService.Instance.UpdateProgress(70, "Excel帳票を作成・保存中...");

                    // Excel出力
                    var excelService = new ExcelService();
                    excelService.CreateMonthlyReport(savePath, summaryData, TargetDate);

                    SystemLogService.Instance.LogInfo($"Excel出力完了: {savePath}");

                    // 完了通知
                    MessageQueue.Enqueue(
                        "月次集計ファイルの出力が完了しました",
                        "開く",
                        () => Process.Start(new ProcessStartInfo(savePath) { UseShellExecute = true })
                    );
                });
            }
            catch (Exception ex)
            {
                SystemLogService.Instance.LogError(ex, "Excel出力処理でエラーが発生しました");

                MessageQueue.Enqueue("Excel出力中にエラーが発生しました", "詳細", () => MessageBox.Show(ex.Message));
            }
            finally
            {
                StatusService.Instance.End("待機中");
            }
        }
    }
}