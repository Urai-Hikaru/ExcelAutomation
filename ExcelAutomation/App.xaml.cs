using ExcelAutomation.Data;
using ExcelAutomation.Services.Common;
using ExcelAutomation.Views;
using Microsoft.EntityFrameworkCore;
using System.Windows;
using System.Windows.Threading;

namespace ExcelAutomation
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            this.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            // ログ機能の初期化
            SetupExceptionHandling();

            // アプリケーション起動ログ
            SystemLogService.Instance.LogInfo("=== アプリケーションを起動しました ===");

            try
            {
                // Dapperでアンダースコア付きカラム名を自動マッピングする設定
                Dapper.DefaultTypeMap.MatchNamesWithUnderscores = true;

                // アプリ起動時に、DBがなければ自動作成
                using (var context = new AppDbContext())
                {
                    context.Database.Migrate();
                }

                // ==============================================
                // 3. ログイン画面の表示ロジック (ここを追加)
                // ==============================================

                // DB接続に成功したらログイン画面を表示
                var loginWindow = new LoginWindow();

                // ShowDialog() は画面が閉じるまで待機
                bool? result = loginWindow.ShowDialog();

                if (result == true)
                {
                    // ログイン成功 -> メイン画面を表示
                    var mainWindow = new MainWindow();

                    this.MainWindow = mainWindow;

                    mainWindow.Show();

                    this.ShutdownMode = ShutdownMode.OnMainWindowClose;
                }
                else
                {
                    // ログインキャンセル/失敗 -> アプリ終了
                    SystemLogService.Instance.LogInfo("ログインがキャンセルされたため、終了します。");
                    Current.Shutdown();
                }
                // ==============================================
            }
            catch (Exception ex)
            {
                SystemLogService.Instance.LogError(ex, "初期化エラー");
                MessageBox.Show($"エラーが発生しました。\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Current.Shutdown();
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            SystemLogService.Instance.LogInfo("=== アプリケーションを終了しました ===");
            base.OnExit(e);
        }

        /// <summary>
        /// 未処理の例外を捕捉する設定
        /// </summary>
        private void SetupExceptionHandling()
        {
            // UIスレッドで発生した未処理例外
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;

            // バックグラウンドスレッド（Task内など）で発生した未処理例外
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            // ログ出力
            SystemLogService.Instance.LogError(e.Exception, "【UIスレッド】予期せぬエラーが発生しました");

            // ユーザーへの通知
            MessageBox.Show($"予期せぬエラーが発生しました。\n\n{e.Exception.Message}",
                            "System Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                // ログ出力
                SystemLogService.Instance.LogError(ex, "【バックグラウンド】予期せぬエラーが発生しました");

                // バックグラウンドスレッドのエラーは、ここでMessageBoxを出しても表示されないことがあるが、念のため記述
                MessageBox.Show($"致命的なシステムエラーが発生しました。\n\n{ex.Message}",
                                "Critical Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }
    }
}