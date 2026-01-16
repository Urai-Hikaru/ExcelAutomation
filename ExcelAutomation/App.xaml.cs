using ExcelAutomation.Data;
using ExcelAutomation.Services;
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

            // ログ機能の初期化（未処理例外の捕捉設定）
            SetupExceptionHandling();

            // アプリケーション起動ログ
            SystemLogService.Instance.LogInfo("=== アプリケーションを起動しました ===");

            try
            {
                // Dapperでアンダースコア付きカラム名(sales_date)を自動マッピングする設定
                Dapper.DefaultTypeMap.MatchNamesWithUnderscores = true;

                // アプリ起動時に、DBがなければ自動作成し、最新の状態にする
                using (var context = new AppDbContext())
                {
                    context.Database.Migrate();
                }
            }
            catch (Exception ex)
            {
                // DB接続やマイグレーションに失敗した場合
                SystemLogService.Instance.LogError(ex, "データベースの初期化に失敗しました。アプリケーションを終了します。");

                MessageBox.Show($"データベース接続エラーが発生しました。\nログファイルを確認してください。\n\n{ex.Message}",
                                "Startup Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);

                // DBが使えないとアプリが動作しないため終了させる
                Current.Shutdown();
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // アプリケーション終了ログ
            SystemLogService.Instance.LogInfo("=== アプリケーションを終了しました ===");
            base.OnExit(e);
        }

        /// <summary>
        /// 未処理の例外を捕捉する設定を行います
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

            // trueにするとアプリの強制終了を防げるが、状態が不安定なため基本的には落とすか、慎重に判断する
            // ここでは false (デフォルト) のままにして、エラー後にアプリを終了させる挙動とします
            // e.Handled = true; 
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