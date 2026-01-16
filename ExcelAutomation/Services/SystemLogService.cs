using System.IO;
using System.Text;

namespace ExcelAutomation.Services
{
    // ログレベルの定義
    public enum LogLevel
    {
        Info,
        Warning,
        Error
    }

    public class SystemLogService
    {
        // シングルトン
        public static SystemLogService Instance { get; } = new SystemLogService();

        // ログファイルの保存先フォルダ（実行ファイル直下の Logs フォルダ）
        private readonly string _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

        // 排他制御用のロックオブジェクト（複数のスレッドから同時に書き込まれても大丈夫なようにする）
        private static readonly object _lockObj = new object();

        private SystemLogService()
        {
            // フォルダがなければ作成
            if (!Directory.Exists(_logDirectory))
            {
                Directory.CreateDirectory(_logDirectory);
            }
        }

        // ==========================================================
        // 公開メソッド
        // ==========================================================

        /// <summary>
        /// 情報ログを出力します
        /// </summary>
        public void LogInfo(string message) => WriteLog(LogLevel.Info, message);

        /// <summary>
        /// 警告ログを出力します
        /// </summary>
        public void LogWarning(string message) => WriteLog(LogLevel.Warning, message);

        /// <summary>
        /// エラーログを出力します（例外オブジェクト対応）
        /// </summary>
        public void LogError(Exception ex, string message = "エラーが発生しました")
        {
            // エラーメッセージ + 改行 + 例外の詳細(Stack Trace)
            string fullMessage = $"{message}{Environment.NewLine}Exception: {ex.GetType().Name}{Environment.NewLine}Message: {ex.Message}{Environment.NewLine}StackTrace: {ex.StackTrace}";
            WriteLog(LogLevel.Error, fullMessage);
        }

        // ==========================================================
        // 内部ロジック
        // ==========================================================

        private void WriteLog(LogLevel level, string message)
        {
            // ファイル名: App_20240116.log のように日付で分ける
            string fileName = $"App_{DateTime.Now:yyyyMMdd}.log";
            string filePath = Path.Combine(_logDirectory, fileName);

            // ログのフォーマット: [2024-01-16 12:00:00] [Info] 処理を開始しました
            string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

            // ロックをかけて書き込む（他スレッドからの干渉を防ぐ）
            lock (_lockObj)
            {
                try
                {
                    // ファイルに追記 (Append)
                    // エンコーディングはUTF-8
                    using (var sw = new StreamWriter(filePath, append: true, encoding: Encoding.UTF8))
                    {
                        sw.WriteLine(logLine);
                    }
                }
                catch
                {
                    // ログ出力自体のエラーは、アプリを止めないために握りつぶすか、デバッグ出力のみにする
                    // System.Diagnostics.Debug.WriteLine("ログ書き込み失敗");
                }
            }
        }
    }
}