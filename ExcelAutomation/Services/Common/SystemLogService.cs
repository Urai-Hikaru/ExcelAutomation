using System.IO;
using System.Text;

namespace ExcelAutomation.Services.Common
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

        // 排他制御用のロックオブジェクト
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
        /// 情報ログを出力
        /// </summary>
        public void LogInfo(string message) => WriteLog(LogLevel.Info, message);

        /// <summary>
        /// 警告ログを出力
        /// </summary>
        public void LogWarning(string message) => WriteLog(LogLevel.Warning, message);

        /// <summary>
        /// エラーログを出力
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
            // ファイル名は日付で分ける
            string fileName = $"App_{DateTime.Now:yyyyMMdd}.log";
            string filePath = Path.Combine(_logDirectory, fileName);

            // ログのフォーマット
            string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

            // ロックをかけて書き込む
            lock (_lockObj)
            {
                try
                {
                    // ファイルに追記
                    using (var sw = new StreamWriter(filePath, append: true, encoding: Encoding.UTF8))
                    {
                        sw.WriteLine(logLine);
                    }
                }
                catch
                {
                    // ログ出力自体のエラーは、アプリを止めないために握りつぶすか、デバッグ出力のみにする
                }
            }
        }
    }
}