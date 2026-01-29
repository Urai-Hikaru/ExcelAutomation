using ExcelAutomation.Models.DTOs.Responses;

namespace ExcelAutomation.Data
{
    // アプリ全体で共有するデータを保持するクラス
    public static class UserSession
    {
        // ==========================================
        // プロパティ
        // ==========================================

        // API認証用トークン
        public static string? AccessToken { get; private set; }

        // ユーザー情報
        public static int UserId { get; private set; }
        public static string? Username { get; private set; }
        public static string? Role { get; private set; }

        // 有効期限
        public static DateTime Expiration { get; private set; }

        // ==========================================
        // 状態判定プロパティ
        // ==========================================

        /// <summary>
        /// ログイン済みかつ有効期限内かどうか
        /// </summary>
        public static bool IsLoggedIn => !string.IsNullOrEmpty(AccessToken) && DateTime.Now < Expiration;

        /// <summary>
        /// 権限確認
        /// </summary>
        public static bool IsAdmin => Role == "管理者";

        // ==========================================
        // 操作メソッド
        // ==========================================

        /// <summary>
        /// ログイン成功時にデータをセットする
        /// </summary>
        public static void Set(LoginResponseDto data)
        {
            AccessToken = data.Token;
            UserId = data.UserId;
            Username = data.Username;
            Role = data.Rolename;
            Expiration = data.Expiration;
        }

        /// <summary>
        /// ログアウト時にデータをクリアする
        /// </summary>
        public static void Clear()
        {
            AccessToken = null;
            UserId = 0;
            Username = null;
            Role = null;
            Expiration = DateTime.MinValue;
        }
    }
}