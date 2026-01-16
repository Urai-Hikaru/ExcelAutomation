using Dapper;
using ExcelAutomation.Models;
using System.Configuration;
using System.Data.SQLite;

namespace ExcelAutomation.Services
{
    public class SalesService
    {
        // 接続文字列
        private static readonly string ConnectionString = ConfigurationManager.AppSettings["SqliteConnectionString"] ?? string.Empty;

        /// <summary>
        /// 指定日の売上明細を取得します
        /// </summary>
        public List<SalesHistoryGridModel> GetDailySales(DateTime targetDate)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT 
                        sales_date, 
                        product_code, 
                        quantity, 
                        unit_price, 
                        total_price
                    FROM SalesHistory 
                    WHERE sales_date = @TargetDate 
                    ORDER BY product_code";

                return conn.Query<SalesHistoryGridModel>(sql, new { TargetDate = targetDate.Date }).ToList();
            }
        }

        /// <summary>
        /// 指定期間の月次集計データを取得します
        /// </summary>
        public List<MonthlySummaryModel> GetMonthlySummary(DateTime startDate, DateTime endDate)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT 
                        product_code, 
                        SUM(quantity) AS TotalQuantity, 
                        SUM(total_price) AS TotalPrice 
                    FROM SalesHistory 
                    WHERE sales_date >= @StartDate AND sales_date <= @EndDate
                    GROUP BY product_code
                    ORDER BY product_code";

                return conn.Query<MonthlySummaryModel>(sql, new { StartDate = startDate, EndDate = endDate }).ToList();
            }
        }

        /// <summary>
        /// 売上データを一括登録します（同日のデータは洗い替え）
        /// </summary>
        public void RegisterSalesBatch(List<SalesHistory> list)
        {
            if (list == null || list.Count == 0) return;

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // 登録対象の日付リスト（重複なし）を取得
                        var targetDates = list.Select(x => x.SalesDate).Distinct().ToList();

                        // 1. 既存データの削除 (対象日のみ)
                        string deleteSql = "DELETE FROM SalesHistory WHERE sales_date = @SalesDate";
                        conn.Execute(deleteSql, targetDates.Select(d => new { SalesDate = d }), transaction: transaction);

                        // 2. 新規データの登録
                        string insertSql = @"
                            INSERT INTO SalesHistory 
                            (sales_date, product_code, quantity, unit_price, total_price, created_by, created_at, updated_by, updated_at) 
                            VALUES 
                            (@SalesDate, @ProductCode, @Quantity, @UnitPrice, @TotalPrice, @CreatedBy, @CreatedAt, @UpdatedBy, @UpdatedAt)";

                        conn.Execute(insertSql, list, transaction: transaction);

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }
    }
}