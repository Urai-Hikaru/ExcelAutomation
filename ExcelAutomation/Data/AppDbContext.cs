using ExcelAutomation.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelAutomation.Data
{
    public class AppDbContext : DbContext
    {
        // 作成したテーブル定義をここに登録する
        public DbSet<SalesHistory> SalesHistories { get; set; }

        // データベースファイルの設定
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // 実行ファイルと同じ場所に "ExcelAutomation.db" というファイルを作る設定
            string dbPath = "ExcelAutomation.db";
            optionsBuilder.UseSqlite($"Data Source={dbPath}");
        }
    }
}