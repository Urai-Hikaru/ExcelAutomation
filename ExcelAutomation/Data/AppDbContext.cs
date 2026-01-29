using ExcelAutomation.Models.Entities;
using Microsoft.EntityFrameworkCore;

namespace ExcelAutomation.Data
{
    public class AppDbContext : DbContext
    {
        public DbSet<SalesHistory> SalesHistories { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // 実行ファイルと同じ場所に "ExcelAutomation.db" というファイルを作る設定
            string dbPath = "ExcelAutomation.db";
            optionsBuilder.UseSqlite($"Data Source={dbPath}");
        }
    }
}