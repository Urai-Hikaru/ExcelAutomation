using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelAutomation.Models.Entities
{
    [Table("sales_history")]
    [PrimaryKey(nameof(SalesDate), nameof(ProductCode))]
    public class SalesHistory
    {
        [Column("sales_date")]
        public DateTime SalesDate { get; set; }

        [Column("product_code")]
        public string ProductCode { get; set; } = string.Empty;

        [Column("quantity")]
        public int Quantity { get; set; }

        [Column("unit_price")]
        public int UnitPrice { get; set; }

        [Column("total_price")]
        public int TotalPrice { get; set; }

        [Column("created_by")]
        public string CreatedBy { get; set; } = string.Empty;

        [Column("created_at")]
        public DateTime CreatedAt { get; set; }

        [Column("updated_by")]
        public string UpdatedBy { get; set; } = string.Empty;

        [Column("updated_at")]
        public DateTime UpdatedAt { get; set; }
    }
    public class SalesHistoryGridModel
    {
        public DateTime SalesDate { get; set; }

        public string ProductCode { get; set; } = string.Empty;

        public int Quantity { get; set; }

        public int UnitPrice { get; set; }

        public int TotalPrice { get; set; }

    }
    public class MonthlySummaryModel
    {
        public string ProductCode { get; set; } = string.Empty;
        public int TotalQuantity { get; set; }
        public int TotalPrice { get; set; }
    }

}