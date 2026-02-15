using System.Data;

namespace ShaiWindowsExcelToSql.Models
{
    public class ResultSet
    {
        public DataTable Data { get; set; } = new DataTable();
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public string SheetName { get; set; } = "Sheet";
    }
}
