using System.Data;

namespace ShaiMauiExcelToSql.Models
{
    public class QueryResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public DataTable? Data { get; set; }
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public TimeSpan ExecutionTime { get; set; }
    }
}