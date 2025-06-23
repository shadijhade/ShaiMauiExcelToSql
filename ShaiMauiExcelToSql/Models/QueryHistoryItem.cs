using System;

namespace ShaiMauiExcelToSql.Models
{
    public class QueryHistoryItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string ConnectionString { get; set; } = string.Empty;
        public string SqlQuery { get; set; } = string.Empty;
        public DateTime ExecutedAt { get; set; } = DateTime.Now;
        public bool Success { get; set; }
        public int RowCount { get; set; }
        public TimeSpan ExecutionTime { get; set; }
        public bool IsFavorite { get; set; }
        public string? Name { get; set; }
    }
}