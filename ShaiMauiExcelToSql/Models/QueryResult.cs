using System.Data;
using System.Collections.Generic;

namespace ShaiMauiExcelToSql.Models
{
    public class ResultSet
    {
        public DataTable Data { get; set; } = new DataTable();
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public string SheetName { get; set; } = "Sheet";
    }

    public class QueryResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public DataTable? Data { get; set; } // Keep for backward compatibility
        public List<ResultSet> ResultSets { get; set; } = new List<ResultSet>();
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public TimeSpan ExecutionTime { get; set; }
        public int ResultSetCount => ResultSets.Count;
        
        // Helper property to check if this is a multi-result set query
        public bool HasMultipleResultSets => ResultSets.Count > 1;
        
        // Helper method to get the primary result set (for backward compatibility)
        public DataTable? GetPrimaryResultSet()
        {
            if (ResultSets.Count > 0)
                return ResultSets[0].Data;
            return Data;
        }
    }
}