using ShaiMauiExcelToSql.Models;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using ClosedXML.Excel;

namespace ShaiMauiExcelToSql.Services
{
    public class DatabaseService
    {
        public async Task<QueryResult> ExecuteQueryAsync(DatabaseConnection connection)
        {
            var result = new QueryResult();
            var stopwatch = new Stopwatch();
            
            try
            {
                using (var sqlConnection = new SqlConnection(connection.ConnectionString))
                {
                    await sqlConnection.OpenAsync();
                    
                    using (var command = new SqlCommand(connection.SqlQuery, sqlConnection))
                    {
                        command.CommandTimeout = 30; // 30 seconds timeout
                        
                        stopwatch.Start();
                        using (var adapter = new SqlDataAdapter(command))
                        {
                            var dataTable = new DataTable();
                            adapter.Fill(dataTable);
                            stopwatch.Stop();
                            
                            result.Success = true;
                            result.Data = dataTable;
                            result.TotalRows = dataTable.Rows.Count;
                            result.TotalColumns = dataTable.Columns.Count;
                            result.ExecutionTime = stopwatch.Elapsed;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            
            return result;
        }
        
        private readonly EnhancedDatabaseService _enhancedService;
        
        public DatabaseService()
        {
            _enhancedService = new EnhancedDatabaseService();
        }
        
        public async Task<string> ExportToExcelAsync(DataTable dataTable, ExcelExportOptions options = null, IProgress<int> progress = null)
        {
            // Use the enhanced service for export
            return await _enhancedService.ExportToExcelWithCustomOptionsAsync(dataTable, options, progress);
        }
        
        public List<string> GetAvailableTableStyles()
        {
            // Return a list of common Excel table styles
            return new List<string>
            {
                "TableStyleLight1", "TableStyleLight2", "TableStyleLight3",
                "TableStyleLight4", "TableStyleLight5", "TableStyleLight6",
                "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3",
                "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6",
                "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9",
                "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12",
                "TableStyleDark1", "TableStyleDark2", "TableStyleDark3",
                "TableStyleDark4", "TableStyleDark5", "TableStyleDark6"
            };
        }
    }
}