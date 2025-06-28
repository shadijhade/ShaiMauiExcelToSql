using ShaiMauiExcelToSql.Models;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Collections.Generic;
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
                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            int resultSetIndex = 0;
                            int totalRows = 0;
                            int totalColumns = 0;
                            
                            do
                            {
                                // Debug: Log current result set info
                                System.Diagnostics.Debug.WriteLine($"Processing result set {resultSetIndex + 1}, FieldCount: {reader.FieldCount}");
                                
                                // Check if the current result set has columns
                                if (reader.FieldCount > 0)
                                {
                                    var dataTable = new DataTable();
                                    
                                    // Manually build the DataTable schema
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        var columnName = reader.GetName(i);
                                        var columnType = reader.GetFieldType(i);
                                        dataTable.Columns.Add(columnName, columnType);
                                    }
                                    
                                    // Read all rows from the current result set
                                    while (await reader.ReadAsync())
                                    {
                                        var row = dataTable.NewRow();
                                        for (int i = 0; i < reader.FieldCount; i++)
                                        {
                                            row[i] = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
                                        }
                                        dataTable.Rows.Add(row);
                                    }
                                    
                                    var resultSet = new ResultSet
                                    {
                                        Data = dataTable,
                                        TotalRows = dataTable.Rows.Count,
                                        TotalColumns = dataTable.Columns.Count,
                                        SheetName = $"ResultSet{resultSetIndex + 1}"
                                    };
                                    
                                    result.ResultSets.Add(resultSet);
                                    totalRows += dataTable.Rows.Count;
                                    totalColumns = Math.Max(totalColumns, dataTable.Columns.Count);
                                    
                                    // Debug: Log result set details
                                    System.Diagnostics.Debug.WriteLine($"Added result set {resultSetIndex + 1}: {dataTable.Rows.Count} rows, {dataTable.Columns.Count} columns");
                                    
                                    resultSetIndex++;
                                }
                                else
                                {
                                    // Handle result sets with no columns (like UPDATE/INSERT statements)
                                    // Skip to next result set
                                    resultSetIndex++;
                                }
                            }
                            while (await reader.NextResultAsync());
                            
                            stopwatch.Stop();
                            
                            result.Success = true;
                            result.TotalRows = totalRows;
                            result.TotalColumns = totalColumns;
                            result.ExecutionTime = stopwatch.Elapsed;
                            
                            // Debug: Log final result
                            System.Diagnostics.Debug.WriteLine($"Total result sets found: {result.ResultSets.Count}");
                            
                            // Set the primary Data property for backward compatibility
                            if (result.ResultSets.Count > 0)
                            {
                                result.Data = result.ResultSets[0].Data;
                            }
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
        
        public async Task<string> ExportMultipleResultSetsToExcelAsync(List<ResultSet> resultSets, ExcelExportOptions options = null, IProgress<int> progress = null)
        {
            // Use the enhanced service for multi-result set export
            return await _enhancedService.ExportMultipleResultSetsToExcelAsync(resultSets, options, progress);
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