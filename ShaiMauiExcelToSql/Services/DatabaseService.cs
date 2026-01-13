using ShaiMauiExcelToSql.Models;
using System.Data;
using Microsoft.Data.SqlClient;
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
                if (string.IsNullOrWhiteSpace(connection.ConnectionString))
                    throw new ArgumentException("Connection string is required.");
                
                if (string.IsNullOrWhiteSpace(connection.SqlQuery))
                    throw new ArgumentException("SQL Query is required.");

                var builder = new SqlConnectionStringBuilder(connection.ConnectionString);
                builder.TrustServerCertificate = true; // Ensure we trust the server certificate

                await using (var sqlConnection = await OpenConnectionWithRetryAsync(builder))
                {
                    if (sqlConnection.State != ConnectionState.Open)
                        throw new InvalidOperationException($"Connection failed to open. Current state: {sqlConnection.State}");

                    await using (var command = new SqlCommand(connection.SqlQuery, sqlConnection))
                    {
                        command.CommandTimeout = 30; // 30 seconds timeout
                        
                        stopwatch.Start();
                        await using (var reader = await command.ExecuteReaderAsync())
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
                                        
                                        // Handle duplicate column names
                                        var originalColumnName = columnName;
                                        int duplicateCount = 1;
                                        while (dataTable.Columns.Contains(columnName))
                                        {
                                            columnName = $"{originalColumnName}_{duplicateCount}";
                                            duplicateCount++;
                                        }

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
                                    // Make sure we still increment index to track separate results
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
            catch (InvalidOperationException ex)
            {
                var builder = new SqlConnectionStringBuilder(connection.ConnectionString);
                result.Success = false;
                result.ErrorMessage = $"Invalid Operation: {ex.Message}. Connection state might be invalid. \nAttempted to connect to: '{builder.DataSource}', Database: '{builder.InitialCatalog}'";
                System.Diagnostics.Debug.WriteLine(result.ErrorMessage);
            }
            catch (SqlException ex)
            {
                 var builder = new SqlConnectionStringBuilder(connection.ConnectionString);
                 result.Success = false;
                 result.ErrorMessage = $"SQL Error: {ex.Message} \nAttempted to connect to: '{builder.DataSource}', Database: '{builder.InitialCatalog}'";
            }
            catch (Exception ex)
            {
                var builder = new SqlConnectionStringBuilder(connection.ConnectionString);
                result.Success = false;
                result.ErrorMessage = $"Error: {ex.GetType().Name} - {ex.Message} \nAttempted to connect to: '{builder.DataSource}', Database: '{builder.InitialCatalog}'";
            }
            
            return result;
        }

        private async Task<SqlConnection> OpenConnectionWithRetryAsync(SqlConnectionStringBuilder builder)
        {
            // Internal helper to create the connection
            // Capture the Original Exception to throw it if retry fails, so we don't confuse the user
            Exception originalException = null;

            // Attempt 1: Standard Connection
            try
            {
                var connection = new SqlConnection(builder.ConnectionString);
                await connection.OpenAsync();
                return connection;
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("Instance failure"))
            {
                originalException = ex;
                System.Diagnostics.Debug.WriteLine($"Initial connection failed with 'Instance failure'. Attempting retry with TCP/IP...");
            }
            catch (PlatformNotSupportedException ex)
            {
                 // Sometimes SNI native might throw this on some platforms/configs
                 originalException = ex;
            }
            catch (Exception ex)
            {
                // For other exceptions (like Auth failure), don't verify retry for now
                throw; 
            }

            // Check if we should retry (Local instance?)
            if (IsLocalDataSource(builder.DataSource))
            {
                 try 
                 {
                    // Construct new DataSource: tcp:localhost\InstanceName
                    string instanceName = "";
                     if (builder.DataSource.Contains("\\"))
                     {
                         // Handle cases like ".\SQLEXPRESS" or ".\\SQLEXPRESS" which might occur in some connection strings
                         var parts = builder.DataSource.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                         if (parts.Length > 1) 
                         {
                             // Usually the last part is the instance name if split correctly
                             instanceName = "\\" + parts.Last(); 
                         }
                     }
                     
                     var originalDataSource = builder.DataSource;
                     builder.DataSource = $"tcp:localhost{instanceName}";
                     System.Diagnostics.Debug.WriteLine($"Retrying connection with DataSource: {builder.DataSource}");
                     
                     var retryConnection = new SqlConnection(builder.ConnectionString);
                     await retryConnection.OpenAsync();
                     return retryConnection;
                 }
                 catch (Exception retryEx)
                 {
                     System.Diagnostics.Debug.WriteLine($"Retry failed: {retryEx.Message}");
                     // If retry fails, throw the ORIGINAL exception so the user sees the primary error "Instance failure" 
                     // OR throw a new one explaining both. 
                     // Let's stick to the structure we have in ExecuteQueryAsync which catches InvalidOperationException
                     if (originalException != null) throw originalException;
                     throw;
                 }
            }
            
            // If not local or no exception captured but fell through (shouldn't happen)
            if (originalException != null) throw originalException;
            throw new InvalidOperationException("Unknown connection error.");
        }

        private bool IsLocalDataSource(string dataSource)
        {
            if (string.IsNullOrEmpty(dataSource)) return false;
            dataSource = dataSource.ToLowerInvariant();
            return dataSource.StartsWith(".") || 
                   dataSource.StartsWith("(local)") || 
                   dataSource.StartsWith("localhost") ||
                   dataSource.Equals("127.0.0.1");
        }
        
        private readonly EnhancedDatabaseService _enhancedService;
        
        public DatabaseService()
        {
            _enhancedService = new EnhancedDatabaseService();
        }
        
        public async Task<string> ExportToExcelAsync(DataTable dataTable, ExcelExportOptions? options = null, IProgress<int>? progress = null)
        {
            // Use the enhanced service for export
            return await _enhancedService.ExportToExcelWithCustomOptionsAsync(dataTable, options, progress);
        }
        
        public async Task<string> ExportMultipleResultSetsToExcelAsync(List<ResultSet> resultSets, ExcelExportOptions? options = null, IProgress<int>? progress = null)
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