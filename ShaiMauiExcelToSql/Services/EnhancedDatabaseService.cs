using ShaiMauiExcelToSql.Models;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ShaiMauiExcelToSql.Services
{
    public class EnhancedDatabaseService
    {
        public async Task<string> ExportToExcelWithCustomOptionsAsync(DataTable dataTable, ExcelExportOptions? options = null, IProgress<int>? progress = null)
        {
            try
            {
                // Use default options if none provided
                options ??= new ExcelExportOptions();
                
                // Create a file path
                string filePath;
                
                if (!string.IsNullOrEmpty(options.CustomSavePath))
                {
                    // Use the custom save path if provided
                    filePath = options.CustomSavePath;
                    
                    // Ensure the directory exists
                    var directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                }
                // Use the default cache directory
                var fileName = $"{options.FileNamePrefix}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                if (options.CustomSavePath != null && !string.IsNullOrEmpty(options.CustomSavePath) && !options.CustomSavePath.EndsWith(".xlsx"))
                {
                    filePath = Path.Combine(options.CustomSavePath, fileName);
                }
                else
                {
                    filePath = Path.Combine(FileSystem.CacheDirectory, fileName);
                }

                // Calculate total rows to export
                int totalRows = options.ExportAllRows ? dataTable.Rows.Count : Math.Min(options.MaxRowsToExport, dataTable.Rows.Count);
                
                // Create progress tracking if not provided
                progress ??= new Progress<int>();
                
                await Task.Run(() =>
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(options.WorksheetName);
                        
                        // Add headers if requested
                        if (options.IncludeHeaderRow)
                        {
                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                var columnName = dataTable.Columns[i].ColumnName;
                                var columnFormatting = options.GetColumnFormatting(columnName);
                                
                                var cell = worksheet.Cell(1, i + 1);
                                
                                // Use custom header text if provided, otherwise use column name
                                cell.Value = !string.IsNullOrEmpty(columnFormatting.CustomHeaderText) 
                                    ? columnFormatting.CustomHeaderText 
                                    : columnName;
                                
                                // Apply header formatting if requested
                                if (options.BoldHeaders)
                                {
                                    cell.Style.Font.Bold = true;
                                }
                                
                                if (!string.IsNullOrEmpty(options.HeaderBackgroundColor))
                                {
                                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml(options.HeaderBackgroundColor);
                                }
                                
                                if (!string.IsNullOrEmpty(options.HeaderFontColor))
                                {
                                    cell.Style.Font.FontColor = XLColor.FromHtml(options.HeaderFontColor);
                                }
                                
                                // Apply column-specific width if set
                                if (columnFormatting.Width > 0)
                                {
                                    worksheet.Column(i + 1).Width = columnFormatting.Width;
                                }
                                
                                // Hide column if requested
                                if (columnFormatting.Hidden)
                                {
                                    worksheet.Column(i + 1).Hide();
                                }
                            }
                        }
                        
                        // Determine starting row (2 if headers included, 1 if not)
                        int startRow = options.IncludeHeaderRow ? 2 : 1;
                        
                        // Add data with intelligent formatting
                        for (int row = 0; row < totalRows; row++)
                        {
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                var columnName = dataTable.Columns[col].ColumnName;
                                var columnFormatting = options.GetColumnFormatting(columnName);
                                var cell = worksheet.Cell(row + startRow, col + 1);
                                var value = dataTable.Rows[row][col];
                                
                                // Apply intelligent data formatting if requested
                                if (options.ApplyDataFormatting && value != DBNull.Value && value != null)
                                {
                                    // Handle different data types
                                    if (options.FormatDatesAsDate && value is DateTime dateValue)
                                    {
                                        cell.Value = dateValue;
                                        
                                        // Apply custom format if provided, otherwise use default
                                        if (!string.IsNullOrEmpty(columnFormatting.CustomFormat))
                                        {
                                            cell.Style.DateFormat.Format = columnFormatting.CustomFormat;
                                        }
                                        else
                                        {
                                            cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                                        }
                                    }
                                    else if (options.FormatNumbersAsNumbers && 
                                             (value is int || value is long || value is float || 
                                              value is double || value is decimal))
                                    {
                                        cell.Value = XLCellValue.FromObject(value);
                                        
                                        // Apply custom format if provided, otherwise use default
                                        if (!string.IsNullOrEmpty(columnFormatting.CustomFormat))
                                        {
                                            cell.Style.NumberFormat.Format = columnFormatting.CustomFormat;
                                        }
                                        else
                                        {
                                            // Apply number format based on the type
                                            if (value is decimal || value is double || value is float)
                                            {
                                                cell.Style.NumberFormat.Format = "#,##0.00";
                                            }
                                            else
                                            {
                                                cell.Style.NumberFormat.Format = "#,##0";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        cell.Value = value.ToString();
                                    }
                                }
                                else
                                {
                                    // Default to string representation
                                    cell.Value = value?.ToString() ?? string.Empty;
                                }
                                
                                // Apply column-specific formatting
                                if (columnFormatting.Bold)
                                {
                                    cell.Style.Font.Bold = true;
                                }
                                
                                if (columnFormatting.Italic)
                                {
                                    cell.Style.Font.Italic = true;
                                }
                                
                                if (!string.IsNullOrEmpty(columnFormatting.BackgroundColor))
                                {
                                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml(columnFormatting.BackgroundColor);
                                }
                                
                                if (!string.IsNullOrEmpty(columnFormatting.FontColor))
                                {
                                    cell.Style.Font.FontColor = XLColor.FromHtml(columnFormatting.FontColor);
                                }
                            }
                            
                            // Apply alternating row colors if requested (and not overridden by column formatting)
                            if (options.AlternatingRowColors && row % 2 == 1)
                            {
                                var rowStyle = worksheet.Row(row + startRow).Style;
                                
                                // Only apply if not already set by column formatting
                                if (rowStyle.Fill.BackgroundColor == XLColor.NoColor)
                                {
                                    rowStyle.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
                                }
                            }
                            
                            // Report progress every 100 rows or for every 1% of progress
                            int progressInterval = Math.Max(100, totalRows / 100);
                            if (row % progressInterval == 0 || row == totalRows - 1)
                            {
                                int percentComplete = (int)((row / (double)totalRows) * 100);
                                progress.Report(percentComplete);
                            }
                        }
                        
                        // Apply table style if requested
                        if (options.ApplyTableStyle && totalRows > 0)
                        {
                            var dataRange = worksheet.Range(
                                startRow - 1, 1, 
                                totalRows + (startRow - 1), dataTable.Columns.Count);
                                
                            var table = dataRange.CreateTable();
                            
                            // Apply built-in table style if specified
                            if (!string.IsNullOrEmpty(options.TableStyleName))
                            {
                                try
                                {
                                    table.Theme = new XLTableTheme(options.TableStyleName);
                                }
                                catch
                                {
                                    // Fallback to default if style name is invalid
                                    table.Theme = new XLTableTheme("TableStyleMedium9");
                                }
                            }
                        }
                        
                        // Auto-fit columns if requested (and not overridden by column width)
                        if (options.AutoFitColumns)
                        {
                            for (int col = 1; col <= dataTable.Columns.Count; col++)
                            {
                                var columnName = dataTable.Columns[col - 1].ColumnName;
                                var columnFormatting = options.GetColumnFormatting(columnName);
                                
                                // Only auto-fit if width is not explicitly set
                                if (columnFormatting.Width <= 0)
                                {
                                    worksheet.Column(col).AdjustToContents();
                                }
                            }
                        }
                        
                        // Freeze the header row if headers are included
                        if (options.IncludeHeaderRow)
                        {
                            worksheet.SheetView.FreezeRows(1);
                        }
                        
                        // Save the workbook
                        workbook.SaveAs(filePath);
                    }
                });
                
                return filePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export to Excel: {ex.Message}", ex);
            }
        }
        
        public async Task<string> ExportMultipleResultSetsToExcelAsync(List<ResultSet> resultSets, ExcelExportOptions? options = null, IProgress<int>? progress = null)
        {
            try
            {
                // Use default options if none provided
                options ??= new ExcelExportOptions();
                
                // Create a file path
                string filePath;
                
                if (!string.IsNullOrEmpty(options.CustomSavePath))
                {
                    // Use the custom save path if provided
                    filePath = options.CustomSavePath;
                    
                    // Ensure the directory exists
                    var directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                }
                var fileName = $"{options.FileNamePrefix}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                if (options.CustomSavePath != null && !string.IsNullOrEmpty(options.CustomSavePath) && !options.CustomSavePath.EndsWith(".xlsx"))
                {
                    filePath = Path.Combine(options.CustomSavePath, fileName);
                }
                else
                {
                    filePath = Path.Combine(FileSystem.CacheDirectory, fileName);
                }

                // Calculate total rows across all result sets
                int totalRowsAllSets = resultSets.Sum(rs => options.ExportAllRows ? rs.Data.Rows.Count : Math.Min(options.MaxRowsToExport, rs.Data.Rows.Count));
                
                // Create progress tracking if not provided
                progress ??= new Progress<int>();
                
                await Task.Run(() =>
                {
                    using (var workbook = new XLWorkbook())
                    {
                        int processedRows = 0;
                        
                        for (int setIndex = 0; setIndex < resultSets.Count; setIndex++)
                        {
                            var resultSet = resultSets[setIndex];
                            var dataTable = resultSet.Data;
                            
                            // Use the custom sheet name from the result set
                            var worksheet = workbook.Worksheets.Add(resultSet.SheetName);
                            
                            // Calculate rows to export for this result set
                            int totalRows = options.ExportAllRows ? dataTable.Rows.Count : Math.Min(options.MaxRowsToExport, dataTable.Rows.Count);
                            
                            // Add headers if requested
                            if (options.IncludeHeaderRow)
                            {
                                for (int i = 0; i < dataTable.Columns.Count; i++)
                                {
                                    var columnName = dataTable.Columns[i].ColumnName;
                                    var columnFormatting = options.GetColumnFormatting(columnName);
                                    
                                    var cell = worksheet.Cell(1, i + 1);
                                    
                                    // Use custom header text if provided, otherwise use column name
                                    cell.Value = !string.IsNullOrEmpty(columnFormatting.CustomHeaderText) 
                                        ? columnFormatting.CustomHeaderText 
                                        : columnName;
                                    
                                    // Apply header formatting if requested
                                    if (options.BoldHeaders)
                                    {
                                        cell.Style.Font.Bold = true;
                                    }
                                    
                                    if (!string.IsNullOrEmpty(options.HeaderBackgroundColor))
                                    {
                                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml(options.HeaderBackgroundColor);
                                    }
                                    
                                    if (!string.IsNullOrEmpty(options.HeaderFontColor))
                                    {
                                        cell.Style.Font.FontColor = XLColor.FromHtml(options.HeaderFontColor);
                                    }
                                    
                                    // Apply column-specific width if set
                                    if (columnFormatting.Width > 0)
                                    {
                                        worksheet.Column(i + 1).Width = columnFormatting.Width;
                                    }
                                    
                                    // Hide column if requested
                                    if (columnFormatting.Hidden)
                                    {
                                        worksheet.Column(i + 1).Hide();
                                    }
                                }
                            }
                            
                            // Determine starting row (2 if headers included, 1 if not)
                            int startRow = options.IncludeHeaderRow ? 2 : 1;
                            
                            // Add data with intelligent formatting
                            for (int row = 0; row < totalRows; row++)
                            {
                                for (int col = 0; col < dataTable.Columns.Count; col++)
                                {
                                    var columnName = dataTable.Columns[col].ColumnName;
                                    var columnFormatting = options.GetColumnFormatting(columnName);
                                    var cell = worksheet.Cell(row + startRow, col + 1);
                                    var value = dataTable.Rows[row][col];
                                    
                                    // Apply intelligent data formatting if requested
                                    if (options.ApplyDataFormatting && value != DBNull.Value && value != null)
                                    {
                                        // Handle different data types
                                        if (options.FormatDatesAsDate && value is DateTime dateValue)
                                        {
                                            cell.Value = dateValue;
                                            
                                            // Apply custom format if provided, otherwise use default
                                            if (!string.IsNullOrEmpty(columnFormatting.CustomFormat))
                                            {
                                                cell.Style.DateFormat.Format = columnFormatting.CustomFormat;
                                            }
                                            else
                                            {
                                                cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                                            }
                                        }
                                        else if (options.FormatNumbersAsNumbers && 
                                                 (value is int || value is long || value is float || 
                                                  value is double || value is decimal))
                                        {
                                            cell.Value = XLCellValue.FromObject(value);
                                            
                                            // Apply custom format if provided, otherwise use default
                                            if (!string.IsNullOrEmpty(columnFormatting.CustomFormat))
                                            {
                                                cell.Style.NumberFormat.Format = columnFormatting.CustomFormat;
                                            }
                                            else
                                            {
                                                // Apply number format based on the type
                                                if (value is decimal || value is double || value is float)
                                                {
                                                    cell.Style.NumberFormat.Format = "#,##0.00";
                                                }
                                                else
                                                {
                                                    cell.Style.NumberFormat.Format = "#,##0";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            cell.Value = value.ToString();
                                        }
                                    }
                                    else
                                    {
                                        // Default to string representation
                                        cell.Value = value?.ToString() ?? string.Empty;
                                    }
                                    
                                    // Apply column-specific formatting
                                    if (columnFormatting.Bold)
                                    {
                                        cell.Style.Font.Bold = true;
                                    }
                                    
                                    if (columnFormatting.Italic)
                                    {
                                        cell.Style.Font.Italic = true;
                                    }
                                    
                                    if (!string.IsNullOrEmpty(columnFormatting.BackgroundColor))
                                    {
                                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml(columnFormatting.BackgroundColor);
                                    }
                                    
                                    if (!string.IsNullOrEmpty(columnFormatting.FontColor))
                                    {
                                        cell.Style.Font.FontColor = XLColor.FromHtml(columnFormatting.FontColor);
                                    }
                                }
                                
                                // Apply alternating row colors if requested (and not overridden by column formatting)
                                if (options.AlternatingRowColors && row % 2 == 1)
                                {
                                    var rowStyle = worksheet.Row(row + startRow).Style;
                                    
                                    // Only apply if not already set by column formatting
                                    if (rowStyle.Fill.BackgroundColor == XLColor.NoColor)
                                    {
                                        rowStyle.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
                                    }
                                }
                                
                                // Report progress every 100 rows or for every 1% of progress
                                processedRows++;
                                int progressInterval = Math.Max(100, totalRowsAllSets / 100);
                                if (processedRows % progressInterval == 0 || processedRows == totalRowsAllSets)
                                {
                                    int percentComplete = (int)((processedRows / (double)totalRowsAllSets) * 100);
                                    progress.Report(percentComplete);
                                }
                            }
                            
                            // Apply table style if requested
                            if (options.ApplyTableStyle && totalRows > 0)
                            {
                                var dataRange = worksheet.Range(
                                    startRow - 1, 1, 
                                    totalRows + (startRow - 1), dataTable.Columns.Count);
                                    
                                var table = dataRange.CreateTable();
                                
                                // Apply built-in table style if specified
                                if (!string.IsNullOrEmpty(options.TableStyleName))
                                {
                                    try
                                    {
                                        table.Theme = new XLTableTheme(options.TableStyleName);
                                    }
                                    catch
                                    {
                                        // Fallback to default if style name is invalid
                                        table.Theme = new XLTableTheme("TableStyleMedium9");
                                    }
                                }
                            }
                            
                            // Auto-fit columns if requested (and not overridden by column width)
                            if (options.AutoFitColumns)
                            {
                                for (int col = 1; col <= dataTable.Columns.Count; col++)
                                {
                                    var columnName = dataTable.Columns[col - 1].ColumnName;
                                    var columnFormatting = options.GetColumnFormatting(columnName);
                                    
                                    // Only auto-fit if width is not explicitly set
                                    if (columnFormatting.Width <= 0)
                                    {
                                        worksheet.Column(col).AdjustToContents();
                                    }
                                }
                            }
                            
                            // Freeze the header row if headers are included
                            if (options.IncludeHeaderRow)
                            {
                                worksheet.SheetView.FreezeRows(1);
                            }
                        }
                        
                        // Save the workbook
                        workbook.SaveAs(filePath);
                    }
                });
                
                return filePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to export multiple result sets to Excel: {ex.Message}", ex);
            }
        }
    }
}