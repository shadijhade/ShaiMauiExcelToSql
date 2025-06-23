using System.Collections.Generic;
using System.Drawing;

namespace ShaiMauiExcelToSql.Models
{
    public class ColumnFormatting
    {
        public string ColumnName { get; set; } = string.Empty;
        public string CustomHeaderText { get; set; } = string.Empty;
        public string CustomFormat { get; set; } = string.Empty;
        public bool Bold { get; set; } = false;
        public bool Italic { get; set; } = false;
        public string BackgroundColor { get; set; } = string.Empty;
        public string FontColor { get; set; } = string.Empty;
        public int Width { get; set; } = 0; // 0 means auto-width
        public bool Hidden { get; set; } = false;
    }

    public class ExcelExportOptions
    {
        // Basic options
        public string WorksheetName { get; set; } = "Query Result";
        public bool IncludeHeaderRow { get; set; } = true;
        public bool AutoFitColumns { get; set; } = true;
        
        // Formatting options
        public bool ApplyTableStyle { get; set; } = true;
        public string TableStyleName { get; set; } = "TableStyleMedium9"; // Excel built-in table style
        public bool AlternatingRowColors { get; set; } = true;
        
        // Header formatting
        public bool BoldHeaders { get; set; } = true;
        public string HeaderBackgroundColor { get; set; } = "#4472C4"; // Excel default blue
        public string HeaderFontColor { get; set; } = "#FFFFFF"; // White
        
        // Data formatting
        public bool ApplyDataFormatting { get; set; } = true;
        public bool FormatDatesAsDate { get; set; } = true;
        public bool FormatNumbersAsNumbers { get; set; } = true;
        
        // Export scope
        public bool ExportAllRows { get; set; } = true;
        public int MaxRowsToExport { get; set; } = 1000;
        
        // File options
        public string FileNamePrefix { get; set; } = "QueryResult";
        public bool OpenAfterExport { get; set; } = true;
        public string CustomSavePath { get; set; } = string.Empty;
        
        // Column-specific formatting
        public List<ColumnFormatting> ColumnFormattings { get; set; } = new List<ColumnFormatting>();
        
        // Helper method to get column formatting for a specific column
        public ColumnFormatting GetColumnFormatting(string columnName)
        {
            return ColumnFormattings.FirstOrDefault(cf => cf.ColumnName == columnName) ?? new ColumnFormatting { ColumnName = columnName };
        }
        
        // Helper method to add or update column formatting
        public void SetColumnFormatting(ColumnFormatting formatting)
        {
            var existing = ColumnFormattings.FirstOrDefault(cf => cf.ColumnName == formatting.ColumnName);
            if (existing != null)
            {
                // Update existing
                var index = ColumnFormattings.IndexOf(existing);
                ColumnFormattings[index] = formatting;
            }
            else
            {
                // Add new
                ColumnFormattings.Add(formatting);
            }
        }
    }
}