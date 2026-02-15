using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ShaiWindowsExcelToSql.Models;
using ShaiWindowsExcelToSql.Services;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading.Tasks;
using System.Windows;

namespace ShaiWindowsExcelToSql.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly DatabaseService _databaseService;
        private readonly QueryHistoryService _queryHistoryService;
        private readonly ExcelExportSettingsService _exportSettingsService;

        [ObservableProperty]
        private string _connectionString = string.Empty;

        [ObservableProperty]
        private string _sqlQuery = string.Empty;

        [ObservableProperty]
        private DataTable? _queryResults;

        [ObservableProperty]
        private bool _isBusy;

        [ObservableProperty]
        private string _statusMessage = "Ready";

        [ObservableProperty]
        private ObservableCollection<ResultSet> _resultSets = new();

        [ObservableProperty]
        private ResultSet? _selectedResultSet;

        public MainViewModel()
        {
            _databaseService = new DatabaseService();
            _queryHistoryService = new QueryHistoryService();
            _exportSettingsService = new ExcelExportSettingsService();

            // Load last used connection string if available (could be added to settings service)
        }

        [RelayCommand]
        private async Task ExecuteQueryAsync()
        {
            if (string.IsNullOrWhiteSpace(ConnectionString) || string.IsNullOrWhiteSpace(SqlQuery))
            {
                StatusMessage = "Please enter both Connection String and SQL Query.";
                return;
            }

            IsBusy = true;
            StatusMessage = "Executing query...";
            QueryResults = null;
            ResultSets.Clear();

            try
            {
                var connection = new DatabaseConnection
                {
                    ConnectionString = ConnectionString,
                    SqlQuery = SqlQuery
                };

                var result = await _databaseService.ExecuteQueryAsync(connection);

                if (result.Success)
                {
                    StatusMessage = $"Query executed successfully in {result.ExecutionTime.TotalSeconds:F2}s. {result.TotalRows} rows affected.";
                    
                    if (result.ResultSets.Count > 0)
                    {
                        foreach (var rs in result.ResultSets)
                        {
                            ResultSets.Add(rs);
                        }
                        SelectedResultSet = ResultSets[0];
                        QueryResults = ResultSets[0].Data;
                        HasMultipleResultSets = ResultSets.Count > 1;
                    }
                    else
                    {
                        HasMultipleResultSets = false;
                    }

                    // Add to history
                    _queryHistoryService.AddToHistory(new QueryHistoryItem
                    {
                        ConnectionString = ConnectionString,
                        SqlQuery = SqlQuery,
                        Success = true,
                        RowCount = result.TotalRows,
                        ExecutionTime = result.ExecutionTime
                    });
                }
                else
                {
                    StatusMessage = $"Error: {result.ErrorMessage}";
                    MessageBox.Show(result.ErrorMessage, "Query Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    HasMultipleResultSets = false;
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Unexpected Error: {ex.Message}";
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                HasMultipleResultSets = false;
            }
            finally
            {
                IsBusy = false;
            }
        }

        [ObservableProperty]
        private bool _hasMultipleResultSets;

        [RelayCommand]
        private async Task ExportToExcelAsync()
        {
            if (ResultSets.Count == 0)
            {
                StatusMessage = "No data to export.";
                return;
            }

            IsBusy = true;
            StatusMessage = "Exporting to Excel...";

            try
            {
                var options = _exportSettingsService.GetSettings();
                
                // Prompt user for save location? 
                // For now, let's use the default logic in EnhancedDatabaseService which uses Temp or CustomSavePath
                // In a real WPF app, likely want a SaveFileDialog here to set CustomSavePath override
                
                var dialog = new Microsoft.Win32.SaveFileDialog
                {
                    FileName = $"{options.FileNamePrefix}_{DateTime.Now:yyyyMMddHHmmss}",
                    DefaultExt = ".xlsx",
                    Filter = "Excel Worksheets|*.xlsx"
                };

                if (dialog.ShowDialog() == true)
                {
                    options.CustomSavePath = dialog.FileName;
                    
                    string filePath;
                    if (ResultSets.Count > 1)
                    {
                        filePath = await _databaseService.ExportMultipleResultSetsToExcelAsync([.. ResultSets], options);
                    }
                    else
                    {
                         filePath = await _databaseService.ExportToExcelAsync(ResultSets[0].Data, options);
                    }

                    StatusMessage = $"Export complete: {filePath}";
                    
                    if (options.OpenAfterExport)
                    {
                        try
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = filePath,
                                UseShellExecute = true
                            };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch { /* Ignore open errors */ }
                    }
                }
                else
                {
                    StatusMessage = "Export cancelled.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export Error: {ex.Message}";
                MessageBox.Show(ex.Message, "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }
        
        partial void OnSelectedResultSetChanged(ResultSet? value)
        {
            if (value != null)
            {
                QueryResults = value.Data;
            }
        }
    }
}
