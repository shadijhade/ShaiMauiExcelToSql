using ShaiMauiExcelToSql.Models;
using System.Text.Json;

namespace ShaiMauiExcelToSql.Services
{
    public class ExcelExportSettingsService
    {
        private const string SettingsKey = "excel_export_settings";
        private ExcelExportOptions _currentSettings;
        
        public ExcelExportSettingsService()
        {
            LoadSettings();
        }
        
        public ExcelExportOptions GetSettings()
        {
            return _currentSettings;
        }
        
        public void SaveSettings(ExcelExportOptions settings)
        {
            _currentSettings = settings;
            
            try
            {
                var json = JsonSerializer.Serialize(settings);
                Preferences.Set(SettingsKey, json);
            }
            catch (Exception)
            {
                // Log error or handle exception
            }
        }
        
        public void ResetToDefaults()
        {
            _currentSettings = new ExcelExportOptions();
            SaveSettings(_currentSettings);
        }
        
        private void LoadSettings()
        {
            try
            {
                var json = Preferences.Get(SettingsKey, string.Empty);
                if (!string.IsNullOrEmpty(json))
                {
                    _currentSettings = JsonSerializer.Deserialize<ExcelExportOptions>(json) ?? new ExcelExportOptions();
                }
                else
                {
                    _currentSettings = new ExcelExportOptions();
                }
            }
            catch (Exception)
            {
                _currentSettings = new ExcelExportOptions();
            }
        }
    }
}