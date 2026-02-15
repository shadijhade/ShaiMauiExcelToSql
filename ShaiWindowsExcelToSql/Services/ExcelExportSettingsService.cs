using ShaiWindowsExcelToSql.Models;

namespace ShaiWindowsExcelToSql.Services
{
    public class ExcelExportSettingsService
    {
        private const string SettingsKey = "excel_export_settings";
        private ExcelExportOptions _currentSettings = new ExcelExportOptions();
        
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
            FileService.Save(SettingsKey, settings);
        }
        
        public void ResetToDefaults()
        {
            _currentSettings = new ExcelExportOptions();
            SaveSettings(_currentSettings);
        }
        
        private void LoadSettings()
        {
            _currentSettings = FileService.Read<ExcelExportOptions>(SettingsKey, new ExcelExportOptions());
        }
    }
}
