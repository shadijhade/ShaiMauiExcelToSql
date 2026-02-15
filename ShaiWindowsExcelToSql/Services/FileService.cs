using System;
using System.IO;
using System.Text.Json;

namespace ShaiWindowsExcelToSql.Services
{
    public static class FileService
    {
        private static readonly string AppDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            "ShaiWindowsExcelToSql");

        static FileService()
        {
            if (!Directory.Exists(AppDataPath))
            {
                Directory.CreateDirectory(AppDataPath);
            }
        }

        public static T Read<T>(string key, T defaultValue)
        {
            try
            {
                string filePath = Path.Combine(AppDataPath, $"{key}.json");
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    return JsonSerializer.Deserialize<T>(json) ?? defaultValue;
                }
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        public static void Save<T>(string key, T data)
        {
            try
            {
                string filePath = Path.Combine(AppDataPath, $"{key}.json");
                string json = JsonSerializer.Serialize(data);
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving data: {ex.Message}");
            }
        }
    }
}
