using System.ComponentModel.DataAnnotations;

namespace ShaiMauiExcelToSql.Models
{
    public class DatabaseConnection
    {
        [Required(ErrorMessage = "Connection string is required")]
        public string ConnectionString { get; set; } = string.Empty;
        
        [Required(ErrorMessage = "SQL query is required")]
        public string SqlQuery { get; set; } = string.Empty;
    }
}