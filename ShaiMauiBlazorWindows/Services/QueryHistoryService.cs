using ShaiMauiExcelToSql.Models;
using System.Text.Json;

namespace ShaiMauiExcelToSql.Services
{
    public class QueryHistoryService
    {
        private const string HistoryKey = "query_history";
        private const string FavoritesKey = "query_favorites";
        private const int MaxHistoryItems = 50;
        
        private List<QueryHistoryItem> _historyItems = new();
        private List<QueryHistoryItem> _favoriteItems = new();
        
        public QueryHistoryService()
        {
            LoadHistory();
            LoadFavorites();
        }
        
        public List<QueryHistoryItem> GetHistory()
        {
            return _historyItems.OrderByDescending(h => h.ExecutedAt).ToList();
        }
        
        public List<QueryHistoryItem> GetFavorites()
        {
            return _favoriteItems.OrderByDescending(f => f.ExecutedAt).ToList();
        }
        
        public void AddToHistory(QueryHistoryItem item)
        {
            // Remove oldest item if we've reached the limit
            if (_historyItems.Count >= MaxHistoryItems)
            {
                var oldest = _historyItems.OrderBy(h => h.ExecutedAt).First();
                _historyItems.Remove(oldest);
            }
            
            _historyItems.Add(item);
            SaveHistory();
        }
        
        public void AddToFavorites(QueryHistoryItem item)
        {
            // Check if this query is already a favorite
            var existingFavorite = _favoriteItems.FirstOrDefault(f => 
                f.ConnectionString == item.ConnectionString && 
                f.SqlQuery == item.SqlQuery);
                
            if (existingFavorite == null)
            {
                item.IsFavorite = true;
                _favoriteItems.Add(item);
                SaveFavorites();
            }
        }
        
        public void RemoveFromFavorites(string id)
        {
            var item = _favoriteItems.FirstOrDefault(f => f.Id == id);
            if (item != null)
            {
                _favoriteItems.Remove(item);
                
                // Also update the item in history if it exists
                var historyItem = _historyItems.FirstOrDefault(h => h.Id == id);
                if (historyItem != null)
                {
                    historyItem.IsFavorite = false;
                }
                
                SaveFavorites();
                SaveHistory();
            }
        }
        
        public void ClearHistory()
        {
            _historyItems.Clear();
            SaveHistory();
        }
        
        public void UpdateFavorite(QueryHistoryItem item)
        {
            var existingItem = _favoriteItems.FirstOrDefault(f => f.Id == item.Id);
            if (existingItem != null)
            {
                existingItem.Name = item.Name;
                SaveFavorites();
            }
        }
        
        private void LoadHistory()
        {
            try
            {
                var json = Preferences.Get(HistoryKey, string.Empty);
                if (!string.IsNullOrEmpty(json))
                {
                    _historyItems = JsonSerializer.Deserialize<List<QueryHistoryItem>>(json) ?? new List<QueryHistoryItem>();
                }
            }
            catch (Exception)
            {
                _historyItems = new List<QueryHistoryItem>();
            }
        }
        
        private void SaveHistory()
        {
            try
            {
                var json = JsonSerializer.Serialize(_historyItems);
                Preferences.Set(HistoryKey, json);
            }
            catch (Exception)
            {
                // Log error or handle exception
            }
        }
        
        private void LoadFavorites()
        {
            try
            {
                var json = Preferences.Get(FavoritesKey, string.Empty);
                if (!string.IsNullOrEmpty(json))
                {
                    _favoriteItems = JsonSerializer.Deserialize<List<QueryHistoryItem>>(json) ?? new List<QueryHistoryItem>();
                }
            }
            catch (Exception)
            {
                _favoriteItems = new List<QueryHistoryItem>();
            }
        }
        
        private void SaveFavorites()
        {
            try
            {
                var json = JsonSerializer.Serialize(_favoriteItems);
                Preferences.Set(FavoritesKey, json);
            }
            catch (Exception)
            {
                // Log error or handle exception
            }
        }
    }
}