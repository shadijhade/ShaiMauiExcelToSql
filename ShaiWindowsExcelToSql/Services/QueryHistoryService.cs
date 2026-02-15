using ShaiWindowsExcelToSql.Models;
using System.Collections.Generic;
using System.Linq;

namespace ShaiWindowsExcelToSql.Services
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
            _historyItems = FileService.Read<List<QueryHistoryItem>>(HistoryKey, new List<QueryHistoryItem>());
        }
        
        private void SaveHistory()
        {
            FileService.Save(HistoryKey, _historyItems);
        }
        
        private void LoadFavorites()
        {
            _favoriteItems = FileService.Read<List<QueryHistoryItem>>(FavoritesKey, new List<QueryHistoryItem>());
        }
        
        private void SaveFavorites()
        {
            FileService.Save(FavoritesKey, _favoriteItems);
        }
    }
}
