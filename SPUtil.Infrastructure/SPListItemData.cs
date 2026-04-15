namespace SPUtil.Infrastructure
{
    public class SPListItemData
    {
        public int Id { get; set; }
        public string Title { get; set; } = string.Empty;
        // Поле для галочки (выбрать для копирования)
        public bool IsSelected { get; set; }
		public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();		
    }
}