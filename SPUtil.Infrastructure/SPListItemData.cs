namespace SPUtil.Infrastructure
{
    public class SPListItemData
    {
        public int Id { get; set; }
        public string Title { get; set; } = string.Empty;
        // Поле для галочки (выбрать для копирования)
        public bool IsSelected { get; set; }
		public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();	
		// Absolute URL to the item's standard display form — used to open
        // the item in a browser directly from the ID column.
        public string DispFormUrl { get; set; } = string.Empty;		
    }
}