namespace SPUtil.Infrastructure
{
    public class SPViewData
    {
        public string Title { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string ViewQuery { get; set; } = string.Empty;
        public string[]? ViewFields { get; set; }
        public bool DefaultView { get; set; }
        public string ServerRelativeUrl { get; set; } = string.Empty;
        public string SchemaXml { get; set; } = string.Empty;
        
        // Добавляем это свойство для счетчиков и сумм
        public string Aggregations { get; set; } = string.Empty; 
    }
}