using Prism.Mvvm;

namespace SPUtil.Infrastructure
{
    public class SPViewData : BindableBase
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

        // 2026-06-10: IsSelected — user checkbox for selecting views to copy.
        // All checks happen at copy time (CopySelectedViewsAsync), not on checkbox click.
        private bool _isSelected;

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }
}